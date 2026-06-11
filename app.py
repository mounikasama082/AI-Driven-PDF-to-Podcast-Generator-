"""
╔══════════════════════════════════════════════════════════════════╗
║  AI PODCAST GENERATOR · Flask Backend (OPTIMIZED)               ║
║  Run: python app.py → http://localhost:5000                      ║
║                                                                  ║
║  Optimizations:                                                  ║
║  ✓ Content-hash deduplication (skip regeneration)               ║
║  ✓ Per-user daily character quota (free tier)                    ║
║  ✓ Per-minute + per-day rate limiting (anti-spam)               ║
║  ✓ Script caching (only TTS runs once per unique file)          ║
║  ✓ Preview uses pyttsx3 (free local TTS, zero EL chars)         ║
║  ✓ ElevenLabs only for final audio output                       ║
║  ✓ Exact char usage tracked and stored per user                 ║
║                                                                  ║
║  Fixes applied:                                                  ║
║  ✓ In-memory _get_pods now sorts by created_at descending       ║
║  ✓ _mem_pods mutations protected by threading.Lock              ║
║  ✓ Linux espeak note added to startup output                    ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os, io, re, struct, base64, logging, textwrap, time, uuid, hashlib, threading
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor
from functools import wraps
from collections import defaultdict

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv
import pypdf
from docx import Document as DocxDocument
from pptx import Presentation
from groq import Groq
from elevenlabs import ElevenLabs
import jwt
import bcrypt

load_dotenv()

# ── OCR (optional) ────────────────────────────────────────────────
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ── Local TTS for preview (optional, free) ────────────────────────
try:
    import pyttsx3
    LOCAL_TTS_AVAILABLE = True
except ImportError:
    LOCAL_TTS_AVAILABLE = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger("PodcastAI")

# ════════════════════════════════════════════════════════════════
# QUOTA & RATE LIMIT CONFIG  (override via .env)
# ════════════════════════════════════════════════════════════════
FREE_DAILY_CHARS = int(os.environ.get("FREE_DAILY_CHARS", 50_000))  # ElevenLabs chars/day/user
RATE_LIMIT_RPM   = int(os.environ.get("RATE_LIMIT_RPM",  5))        # max generate requests/minute/user
RATE_LIMIT_RPD   = int(os.environ.get("RATE_LIMIT_RPD",  20))       # max generate requests/day/user

# In-memory rate-limit store  { uid: {"minute": [ts, ...], "day": [ts, ...]} }
_rate_store: dict = defaultdict(lambda: {"minute": [], "day": []})
_rate_store_lock = threading.Lock()

# In-memory usage store  { uid: {"date": "YYYY-MM-DD", "chars": int} }
_usage_store: dict = {}
_usage_store_lock = threading.Lock()

# ── MongoDB ────────────────────────────────────────────────────────
MONGO_URI = os.environ.get("MONGO_URI", "")
db = None
users_col = None
podcasts_col = None
cache_col = None   # script cache keyed by content hash
usage_col = None   # per-user daily ElevenLabs char usage

if MONGO_URI:
    try:
        from pymongo import MongoClient
        from pymongo.server_api import ServerApi
        _mc = MongoClient(MONGO_URI, server_api=ServerApi("1"), serverSelectionTimeoutMS=5000)
        _mc.admin.command("ping")
        db           = _mc["podcastai"]
        users_col    = db["users"]
        podcasts_col = db["podcasts"]
        cache_col    = db["script_cache"]
        usage_col    = db["usage"]
        users_col.create_index("email", unique=True)
        podcasts_col.create_index("user_id")
        podcasts_col.create_index("share_id",    unique=True, sparse=True)
        podcasts_col.create_index("content_hash")
        cache_col.create_index("content_hash",   unique=True)
        usage_col.create_index([("user_id", 1), ("date", 1)], unique=True)
        print("MongoDB connected.")
    except Exception as _e:
        print(f"MongoDB unavailable: {_e} — using in-memory.")
        db = None
else:
    print("MONGO_URI not set — using in-memory storage.")

_mem_users: dict = {}
_mem_pods:  list = []
_mem_cache: dict = {}   # content_hash → {script, created_at}

# FIX: Lock to protect _mem_pods from concurrent read/write in threaded Flask
_mem_pods_lock = threading.Lock()

# ── Credentials ───────────────────────────────────────────────────
def _env(key, default="", required=True):
    v = os.environ.get(key, default)
    if required and not v:
        log.error("Missing env variable: %s — add it to .env", key)
        return ""
    return v

GROQ_API_KEY       = _env("GROQ_API_KEY")
ELEVENLABS_API_KEY = _env("ELEVENLABS_API_KEY")
JWT_SECRET         = _env("JWT_SECRET",       default="dev-jwt-secret-change-in-prod-32c", required=False)
FLASK_SECRET_KEY   = _env("FLASK_SECRET_KEY", default="dev-flask-secret-change-in-prod",   required=False)

groq_client = Groq(api_key=GROQ_API_KEY)            if GROQ_API_KEY       else None
el_client   = ElevenLabs(api_key=ELEVENLABS_API_KEY) if ELEVENLABS_API_KEY else None
GROQ_MODEL  = "llama-3.3-70b-versatile"

# ── Fallback voices ───────────────────────────────────────────────
FALLBACK_VOICES = [
    {"voice_id": "nPczCjzI2devNBz1zQrb", "name": "Brian",    "gender": "male",   "category": "general", "preview_url": ""},
    {"voice_id": "cgSgspJ2msm6clMCkdW9", "name": "Jessica",  "gender": "female", "category": "general", "preview_url": ""},
    {"voice_id": "EXAVITQu4vr4xnSDxMaL", "name": "Sarah",    "gender": "female", "category": "general", "preview_url": ""},
    {"voice_id": "TX3LPaxmHKxFdv7VOQHJ", "name": "Liam",     "gender": "male",   "category": "general", "preview_url": ""},
    {"voice_id": "XB0fDUnXU5powFXDhCwa", "name": "Charlotte","gender": "female", "category": "general", "preview_url": ""},
    {"voice_id": "iP95p4xoKVk53GoZ742B", "name": "Chris",    "gender": "male",   "category": "general", "preview_url": ""},
    {"voice_id": "onwK4e9ZLuTAKqWW03F9", "name": "Daniel",   "gender": "male",   "category": "general", "preview_url": ""},
    {"voice_id": "9BWtsMINqrJLrRacOk9x", "name": "Aria",     "gender": "female", "category": "general", "preview_url": ""},
    {"voice_id": "CwhRBWXzGAHq8TQ4Fs17", "name": "Roger",    "gender": "male",   "category": "general", "preview_url": ""},
    {"voice_id": "FGY2WhTYpPnrIDTdsKH5", "name": "Laura",    "gender": "female", "category": "general", "preview_url": ""},
]

_voice_cache: list = []
_voice_ts:    float = 0
VOICE_TTL = 3600

def _fetch_voices():
    if not el_client:
        return list(FALLBACK_VOICES)
    try:
        resp   = el_client.voices.get_all()
        raw    = resp.voices if hasattr(resp, "voices") else []
        result = []
        for v in raw:
            vid = getattr(v, "voice_id", None)
            if not vid:
                continue
            name     = getattr(v, "name", "Unknown")
            labels   = getattr(v, "labels", {}) or {}
            gender   = (labels.get("gender")   or "").lower()
            use_case = (labels.get("use_case") or "").lower()
            preview  = getattr(v, "preview_url", None) or ""
            result.append({
                "voice_id":    vid,
                "name":        name,
                "gender":      gender if gender in ("male", "female") else "other",
                "category":    use_case or "general",
                "preview_url": preview,
            })
        result.sort(key=lambda x: x["name"].lower())
        return result[:20] if result else list(FALLBACK_VOICES)
    except Exception as exc:
        log.warning("Voice fetch failed: %s — using fallbacks.", exc)
        return list(FALLBACK_VOICES)

def get_voices():
    global _voice_cache, _voice_ts
    if not _voice_cache or (time.time() - _voice_ts) > VOICE_TTL:
        _voice_cache = _fetch_voices()
        _voice_ts    = time.time()
    return _voice_cache

def _default_voices():
    v  = get_voices()
    m  = [x for x in v if x["gender"] == "male"]
    f  = [x for x in v if x["gender"] == "female"]
    v1 = m[0]["voice_id"] if m else FALLBACK_VOICES[0]["voice_id"]
    v2 = f[0]["voice_id"] if f else FALLBACK_VOICES[1]["voice_id"]
    return v1, v2

SUPPORTED_LANGUAGES = {
    "en": "English", "hi": "Hindi",    "te": "Telugu",     "es": "Spanish",
    "fr": "French",  "de": "German",   "zh": "Chinese",    "ar": "Arabic",
    "pt": "Portuguese","ja":"Japanese","ko": "Korean",     "it": "Italian",
    "ru": "Russian", "bn": "Bengali",  "ta": "Tamil",
}

# ── Flask ─────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR, static_url_path="")
app.secret_key = FLASK_SECRET_KEY

CORS(app,
     supports_credentials=False,
     origins="*",
     allow_headers=["Content-Type", "Authorization"],
     methods=["GET", "POST", "PATCH", "DELETE", "OPTIONS"])

ALLOWED_EXT = {"pdf", "txt", "docx", "doc", "ppt", "pptx"}

def _allowed(fn):
    return "." in fn and fn.rsplit(".", 1)[1].lower() in ALLOWED_EXT

def _now():
    return datetime.utcnow().isoformat()

def _today():
    return datetime.utcnow().strftime("%Y-%m-%d")

# ════════════════════════════════════════════════════════════════
# CONTENT HASHING  — deduplication key
# SHA-256 of (file bytes + lang + speaker names)
# Same file + same settings → same hash → return saved podcast
# ════════════════════════════════════════════════════════════════
def _content_hash(file_bytes: bytes, lang: str, sp1: str, sp2: str) -> str:
    h = hashlib.sha256()
    h.update(file_bytes)
    h.update(lang.encode())
    h.update(sp1.lower().encode())
    h.update(sp2.lower().encode())
    return h.hexdigest()

# ════════════════════════════════════════════════════════════════
# RATE LIMITING
# ════════════════════════════════════════════════════════════════
def _check_rate_limit(uid: str) -> tuple:
    """Returns (allowed: bool, reason: str)."""
    now   = time.time()
    with _rate_store_lock:
        store = _rate_store[uid]
        store["minute"] = [t for t in store["minute"] if now - t < 60]
        store["day"]    = [t for t in store["day"]    if now - t < 86400]
        if len(store["minute"]) >= RATE_LIMIT_RPM:
            return False, f"Rate limit: max {RATE_LIMIT_RPM} generations per minute. Please wait."
        if len(store["day"]) >= RATE_LIMIT_RPD:
            return False, f"Daily limit: max {RATE_LIMIT_RPD} generations per day reached."
        store["minute"].append(now)
        store["day"].append(now)
    return True, ""

# ════════════════════════════════════════════════════════════════
# USAGE TRACKING  — ElevenLabs character quota per user per day
# ════════════════════════════════════════════════════════════════
def _get_usage(uid: str) -> dict:
    today = _today()
    if db is not None:
        doc = usage_col.find_one({"user_id": uid, "date": today})
        return doc or {"user_id": uid, "date": today, "chars": 0}
    with _usage_store_lock:
        entry = _usage_store.get(uid, {})
        if entry.get("date") != today:
            return {"user_id": uid, "date": today, "chars": 0}
        return entry

def _add_usage(uid: str, chars: int):
    today = _today()
    if db is not None:
        usage_col.update_one(
            {"user_id": uid, "date": today},
            {"$inc": {"chars": chars}},
            upsert=True
        )
    else:
        with _usage_store_lock:
            entry = _usage_store.get(uid, {})
            if entry.get("date") != today:
                _usage_store[uid] = {"user_id": uid, "date": today, "chars": chars}
            else:
                _usage_store[uid]["chars"] = entry.get("chars", 0) + chars

def _check_quota(uid: str) -> tuple:
    """Gate-check before generation using current usage."""
    usage = _get_usage(uid)
    used  = usage.get("chars", 0)
    if used >= FREE_DAILY_CHARS:
        return False, (
            f"Daily character quota exceeded. "
            f"Used {used:,}/{FREE_DAILY_CHARS:,} chars today. Resets at midnight UTC."
        )
    return True, ""

# ════════════════════════════════════════════════════════════════
# SCRIPT CACHE  — keyed by content hash, saves LLM calls
# ════════════════════════════════════════════════════════════════
def _cache_get(content_hash: str) -> dict:
    if db is not None:
        return cache_col.find_one({"content_hash": content_hash}, {"_id": 0})
    return _mem_cache.get(content_hash)

def _cache_set(content_hash: str, data: dict):
    data = {**data, "content_hash": content_hash, "created_at": _now()}
    if db is not None:
        cache_col.update_one(
            {"content_hash": content_hash},
            {"$set": data},
            upsert=True
        )
    else:
        _mem_cache[content_hash] = data

# ════════════════════════════════════════════════════════════════
# AUTH HELPERS
# ════════════════════════════════════════════════════════════════
def _hash(pw):
    return bcrypt.hashpw(pw.encode(), bcrypt.gensalt()).decode()

def _verify(pw, h):
    return bcrypt.checkpw(pw.encode(), h.encode())

def _make_token(uid, email, name):
    now = datetime.now(timezone.utc)
    return jwt.encode(
        {"sub": uid, "email": email, "name": name,
         "iat": int(now.timestamp()),
         "exp": int((now + timedelta(days=30)).timestamp())},
        JWT_SECRET, algorithm="HS256"
    )

def require_auth(f):
    @wraps(f)
    def wrapper(*a, **kw):
        auth = request.headers.get("Authorization", "")
        if not auth.startswith("Bearer "):
            return jsonify({"error": "Authentication required."}), 401
        try:
            request.user = jwt.decode(auth.split()[1], JWT_SECRET, algorithms=["HS256"])
        except jwt.ExpiredSignatureError:
            return jsonify({"error": "Session expired."}), 401
        except jwt.InvalidTokenError:
            return jsonify({"error": "Invalid token."}), 401
        return f(*a, **kw)
    return wrapper

# ── DB helpers ────────────────────────────────────────────────────
def _find_user(email):
    if db is not None:
        return users_col.find_one({"email": email})
    return _mem_users.get(email)

def _create_user(email, name, pw):
    uid  = str(uuid.uuid4())
    user = {"_id": uid, "email": email, "name": name,
            "password": _hash(pw), "created_at": _now()}
    if db is not None:
        users_col.insert_one(user)
    else:
        _mem_users[email] = user
    return uid

# FIX: _get_pods now sorts in-memory results by created_at descending,
# matching the MongoDB path's .sort("created_at", -1) behaviour.
def _get_pods(uid):
    if db is not None:
        return list(podcasts_col.find({"user_id": uid}, {"_id": 0})
                                .sort("created_at", -1).limit(100))
    with _mem_pods_lock:
        pods = [p for p in _mem_pods if p.get("user_id") == uid]
    return sorted(pods, key=lambda p: p.get("created_at", ""), reverse=True)

def _find_pod_by_hash(uid: str, content_hash: str):
    """Return an existing podcast if this exact file+settings were already generated."""
    if db is not None:
        return podcasts_col.find_one(
            {"user_id": uid, "content_hash": content_hash}, {"_id": 0}
        )
    with _mem_pods_lock:
        return next(
            (p for p in _mem_pods
             if p.get("user_id") == uid and p.get("content_hash") == content_hash),
            None
        )

# FIX: _save_pod and _del_pod both acquire _mem_pods_lock before mutating
# the shared list, preventing data races under threaded=True Flask.
def _save_pod(e):
    if db is not None:
        podcasts_col.insert_one(e)
    else:
        with _mem_pods_lock:
            _mem_pods.insert(0, e)

def _del_pod(pid, uid):
    if db is not None:
        podcasts_col.delete_one({"podcast_id": pid, "user_id": uid})
    else:
        with _mem_pods_lock:
            _mem_pods[:] = [
                p for p in _mem_pods
                if not (p["podcast_id"] == pid and p["user_id"] == uid)
            ]

def _find_shared(sid):
    if db is not None:
        return podcasts_col.find_one({"share_id": sid}, {"_id": 0})
    with _mem_pods_lock:
        return next((p for p in _mem_pods if p.get("share_id") == sid), None)

def _rename_pod(pid, uid, title):
    if db is not None:
        podcasts_col.update_one(
            {"podcast_id": pid, "user_id": uid},
            {"$set": {"title": title}}
        )
    else:
        with _mem_pods_lock:
            for p in _mem_pods:
                if p["podcast_id"] == pid and p["user_id"] == uid:
                    p["title"] = title
                    break

# ════════════════════════════════════════════════════════════════
# VOICE ROUTES
# ════════════════════════════════════════════════════════════════
@app.route("/api/voices", methods=["GET"])
def list_voices():
    return jsonify({
        "voices":    get_voices(),
        "languages": SUPPORTED_LANGUAGES,
        "local_tts": LOCAL_TTS_AVAILABLE,
    }), 200

@app.route("/api/voices/<voice_id>/preview", methods=["GET"])
def preview_voice(voice_id):
    """
    Optimization: Use free local TTS (pyttsx3) for voice previews.
    This saves all ElevenLabs characters for actual podcast generation.
    Falls back to ElevenLabs only if pyttsx3 is not installed.
    """
    voices = get_voices()
    vname  = next((v["name"] for v in voices if v["voice_id"] == voice_id), "your host")
    text   = (f"Hi, I'm {vname}! I'll be one of your podcast hosts today. "
              "Welcome to PodcastAI, where documents become engaging conversations.")

    if LOCAL_TTS_AVAILABLE:
        try:
            wav = _local_tts(text)
            return jsonify({"ok": True, "audio": base64.b64encode(wav).decode(),
                            "name": vname, "source": "local"}), 200
        except Exception as e:
            log.warning("Local TTS preview failed: %s — falling back to ElevenLabs.", e)

    if not el_client:
        return jsonify({"ok": False, "error": "ElevenLabs API key not configured."}), 503
    try:
        pcm = _tts_line(text, voice_id)
        wav = _pcm_to_wav(pcm)
        return jsonify({"ok": True, "audio": base64.b64encode(wav).decode(),
                        "name": vname, "source": "elevenlabs"}), 200
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ════════════════════════════════════════════════════════════════
# AUTH ROUTES
# ════════════════════════════════════════════════════════════════
@app.route("/api/auth/signup", methods=["POST"])
def signup():
    d     = request.get_json() or {}
    name  = (d.get("name")     or "").strip()
    email = (d.get("email")    or "").strip().lower()
    pw    = (d.get("password") or "")
    if not name or not email or not pw:
        return jsonify({"error": "Name, email and password are required."}), 400
    if len(pw) < 8:
        return jsonify({"error": "Password must be at least 8 characters."}), 400
    if _find_user(email):
        return jsonify({"error": "An account with this email already exists."}), 409
    uid   = _create_user(email, name, pw)
    token = _make_token(uid, email, name)
    return jsonify({"token": token, "user": {"id": uid, "email": email, "name": name}}), 201

@app.route("/api/auth/login", methods=["POST"])
def login():
    d     = request.get_json() or {}
    email = (d.get("email")    or "").strip().lower()
    pw    = (d.get("password") or "")
    user  = _find_user(email)
    if not user or not _verify(pw, user["password"]):
        return jsonify({"error": "Invalid email or password."}), 401
    uid   = str(user["_id"])
    token = _make_token(uid, email, user["name"])
    return jsonify({"token": token,
                    "user": {"id": uid, "email": email, "name": user["name"]}}), 200

@app.route("/api/auth/me", methods=["GET"])
@require_auth
def me():
    u = request.user
    return jsonify({"user": {"id": u["sub"], "email": u["email"], "name": u["name"]}}), 200

# ════════════════════════════════════════════════════════════════
# USAGE ROUTE  — frontend can show the user their quota
# ════════════════════════════════════════════════════════════════
@app.route("/api/usage", methods=["GET"])
@require_auth
def get_usage_route():
    uid   = request.user["sub"]
    usage = _get_usage(uid)
    used  = usage.get("chars", 0)
    return jsonify({
        "used":      used,
        "limit":     FREE_DAILY_CHARS,
        "remaining": max(0, FREE_DAILY_CHARS - used),
        "date":      _today(),
        "pct":       round(used / FREE_DAILY_CHARS * 100, 1),
    }), 200

# ════════════════════════════════════════════════════════════════
# PODCAST CRUD
# ════════════════════════════════════════════════════════════════
@app.route("/api/podcasts", methods=["GET"])
@require_auth
def list_podcasts():
    items = _get_pods(request.user["sub"])
    return jsonify({"podcasts": [
        {
            "podcast_id":     p.get("podcast_id"),
            "title":          p.get("title") or p.get("filename", "Untitled"),
            "filename":       p.get("filename"),
            "created_at":     p.get("created_at"),
            "share_id":       p.get("share_id"),
            "language":       p.get("language", "en"),
            "speaker1":       p.get("speaker1", "Alex"),
            "speaker2":       p.get("speaker2", "Jordan"),
            "script_preview": (p.get("script") or "")[:200],
            "from_cache":     p.get("from_cache", False),
            "chars_used":     p.get("chars_used", 0),
        }
        for p in items
    ]}), 200

@app.route("/api/podcasts/<pid>", methods=["GET"])
@require_auth
def get_podcast(pid):
    p = next((x for x in _get_pods(request.user["sub"]) if x.get("podcast_id") == pid), None)
    if not p:
        return jsonify({"error": "Not found."}), 404
    return jsonify({"podcast": {k: v for k, v in p.items() if k != "_id"}}), 200

@app.route("/api/podcasts/<pid>", methods=["DELETE"])
@require_auth
def delete_podcast(pid):
    _del_pod(pid, request.user["sub"])
    return jsonify({"ok": True}), 200

@app.route("/api/podcasts/<pid>/rename", methods=["PATCH"])
@require_auth
def rename_podcast(pid):
    t = ((request.get_json() or {}).get("title") or "").strip()
    if not t:
        return jsonify({"error": "Title required."}), 400
    _rename_pod(pid, request.user["sub"], t)
    return jsonify({"ok": True}), 200

@app.route("/api/share/<sid>", methods=["GET"])
def get_shared(sid):
    p = _find_shared(sid)
    if not p:
        return jsonify({"error": "Shared podcast not found."}), 404
    return jsonify({
        "title":      p.get("title") or p.get("filename", "Podcast"),
        "script":     p.get("script", ""),
        "audio":      p.get("audio", ""),
        "filename":   p.get("filename", ""),
        "created_at": p.get("created_at", ""),
        "language":   p.get("language", "en"),
        "speaker1":   p.get("speaker1", "Alex"),
        "speaker2":   p.get("speaker2", "Jordan"),
    }), 200

# ════════════════════════════════════════════════════════════════
# TEXT EXTRACTION
# ════════════════════════════════════════════════════════════════
def extract_text(fb, fname):
    ext = fname.rsplit(".", 1)[1].lower()
    if ext == "pdf":
        text = _extract_pdf(fb)
    elif ext in ("docx", "doc"):
        text = "\n".join(p.text for p in DocxDocument(io.BytesIO(fb)).paragraphs)
    elif ext in ("ppt", "pptx"):
        prs   = Presentation(io.BytesIO(fb))
        lines = [
            para.text
            for slide in prs.slides
            for shape in slide.shapes
            if shape.has_text_frame
            for para in shape.text_frame.paragraphs
        ]
        text = "\n".join(lines)
    elif ext == "txt":
        try:
            text = fb.decode("utf-8")
        except Exception:
            text = fb.decode("latin-1", errors="replace")
    else:
        raise ValueError(f"Unsupported file type: .{ext}")
    cleaned = _clean(text)
    if not cleaned:
        raise ValueError("No text could be extracted. If this is a scanned PDF, install tesseract + pdf2image.")
    return cleaned

def _extract_pdf(fb):
    reader = pypdf.PdfReader(io.BytesIO(fb))
    total  = len(reader.pages)
    if total == 0:
        raise ValueError("PDF has no pages.")
    pages, empty = [], 0
    for page in reader.pages:
        t = page.extract_text() or ""
        if t.strip():
            pages.append(t)
        else:
            empty += 1
    if (total - empty) / total >= 0.5:
        return "\n".join(pages)
    if not OCR_AVAILABLE:
        raise ValueError("Scanned PDF detected — install tesseract-ocr + pdf2image for OCR support.")
    return _ocr_pdf(fb, total)

def _ocr_pdf(fb, total):
    dpi    = 150 if total > 30 else 200
    t0     = time.time()
    images = convert_from_bytes(fb, dpi=dpi, fmt="jpeg", thread_count=4)
    res    = {}
    def _do(args):
        i, img = args
        try:
            return i, pytesseract.image_to_string(img, lang="eng", config="--oem 3 --psm 3")
        except Exception:
            return i, ""
    with ThreadPoolExecutor(max_workers=4) as ex:
        for i, t in ex.map(_do, enumerate(images)):
            res[i] = t
    combined = "\n".join(res.get(i, "") for i in range(len(images)))
    log.info("OCR: %d chars in %.1fs.", len(combined), time.time() - t0)
    if not combined.strip():
        raise ValueError("OCR returned no text.")
    return combined

def _clean(raw):
    t = re.sub(r"[ \t]+",         " ",    raw)
    t = re.sub(r"\n{3,}",         "\n\n", t)
    t = re.sub(r"[^\x20-\x7E\n]", "",    t)
    return t.strip()

# ════════════════════════════════════════════════════════════════
# BOILERPLATE FILTER
# ════════════════════════════════════════════════════════════════
_BP_LINE = re.compile("|".join([
    r"^page\s*(no|number)?[\s.:]*\d*$", r"^-\s*\d+\s*-$",
    r"^\d+\s*\|?\s*p\s*a\s*g\s*e$",
    r"signature\s*of\s*(the\s*)?(faculty|examiner|supervisor|guide|hod)",
    r"^(faculty|examiner|hod|supervisor|guide)\s*(signature)?[\s.:_-]*$",
    r"verified\s*by", r"approved\s*by",
    r"roll\s*(no|number)\s*[:\-.]?\s*\d*",
    r"register\s*(no|number)\s*[:\-.]?\s*\d*",
    r"^(table\s*of\s*contents?|index|contents?)$",
    r"(university|institute|college|department)\s*of\s*(technology|science|engineering|arts)",
    r"(autonomous|affiliated\s*to|accredited\s*by)",
    r"(naac|nba|ugc|aicte)\s*(accredited|approved|recognized)?",
    r"(academic\s*year|batch)\s*[:\-.]?\s*\d{4}",
    r"^about\s*(the\s*)?(author|writer)s?$",
    r"all\s*rights?\s*reserved", r"copyright\s*©?\s*\d{4}",
    r"isbn\s*[:\-.]?\s*[\d\-]+", r"www\.[a-z0-9\-]+\.[a-z]{2,}",
    r"^https?://", r"^[_\-=*#~.]{3,}$", r"^[\W\s]{0,3}$",
]), re.IGNORECASE)

_BP_PARA = re.compile(
    r"(table\s*of\s*contents?|pin\s*code|phone\s*no|fax\s*no)", re.IGNORECASE
)

def _regex_filter(text):
    kept = []
    for para in re.split(r"\n{2,}", text):
        if _BP_PARA.search(para):
            continue
        lines = para.splitlines()
        good  = [l for l in lines if not _BP_LINE.match(l.strip())]
        if len(good) < 2 and len(lines) > 3:
            continue
        c = "\n".join(good).strip()
        if c:
            kept.append(c)
    return "\n\n".join(kept)

def _llm_filter(chunks):
    if not groq_client:
        return chunks
    kept = []
    for chunk in chunks:
        if len(chunk.split()) < 20:
            continue
        try:
            v = _groq(
                f"Does this chunk have MEANINGFUL CONTENT for a podcast?\n"
                f"Reply ONLY: KEEP or DISCARD\n\nCHUNK:\n{chunk[:1500]}",
                max_tokens=5
            ).strip().upper()
            if v.startswith("KEEP"):
                kept.append(chunk)
        except Exception:
            kept.append(chunk)
    return kept

def filter_boilerplate(text):
    after = _regex_filter(text)
    if not after.strip():
        after = text
    cands = [c.strip() for c in textwrap.wrap(after, width=3000, break_long_words=False) if c.strip()]
    kept  = _llm_filter(cands)
    return "\n\n".join(kept) if kept else after

# ════════════════════════════════════════════════════════════════
# GROQ WRAPPER
# ════════════════════════════════════════════════════════════════
def _groq(prompt, max_tokens=1024, retries=3):
    if not groq_client:
        raise RuntimeError("Groq API key not configured. Add GROQ_API_KEY to .env")
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            r = groq_client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
                temperature=0.7,
            )
            return r.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
            log.warning("Groq attempt %d/%d: %s", attempt, retries, e)
            if attempt < retries:
                time.sleep(3 * attempt)
    raise RuntimeError(f"Groq failed after {retries} attempts: {last_err}")

# ════════════════════════════════════════════════════════════════
# PAGE ESTIMATION & WORD TARGET
# ════════════════════════════════════════════════════════════════
WORDS_PER_PAGE = 320
WPM            = 140

def _estimate_pages(text: str) -> float:
    return max(1.0, len(text.split()) / WORDS_PER_PAGE)

def _target_script_words(pages: float) -> int:
    if pages <= 2:
        target = int(pages * 140)
    elif pages <= 10:
        target = int(280 + (pages - 2) * 115)
    else:
        target = int(1200 + (pages - 10) * 30)
    return max(140, min(target, 1800))

# ════════════════════════════════════════════════════════════════
# SUMMARISE
# ════════════════════════════════════════════════════════════════
def summarise_chunks(chunks: list, target_script_words: int) -> list:
    total_chunks         = len(chunks)
    total_summary_budget = int(target_script_words / 0.85)
    words_each           = max(60, min(total_summary_budget // max(total_chunks, 1), 500))
    log.info("Summarising %d chunks | %d words each | script target: %d",
             total_chunks, words_each, target_script_words)
    sums = []
    for i, chunk in enumerate(chunks, 1):
        sums.append(_groq(
            f"Summarise in {words_each}-{words_each + 30} words. "
            f"Preserve ALL key ideas, facts, examples, and arguments. Be concise but complete:\n\nTEXT:\n{chunk}",
            max_tokens=max(200, words_each * 2),
        ))
        if i % 5 == 0 or i == total_chunks:
            log.info("  Summarised %d/%d", i, total_chunks)
    return sums

def hierarchical_summarise(sums: list, target_script_words: int) -> list:
    if len(sums) <= 20:
        return sums
    desired_blocks = max(5, min(int(target_script_words / 200), 30))
    batch_size     = max(2, len(sums) // desired_blocks)
    batches        = [sums[i:i + batch_size] for i in range(0, len(sums), batch_size)]
    total_sw       = sum(len(s.split()) for s in sums)
    wpb            = max(150, total_sw // max(len(batches), 1))
    result = []
    for b in batches:
        cw = sum(len(s.split()) for s in b)
        tw = max(150, min(cw // 2, wpb))
        result.append(_groq(
            f"Merge these summaries into one cohesive passage of {tw}-{tw + 50} words. "
            f"Preserve ALL distinct topics, facts, and arguments:\n\n"
            + "\n\n".join(f"[{i+1}] {s}" for i, s in enumerate(b)),
            max_tokens=max(400, tw * 2),
        ))
    return result

# ════════════════════════════════════════════════════════════════
# SCRIPT GENERATION
# ════════════════════════════════════════════════════════════════
SCRIPT_PROMPT = """\
You are a podcast scriptwriter. Write a natural, engaging podcast between {sp1} and {sp2}.

RULES:
- Format EVERY line EXACTLY as: {sp1}: <text>  OR  {sp2}: <text>
- No narration, stage directions, or any other prefix.
- Alternate speakers naturally. Short conversational sentences.
- No repeated points. EXACTLY ~{target} words total. Cover ALL key points.
- {sp1} greets the listener and introduces the topic.
- Both hosts wrap up warmly at the end.
- Stay strictly within the word count.

SUMMARIES:
{summaries}

BEGIN SCRIPT:
"""

def generate_script(sums, sp1="Alex", sp2="Jordan", target_script_words=500):
    combined = "\n\n".join(sums)
    if len(combined) > 60000:
        combined = combined[:60000] + "\n[truncated]"
    log.info("Generating script (target ~%d words)…", target_script_words)
    return _groq(
        SCRIPT_PROMPT.format(sp1=sp1, sp2=sp2, summaries=combined, target=target_script_words),
        max_tokens=min(int(target_script_words * 1.6) + 500, 8192),
    )

# ════════════════════════════════════════════════════════════════
# TRANSLATION
# ════════════════════════════════════════════════════════════════
TRANSLATE_PROMPT = """\
Translate this podcast script into {lang}.
CRITICAL:
- Keep speaker prefixes EXACTLY as-is: "{sp1}:" or "{sp2}:" — never translate names.
- Translate only the dialogue text after the colon.
- Keep sentences short and speech-friendly.
- Output ONLY the translated script, nothing else.

SCRIPT:
{script}
"""

def translate_script(script, lang_code, sp1="Alex", sp2="Jordan"):
    if lang_code == "en":
        return script
    lang_name = SUPPORTED_LANGUAGES.get(lang_code, lang_code)
    log.info("Translating to %s…", lang_name)
    lines   = script.strip().splitlines()
    batches = [lines[i:i + 50] for i in range(0, len(lines), 50)]
    parts   = []
    for batch in batches:
        try:
            parts.append(_groq(
                TRANSLATE_PROMPT.format(lang=lang_name, sp1=sp1, sp2=sp2,
                                        script="\n".join(batch)),
                max_tokens=min(len(" ".join(batch).split()) * 3, 4096),
                retries=2,
            ))
        except Exception as e:
            log.warning("Translation batch failed: %s — using original.", e)
            parts.append("\n".join(batch))
    return "\n\n".join(parts)

# ════════════════════════════════════════════════════════════════
# AUDIO
# ════════════════════════════════════════════════════════════════
SAMPLE_RATE     = 22050
CHANNELS        = 1
BITS_PER_SAMPLE = 16
PAUSE_MS        = 400
TTS_MAX         = 250

def _pcm_pause(ms):
    return b"\x00\x00" * int(SAMPLE_RATE * ms / 1000)

def _pcm_to_wav(pcm):
    br = SAMPLE_RATE * CHANNELS * BITS_PER_SAMPLE // 8
    ba = CHANNELS * BITS_PER_SAMPLE // 8
    ds = len(pcm)
    return struct.pack(
        "<4sI4s4sIHHIIHH4sI",
        b"RIFF", 36 + ds, b"WAVE",
        b"fmt ", 16, 1, CHANNELS, SAMPLE_RATE, br, ba, BITS_PER_SAMPLE,
        b"data", ds,
    ) + pcm

def _split_tts(text, max_chars=TTS_MAX):
    if len(text) <= max_chars:
        return [text]
    sentences   = re.split(r"(?<=[.!?])\s+", text)
    chunks, cur = [], ""
    for s in sentences:
        if len(cur) + len(s) + 1 <= max_chars:
            cur = (cur + " " + s).strip()
        else:
            if cur:
                chunks.append(cur)
            cur = s[:max_chars]
    if cur:
        chunks.append(cur)
    return chunks or [text[:max_chars]]

# ── ElevenLabs TTS  (premium quality, used for final output only) ─
def _tts_line(text, voice_id):
    if not el_client:
        raise RuntimeError("ElevenLabs API key not configured. Add ELEVENLABS_API_KEY to .env")
    gen    = el_client.text_to_speech.convert(
        text=text, voice_id=voice_id,
        model_id="eleven_turbo_v2", output_format="pcm_22050"
    )
    chunks = [c for c in gen if isinstance(c, bytes) and c]
    if not chunks:
        raise ValueError("ElevenLabs returned zero bytes — check API key and credits.")
    return b"".join(chunks)

# ── Local TTS  (free offline, used for voice previews only) ──────
def _local_tts(text: str) -> bytes:
    """Synthesise speech with pyttsx3 (free, no API calls). Returns WAV bytes."""
    if not LOCAL_TTS_AVAILABLE:
        raise RuntimeError("pyttsx3 not installed. Run: pip install pyttsx3")
    import tempfile
    engine = pyttsx3.init()
    engine.setProperty("rate", 160)
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
        tmp = f.name
    try:
        engine.save_to_file(text, tmp)
        engine.runAndWait()
        with open(tmp, "rb") as fh:
            return fh.read()
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass

# ════════════════════════════════════════════════════════════════
# PARSE & GENERATE AUDIO
# ════════════════════════════════════════════════════════════════
def parse_script(script, sp1="Alex", sp2="Jordan"):
    pat      = re.compile(rf"^({re.escape(sp1)}|{re.escape(sp2)})\s*:\s*(.+)$", re.IGNORECASE)
    dialogue = []
    for line in script.splitlines():
        line = line.strip()
        if not line:
            continue
        m = pat.match(line)
        if not m:
            continue
        speaker = m.group(1).capitalize()
        text    = m.group(2).strip()
        if not text:
            continue
        dialogue.append({"speaker": speaker, "text": text})
    return dialogue

def generate_audio(dialogue, voice1, voice2, sp1="Alex", sp2="Jordan") -> tuple:
    """Returns (wav_bytes, total_chars_consumed_by_elevenlabs)."""
    log.info("Generating audio: %d lines, v1=%s v2=%s…", len(dialogue), voice1, voice2)
    segments    = []
    pause       = _pcm_pause(PAUSE_MS)
    fails       = 0
    total_chars = 0

    for idx, entry in enumerate(dialogue, 1):
        if idx % 20 == 0 or idx == len(dialogue):
            log.info("  Audio %d/%d", idx, len(dialogue))
        vid       = voice1 if entry["speaker"].lower() == sp1.lower() else voice2
        got_audio = False
        for sub in _split_tts(entry["text"]):
            total_chars += len(sub)
            try:
                segments.append(_tts_line(sub, vid))
                got_audio = True
            except Exception as e:
                fails += 1
                log.warning("  TTS line %d: %s", idx, e)
        if got_audio:
            segments.append(pause)

    if not segments:
        raise RuntimeError(
            "All TTS calls failed.\n"
            "Check ELEVENLABS_API_KEY and credits.\n"
            "Test: http://localhost:5000/api/test-tts"
        )
    if fails:
        log.warning("  %d chunk(s) failed — audio may have gaps.", fails)

    pcm = b"".join(segments)
    wav = _pcm_to_wav(pcm)
    dur = (len(wav) - 44) / (SAMPLE_RATE * CHANNELS * BITS_PER_SAMPLE // 8)
    log.info("Audio: %.1f s (%.1f min). EL chars: %d", dur, dur / 60, total_chars)
    return wav, total_chars

# ════════════════════════════════════════════════════════════════
# PIPELINE
# ════════════════════════════════════════════════════════════════
def run_pipeline(fb, fname, voice1, voice2, lang="en",
                 sp1="Alex", sp2="Jordan", content_hash="") -> dict:
    t0 = time.time()
    log.info("=== Pipeline: %s | lang=%s | v1=%s v2=%s ===", fname, lang, voice1, voice2)

    # ── Step 1: Script cache check — skip all LLM work if same file seen before
    cached = _cache_get(content_hash) if content_hash else None
    if cached and cached.get("script"):
        log.info("Script cache HIT — skipping extraction + summarisation + script gen.")
        script = cached["script"]
    else:
        log.info("Script cache MISS — running full LLM pipeline.")
        text     = extract_text(fb, fname)
        filtered = filter_boilerplate(text)
        chunks   = [
            c.strip()
            for c in textwrap.wrap(filtered, width=3000, break_long_words=False)
            if c.strip()
        ]
        if not chunks:
            raise ValueError("No meaningful content remained after filtering.")

        est_pages           = _estimate_pages(filtered)
        target_script_words = _target_script_words(est_pages)
        log.info("Pages: %.1f → script target: %d words (~%.1f min)",
                 est_pages, target_script_words, target_script_words / WPM)

        sums    = summarise_chunks(chunks, target_script_words)
        final_s = hierarchical_summarise(sums, target_script_words)
        script  = generate_script(final_s, sp1=sp1, sp2=sp2,
                                  target_script_words=target_script_words)
        if lang != "en":
            script = translate_script(script, lang, sp1=sp1, sp2=sp2)

        # Cache script so future same-file requests skip all LLM calls
        if content_hash:
            _cache_set(content_hash, {"script": script})

    # ── Step 2: ElevenLabs audio synthesis (premium, final output only)
    dialogue = parse_script(script, sp1=sp1, sp2=sp2)
    if not dialogue:
        raise ValueError("Script parsing produced no dialogue lines.")

    wav, chars_used = generate_audio(dialogue, voice1, voice2, sp1, sp2)
    log.info("=== Pipeline DONE in %.1fs ===", time.time() - t0)
    return {"script": script, "audio": base64.b64encode(wav).decode(), "chars_used": chars_used}

# ════════════════════════════════════════════════════════════════
# GENERATE ROUTE  ← all optimizations applied here
# ════════════════════════════════════════════════════════════════
@app.route("/api/generate-podcast", methods=["POST"])
@require_auth
def generate_podcast():
    if not groq_client or not el_client:
        return jsonify({"error": "Server not configured: missing API keys in .env"}), 503
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    f = request.files["file"]
    if not f.filename or not _allowed(f.filename):
        return jsonify({"error": "Unsupported file type. Use PDF, TXT, DOCX, or PPTX."}), 400

    raw = f.read()
    if len(raw) > 10 * 1024 * 1024:
        return jsonify({"error": "File exceeds 10 MB."}), 413
    if not raw:
        return jsonify({"error": "Empty file."}), 400

    dv1, dv2   = _default_voices()
    voice1     = request.form.get("voice1",     "").strip() or dv1
    voice2     = request.form.get("voice2",     "").strip() or dv2
    same_voice = request.form.get("same_voice", "false").lower() == "true"
    lang       = request.form.get("language",   "en").strip()
    sp1        = request.form.get("speaker1",   "Alex").strip()   or "Alex"
    sp2        = request.form.get("speaker2",   "Jordan").strip() or "Jordan"

    if same_voice:
        voice2 = voice1
    if lang not in SUPPORTED_LANGUAGES:
        lang = "en"

    uid = request.user["sub"]
    log.info("user=%s file=%s lang=%s v1=%s v2=%s",
             request.user["email"], f.filename, lang, voice1, voice2)

    # ── Guard 1: Rate limiting ────────────────────────────────────
    allowed, reason = _check_rate_limit(uid)
    if not allowed:
        return jsonify({"error": reason}), 429

    # ── Guard 2: Daily char quota ─────────────────────────────────
    ok, reason = _check_quota(uid)
    if not ok:
        return jsonify({"error": reason}), 429

    # ── Guard 3: Deduplication — same file already generated? ─────
    chash    = _content_hash(raw, lang, sp1, sp2)
    existing = _find_pod_by_hash(uid, chash)
    if existing:
        log.info("Duplicate — returning saved podcast id=%s (0 EL chars used).",
                 existing["podcast_id"])
        return jsonify({
            "podcast_id": existing["podcast_id"],
            "script":     existing["script"],
            "audio":      existing["audio"],
            "share_id":   existing.get("share_id", ""),
            "speaker1":   existing.get("speaker1", sp1),
            "speaker2":   existing.get("speaker2", sp2),
            "from_cache": True,
            "chars_used": 0,
            "message":    "This file was already generated. Returning your saved podcast.",
        }), 200

    # ── Run pipeline ──────────────────────────────────────────────
    try:
        result = run_pipeline(raw, f.filename, voice1, voice2, lang, sp1, sp2,
                              content_hash=chash)
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 502
    except Exception as e:
        import traceback
        log.error(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

    # ── Track actual EL char usage ────────────────────────────────
    _add_usage(uid, result["chars_used"])

    share_id   = uuid.uuid4().hex[:16]
    podcast_id = str(uuid.uuid4())
    _save_pod({
        "podcast_id":   podcast_id,
        "user_id":      uid,
        "filename":     f.filename,
        "title":        f.filename.rsplit(".", 1)[0],
        "script":       result["script"],
        "audio":        result["audio"],
        "created_at":   _now(),
        "share_id":     share_id,
        "language":     lang,
        "voice1":       voice1,
        "voice2":       voice2,
        "speaker1":     sp1,
        "speaker2":     sp2,
        "content_hash": chash,
        "chars_used":   result["chars_used"],
        "from_cache":   False,
    })

    return jsonify({
        "podcast_id": podcast_id,
        "script":     result["script"],
        "audio":      result["audio"],
        "share_id":   share_id,
        "speaker1":   sp1,
        "speaker2":   sp2,
        "from_cache": False,
        "chars_used": result["chars_used"],
    }), 200

# ════════════════════════════════════════════════════════════════
# STATUS / DIAGNOSTICS
# ════════════════════════════════════════════════════════════════
@app.route("/api/status", methods=["GET"])
def status():
    try:
        import elevenlabs as _el
        el_ver = getattr(_el, "__version__", "?")
    except Exception:
        el_ver = "?"
    return jsonify({
        "status":           "online",
        "db":               "mongodb" if db is not None else "memory",
        "ocr":              OCR_AVAILABLE,
        "local_tts":        LOCAL_TTS_AVAILABLE,
        "elevenlabs_sdk":   el_ver,
        "voice_count":      len(get_voices()),
        "languages":        len(SUPPORTED_LANGUAGES),
        "groq_ready":       bool(groq_client),
        "tts_ready":        bool(el_client),
        "free_daily_chars": FREE_DAILY_CHARS,
        "rate_limit_rpm":   RATE_LIMIT_RPM,
        "rate_limit_rpd":   RATE_LIMIT_RPD,
    }), 200

@app.route("/api/test-tts", methods=["GET"])
def test_tts():
    if not el_client:
        return jsonify({"ok": False, "error": "ElevenLabs API key not configured."}), 503
    vid = get_voices()[0]["voice_id"] if get_voices() else FALLBACK_VOICES[0]["voice_id"]
    try:
        pcm = _tts_line("PodcastAI audio test.", vid)
        wav = _pcm_to_wav(pcm)
        dur = (len(wav) - 44) / (SAMPLE_RATE * CHANNELS * BITS_PER_SAMPLE // 8)
        return jsonify({"ok": True, "pcm_bytes": len(pcm),
                        "duration_sec": round(dur, 2), "voice_id": vid})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/test-groq", methods=["GET"])
def test_groq():
    if not groq_client:
        return jsonify({"ok": False, "error": "Groq API key not configured."}), 503
    try:
        reply = _groq("Say exactly: GROQ OK", max_tokens=10)
        return jsonify({"ok": True, "reply": reply, "model": GROQ_MODEL})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ════════════════════════════════════════════════════════════════
# STATIC / SPA FALLBACK
# ════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/<path:path>")
def serve(path):
    if path.startswith("api/"):
        return jsonify({"error": "Not found"}), 404
    file_path = os.path.join(BASE_DIR, path)
    if os.path.isfile(file_path):
        return send_from_directory(BASE_DIR, path)
    return send_from_directory(BASE_DIR, "index.html")

# ════════════════════════════════════════════════════════════════
# ENTRY POINT
# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("\n" + "=" * 62)
    print(f"  PodcastAI  →  http://localhost:{port}")
    print(f"  DB:           {'MongoDB' if db is not None else 'In-memory'}")
    print(f"  OCR:          {'Enabled' if OCR_AVAILABLE        else 'Disabled'}")
    if LOCAL_TTS_AVAILABLE:
        print(f"  Local TTS:    Enabled (pyttsx3)")
    else:
        print(f"  Local TTS:    Disabled  →  pip install pyttsx3")
        print(f"                (Linux also needs: sudo apt install espeak)")
    print(f"  Groq:         {'Ready ✓' if groq_client else '✗ MISSING KEY'}")
    print(f"  ElevenLabs:   {'Ready ✓' if el_client   else '✗ MISSING KEY'}")
    print(f"  Quota/day:    {FREE_DAILY_CHARS:,} chars per user")
    print(f"  Rate limit:   {RATE_LIMIT_RPM} req/min  |  {RATE_LIMIT_RPD} req/day  per user")
    try:
        vv = get_voices()
        print(f"  Voices:       {len(vv)} loaded from ElevenLabs")
    except Exception as e:
        print(f"  Voices:       fallback ({e})")
    print("=" * 62 + "\n")
    app.run(
        host="0.0.0.0",
        port=port,
        debug=os.environ.get("FLASK_ENV") == "development",
        threaded=True,
        use_reloader=False,
    )