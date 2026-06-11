# AI Podcast Generator

An optimized Flask & Vanilla JS single-page web application that converts documents (PDF, DOCX, PPTX, TXT) into engaging, podcast-style audio conversations using LLMs and Text-to-Speech (TTS).

---

## 🚀 Key Features

*   **Document Text Extraction**: Supports text files, PDF extraction (with hierarchical fallback to Tesseract OCR for scanned documents), Word documents (`.docx`), and PowerPoint presentations (`.pptx`).
*   **Boilerplate Filtering**: Employs Regex and LLM-driven filtering to strip out academic and administrative boilerplate (e.g. roll numbers, headers/footers, universities) before generating scripts.
*   **Groq LLM Generation**: Uses the fast `llama-3.3-70b-versatile` model on Groq to write natural, engaging podcast conversations between two hosts.
*   **Dual-Tier TTS System (Optimized & Free)**:
    *   **Local TTS Previews**: Uses `pyttsx3` locally on your machine for zero-cost audio previews, saving ElevenLabs characters.
    *   **ElevenLabs Final Audio**: Generates high-fidelity final podcast audio using professional ElevenLabs voices.
*   **Content-Hash Deduplication**: Hashes document inputs along with configurations to bypass regeneration and save API costs if the same file is submitted with identical settings.
*   **User Management & Quotas**: Built-in authentication, daily character quotas, and rate limiting (requests/minute and requests/day).
*   **Flexible Storage**: Automatic fallback to local in-memory storage if MongoDB is unavailable.

---

## 🛠️ Tech Stack

*   **Backend**: Python, Flask, Flask-CORS, PyJWT, Bcrypt
*   **Database**: MongoDB (optional, with automatic in-memory fallback)
*   **LLM API**: Groq API
*   **TTS API**: ElevenLabs API & `pyttsx3` (Local)
*   **Frontend**: HTML5, Vanilla CSS, Vanilla JavaScript (Single-Page App)

---

## ⚙️ Prerequisites

*   Python 3.8 or higher
*   (Optional but recommended for scanned PDFs) [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
*   (Optional but recommended for scanned PDFs) [poppler-utils](https://poppler.freedesktop.org/) (required by `pdf2image`)

---

## 💻 Installation & Setup

1.  **Clone or Download this Repository**
    ```bash
    git clone <your-repository-url>
    cd AI-Podcast-Generator
    ```

2.  **Create a Virtual Environment**
    ```bash
    python -m venv venv
    # On Windows:
    .\venv\Scripts\activate
    # On macOS/Linux:
    source venv/bin/activate
    ```

3.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Set Up Environment Variables**
    *   Copy the `.env.example` template:
        ```bash
        copy .env.example .env
        ```
    *   Open `.env` and fill in your API keys:
        *   Get a Groq API key from the [Groq Console](https://console.groq.com/).
        *   Get an ElevenLabs API key from [ElevenLabs](https://elevenlabs.io/).
    *   Generate secure secret keys for Flask and JWT:
        ```bash
        python generate_keys.py
        ```
        Copy-paste the generated keys into `JWT_SECRET` and `FLASK_SECRET_KEY` in your `.env`.

5.  **Run the Flask Server**
    ```bash
    python app.py
    ```
    The application will start on `http://localhost:5000`.

---

## 📂 Project Structure

```text
├── app.py                # Main Flask server & API endpoints
├── index.html            # Frontend Single Page App (HTML, CSS, JS)
├── requirements.txt      # Python dependencies
├── generate_keys.py      # Secret key generation utility
├── .env                  # Environment secrets (IGNORED BY GIT)
├── .env.example          # Environment secrets template
├── .gitignore            # Git ignore rules
└── scratch/              # Debug & scratch testing scripts
```

---

## 📝 License

This project is open-source and available under the [MIT License](LICENSE).

