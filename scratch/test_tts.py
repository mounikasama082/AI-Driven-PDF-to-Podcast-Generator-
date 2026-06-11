import pyttsx3
try:
    engine = pyttsx3.init()
    engine.say("Testing pyttsx3 on Windows")
    engine.runAndWait()
    print("pyttsx3 initialized successfully")
except Exception as e:
    print(f"pyttsx3 failed: {e}")
