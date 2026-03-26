import os
import sys
import re
import json
import subprocess
import pyttsx3
import speech_recognition as sr

# Initialize engine globally to avoid re-initialization issues
engine = None

def init_engine():
    global engine
    if engine is None:
        try:
            engine = pyttsx3.init()
            # Set voice to deep male voice on Windows
            if sys.platform == "win32":
                voices = engine.getProperty('voices')
                # Try to find a male voice
                for voice in voices:
                    if "male" in voice.name.lower() or "david" in voice.name.lower() or "mark" in voice.name.lower():
                        engine.setProperty('voice', voice.id)
                        break
                # Set properties for better speech
                engine.setProperty('rate', 175)  # Speed of speech
                engine.setProperty('volume', 1.0)  # Volume (0.0 to 1.0)
            elif sys.platform == "darwin":
                voices = engine.getProperty('voices')
                for voice in voices:
                    if "Daniel" in voice.name:
                        engine.setProperty('voice', voice.id)
                        break
        except Exception as e:
            print(f"TTS Engine initialization error: {e}")
    return engine

# Global flag to check if Jarvis is speaking
is_speaking = False

def _clean_for_speech(text: str) -> str:
    """Convert text to a clean, speakable string."""
    if not text:
        return "Done."
    text = str(text).strip()

    # If it looks like JSON, try to extract a human-readable message
    if text.startswith("{") or text.startswith("["):
        try:
            parsed = json.loads(text)
            if isinstance(parsed, dict):
                # Prefer "message" field, then "error", then "content", then "status"
                for key in ("message", "error", "content", "status", "result"):
                    val = parsed.get(key)
                    if val and isinstance(val, str) and len(val) > 2:
                        text = val
                        break
                else:
                    text = "Task completed."
        except Exception:
            # Not valid JSON - just clean up the braces noise
            text = re.sub(r'\{.*?\}', '', text, flags=re.DOTALL).strip() or "Task completed."

    # Strip markdown syntax: **bold**, *italic*, ##heading, `code`, ---
    text = re.sub(r'\*{1,2}(.*?)\*{1,2}', r'\1', text)
    text = re.sub(r'#{1,6}\s*', '', text)
    text = re.sub(r'`{1,3}(.*?)`{1,3}', r'\1', text, flags=re.DOTALL)
    text = re.sub(r'-{2,}', '', text)
    text = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', text)  # [link text](url)
    text = re.sub(r'\n{2,}', '. ', text)
    text = re.sub(r'\n', ' ', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()


def speak(text):
    global is_speaking
    text = _clean_for_speech(text)

    # Print first so user sees it even if audio fails
    print(f"JARVIS: {text}")

    # Set flag to True before speaking
    is_speaking = True

    try:
        # On Windows, use native SAPI first because pyttsx3 can be silent on some setups.
        if sys.platform == "win32":
            escaped = text.replace("'", "''")
            ps = (
                "Add-Type -AssemblyName System.Speech; "
                "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                "$s.Rate = 0; "
                f"$s.Speak('{escaped}');"
            )
            win_tts = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps],
                capture_output=True,
                text=True,
            )
            if win_tts.returncode == 0:
                return

        # Cross-platform/default path.
        try:
            tts_engine = init_engine()
            if tts_engine:
                tts_engine.say(text)
                tts_engine.runAndWait()
            else:
                print("TTS engine not available")
        except Exception as e:
            print(f"TTS Error: {e}")
            # Try reinitializing engine on error
            try:
                global engine
                engine = None
                tts_engine = init_engine()
                if tts_engine:
                    tts_engine.say(text)
                    tts_engine.runAndWait()
            except Exception as e2:
                print(f"TTS Retry Error: {e2}")

    finally:
        # Ensure flag is reset to False even if errors occur
        is_speaking = False

def listen():
    global is_speaking
    # if system is speaking, don't listen
    if is_speaking:
        return "none"

    r = sr.Recognizer()
    r.dynamic_energy_threshold = True
    r.energy_threshold = 250
    r.pause_threshold = 0.9
    r.non_speaking_duration = 0.5
    r.operation_timeout = 8

    with sr.Microphone() as source:
        print("Listening...")
        try:
            r.adjust_for_ambient_noise(source, duration=1.0)
        except Exception:
            pass

        # Retry a couple of times to reduce missed commands in noisy environments
        for _ in range(2):
            try:
                audio = r.listen(source, timeout=6, phrase_time_limit=10)
                print("Recognizing...")

                # Ask for alternatives and pick the best hypothesis when available
                raw_result = r.recognize_google(audio, show_all=True)
                if isinstance(raw_result, dict) and raw_result.get("alternative"):
                    alternatives = raw_result.get("alternative", [])
                    best = alternatives[0].get("transcript", "")
                    best_conf = alternatives[0].get("confidence", 0.0)
                    for alt in alternatives:
                        conf = alt.get("confidence", 0.0)
                        if conf > best_conf:
                            best_conf = conf
                            best = alt.get("transcript", best)
                    if best:
                        return best.lower().strip()

                # Fallback path
                query = r.recognize_google(audio)
                return query.lower().strip()

            except sr.UnknownValueError:
                continue
            except sr.WaitTimeoutError:
                continue
            except Exception:
                continue

        return "none"
