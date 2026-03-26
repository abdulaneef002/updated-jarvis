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


def _extract_best_transcript(raw_result) -> str:
    """Pick the most confident transcript from recognize_google(show_all=True)."""
    if not isinstance(raw_result, dict):
        return ""

    alternatives = raw_result.get("alternative", [])
    if not alternatives:
        return ""

    best_text = ""
    best_score = -1.0
    for alt in alternatives:
        transcript = (alt.get("transcript") or "").strip()
        if not transcript:
            continue
        confidence = alt.get("confidence")
        if confidence is None:
            confidence = 0.55
        # Slightly favor longer candidates when confidence is tied.
        score = float(confidence) + (min(len(transcript), 80) / 1000.0)
        if score > best_score:
            best_score = score
            best_text = transcript

    return best_text

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
    r.dynamic_energy_adjustment_damping = 0.2
    r.dynamic_energy_ratio = 1.7
    r.energy_threshold = int(os.environ.get("JARVIS_ENERGY_THRESHOLD", "220"))
    r.pause_threshold = 0.75
    r.phrase_threshold = 0.25
    r.non_speaking_duration = 0.35
    r.operation_timeout = 10

    with sr.Microphone() as source:
        print("Listening...")
        try:
            r.adjust_for_ambient_noise(source, duration=1.2)
        except Exception:
            pass

        # Retry with progressively larger listen windows for noisy/slow speech.
        listen_profiles = [
            {"timeout": 5, "phrase_time_limit": 8},
            {"timeout": 7, "phrase_time_limit": 12},
            {"timeout": 8, "phrase_time_limit": 15},
        ]

        for profile in listen_profiles:
            try:
                audio = r.listen(
                    source,
                    timeout=profile["timeout"],
                    phrase_time_limit=profile["phrase_time_limit"],
                )
                print("Recognizing...")

                raw_result = r.recognize_google(audio, show_all=True)
                best = _extract_best_transcript(raw_result)
                if best:
                    return best.lower().strip()

                fallback = r.recognize_google(audio)
                if fallback:
                    return fallback.lower().strip()

            except sr.UnknownValueError:
                # Reduce threshold slightly after misses to catch softer speech.
                r.energy_threshold = max(120, int(r.energy_threshold * 0.9))
                continue
            except sr.WaitTimeoutError:
                continue
            except sr.RequestError:
                continue
            except Exception:
                continue

        return "none"
