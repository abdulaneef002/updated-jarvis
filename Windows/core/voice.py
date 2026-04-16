import os
import sys
import pyttsx3
import speech_recognition as sr

# Initialize engine globally to avoid re-initialization issues
engine = pyttsx3.init()

# Set voice to deep male voice
def set_deep_male_voice():
    voices = engine.getProperty('voices')
    for voice in voices:
        # Prefer "Daniel" for deep male voice on Mac
        if "Daniel" in voice.name:
            engine.setProperty('voice', voice.id)
            return
    # Fallback to any male voice if Daniel not found
    for voice in voices:
        if "male" in voice.name.lower() or "male" in str(voice.gender).lower():
             engine.setProperty('voice', voice.id)
             return

set_deep_male_voice()


def _asr_debug_enabled() -> bool:
    return os.environ.get("JARVIS_ASR_DEBUG", "true").lower() in {"1", "true", "yes", "on"}


def _asr_debug(msg: str) -> None:
    if _asr_debug_enabled():
        print(f"[ASR] {msg}")


def _score_candidate(text: str, confidence: float | None = None) -> float:
    t = (text or "").strip().lower()
    if not t:
        return -1.0
    conf = float(confidence) if confidence is not None else 0.55
    score = 0.1 + min(max(conf, 0.0), 1.0) * 0.7
    score += min(len(t), 120) / 1300.0
    word_count = len([w for w in t.split() if w])
    score += min(word_count, 12) * 0.02
    return score


def _recognize_best_for_language(recognizer: sr.Recognizer, audio: sr.AudioData, language: str) -> tuple[str, float]:
    _asr_debug(f"Trying language={language}")
    best_text = ""
    best_score = -1.0

    try:
        raw_result = recognizer.recognize_google(audio, show_all=True, language=language)
        alternatives = raw_result.get("alternative", []) if isinstance(raw_result, dict) else []
        for alt in alternatives:
            transcript = (alt.get("transcript") or "").strip().lower()
            if not transcript:
                continue
            confidence = alt.get("confidence")
            try:
                conf_val = float(confidence) if confidence is not None else None
            except Exception:
                conf_val = None
            score = _score_candidate(transcript, conf_val)
            if score > best_score:
                best_text = transcript
                best_score = score
    except Exception:
        pass

    if not best_text:
        try:
            plain = recognizer.recognize_google(audio, language=language)
            plain = (plain or "").strip().lower()
            if plain:
                best_text = plain
                best_score = _score_candidate(plain, None)
        except Exception:
            pass

    if best_text:
        _asr_debug(f"Candidate ({language}): {best_text} | score={best_score:.3f}")
    return best_text, best_score

def speak(text):
    if "{" in text and "}" in text and "status" in text:
        text = "Task completed."
    
    # Print first so user sees it even if audio fails
    print(f"JARVIS: {text}")

    # Use pyttsx3 for Windows (SAPI5)
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"TTS Error: {e}")

def listen():
    r = sr.Recognizer()
    r.dynamic_energy_threshold = True
    r.dynamic_energy_adjustment_damping = 0.18
    r.dynamic_energy_ratio = 1.15
    r.energy_threshold = 210
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.7
        r.phrase_threshold = 0.22
        r.non_speaking_duration = 0.3
        r.adjust_for_ambient_noise(source)

        listen_profiles = [
            {"timeout": 5, "phrase_time_limit": 10},
            {"timeout": 7, "phrase_time_limit": 14},
            {"timeout": 9, "phrase_time_limit": 18},
        ]

        for profile in listen_profiles:
            try:
                audio = r.listen(source, timeout=profile["timeout"], phrase_time_limit=profile["phrase_time_limit"])
                print("Recognizing...")
                best_text = ""
                best_score = -1.0
                for language in ("en-IN", "en-US", "ta-IN", "en"):
                    candidate, score = _recognize_best_for_language(r, audio, language)
                    if candidate and score > best_score:
                        best_text = candidate
                        best_score = score

                if best_text:
                    _asr_debug(f"Accepted best candidate: {best_text}")
                    return best_text
            except Exception:
                continue

        return "none"
