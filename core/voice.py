import os
import sys
import re
import json
import audioop
import tempfile
import subprocess
from pathlib import Path
import pyttsx3
import speech_recognition as sr

# Initialize engine globally to avoid re-initialization issues
engine = None

ASR_KEYWORD_BOOST = {
    "pdf": 0.20,
    "file": 0.08,
    "document": 0.10,
    "folder": 0.08,
    "open": 0.05,
    "read": 0.05,
}

ASR_INTENT_HINTS = (
    "open",
    "file",
    "folder",
    "document",
    "read",
    "find",
    "search",
)

DEFAULT_PHRASE_HINTS = (
    "pdf",
    "docx",
    "resume",
    "telegram",
    "whatsapp",
    "youtube",
    "vlc",
    "movie",
    "video",
)

ASR_ADAPTATION_FILENAME = "asr_adaptation.json"
FILE_EXTENSIONS = "pdf|docx|txt|xlsx|xls|pptx|ppt|png|jpg|jpeg|gif|webp|mp3|mp4|mkv|avi|mov|zip|rar|7z"
DOT_MARKER_PATTERN = r"(?:dot|point|period|full\s*stop|short|shot|sort|dart)"
PRIMARY_ASR_LANGUAGE = (os.environ.get("JARVIS_PRIMARY_ASR_LANGUAGE", "en-IN") or "en-IN").strip()


def _asr_debug_enabled() -> bool:
    return os.environ.get("JARVIS_ASR_DEBUG", "false").lower() in {"1", "true", "yes", "on"}


def _asr_debug(msg: str) -> None:
    if _asr_debug_enabled():
        print(f"[ASR] {msg}")


def _env_int(name: str, default: int, minimum: int | None = None, maximum: int | None = None) -> int:
    raw = os.environ.get(name)
    try:
        value = int(raw) if raw is not None else int(default)
    except Exception:
        value = int(default)
    if minimum is not None:
        value = max(minimum, value)
    if maximum is not None:
        value = min(maximum, value)
    return value


def _env_float(name: str, default: float, minimum: float | None = None, maximum: float | None = None) -> float:
    raw = os.environ.get(name)
    try:
        value = float(raw) if raw is not None else float(default)
    except Exception:
        value = float(default)
    if minimum is not None:
        value = max(minimum, value)
    if maximum is not None:
        value = min(maximum, value)
    return value


def _get_phrase_hints() -> list[str]:
    raw = os.environ.get("JARVIS_PHRASE_HINTS", "")
    hints = [x.strip().lower() for x in raw.split(",") if x.strip()]
    if not hints:
        hints = list(DEFAULT_PHRASE_HINTS)

    adaptation = _load_asr_adaptation()
    learned_hints = adaptation.get("phrase_hints", []) if isinstance(adaptation, dict) else []
    for hint in learned_hints:
        clean = str(hint).strip().lower()
        if clean and clean not in hints:
            hints.append(clean)
    return hints


def _get_project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _get_asr_adaptation_path() -> Path:
    custom = os.environ.get("JARVIS_ASR_ADAPT_FILE", "").strip()
    if custom:
        return Path(os.path.expandvars(os.path.expanduser(custom)))
    return _get_project_root() / ASR_ADAPTATION_FILENAME


def _load_asr_adaptation() -> dict:
    path = _get_asr_adaptation_path()
    try:
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                data.setdefault("phrase_hints", [])
                data.setdefault("replacements", {})
                return data
    except Exception:
        pass
    return {"phrase_hints": [], "replacements": {}}


def _save_asr_adaptation(data: dict) -> None:
    path = _get_asr_adaptation_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=True, indent=2), encoding="utf-8")
    except Exception:
        pass


def _apply_adaptive_replacements(text: str) -> str:
    if not text:
        return text
    adaptation = _load_asr_adaptation()
    replacements = adaptation.get("replacements", {}) if isinstance(adaptation, dict) else {}
    if not isinstance(replacements, dict) or not replacements:
        return text

    out = text
    for wrong in sorted(replacements.keys(), key=len, reverse=True):
        right = str(replacements.get(wrong, "")).strip().lower()
        wrong_clean = str(wrong).strip().lower()
        if not wrong_clean or not right:
            continue
        out = re.sub(rf"\b{re.escape(wrong_clean)}\b", right, out)
    return out


def _learn_from_transcript(text: str, phrase_hints: list[str]) -> None:
    normalized = (text or "").strip().lower()
    if not normalized:
        return

    learned_tokens: set[str] = set()
    for match in re.finditer(rf"\b([a-z0-9][a-z0-9 _-]{{1,80}}\.(?:{FILE_EXTENSIONS}))\b", normalized):
        filename = match.group(1).strip().lower()
        stem = Path(filename).stem.lower().strip()
        if filename:
            learned_tokens.add(filename)
        if stem and len(stem) >= 3:
            learned_tokens.add(stem)
            for part in re.split(r"[_\-\s]+", stem):
                if len(part) >= 3:
                    learned_tokens.add(part)

    if not learned_tokens:
        return

    adaptation = _load_asr_adaptation()
    existing = adaptation.get("phrase_hints", []) if isinstance(adaptation, dict) else []
    if not isinstance(existing, list):
        existing = []
    merged = {str(x).strip().lower() for x in existing if str(x).strip()}
    merged.update(learned_tokens)
    merged.update(str(x).strip().lower() for x in phrase_hints if str(x).strip())

    # Keep adaptation file bounded and focused.
    adaptation["phrase_hints"] = sorted(merged)[:300]
    if "replacements" not in adaptation or not isinstance(adaptation["replacements"], dict):
        adaptation["replacements"] = {}
    _save_asr_adaptation(adaptation)


def _is_weak_signal(audio: sr.AudioData, min_rms_override: int | None = None) -> bool:
    """Reject very low-energy clips that are likely background noise only."""
    try:
        sample_width = getattr(audio, "sample_width", 2)
        rms = audioop.rms(audio.frame_data, sample_width)
        min_rms = min_rms_override if min_rms_override is not None else _env_int("JARVIS_MIN_RMS", 90, minimum=20, maximum=5000)
        return rms < min_rms
    except Exception:
        return False


def _is_probable_noise(audio: sr.AudioData) -> bool:
    """Reject clips with extreme zero-crossing ratio that often indicates hiss/static."""
    try:
        raw = audio.get_raw_data(convert_rate=16000, convert_width=2)
        sample_count = max(1, len(raw) // 2)
        zc = audioop.cross(raw, 2)
        zcr = zc / sample_count
        max_zcr = _env_float("JARVIS_MAX_ZCR", 0.62, minimum=0.20, maximum=0.9)
        rms = audioop.rms(raw, 2)
        if rms <= 0:
            return True
        peak = audioop.max(raw, 2)
        peak_to_rms = peak / max(rms, 1)
        max_peak_to_rms = _env_float("JARVIS_MAX_PEAK_TO_RMS", 18.0, minimum=3.0, maximum=45.0)
        return zcr > max_zcr and peak_to_rms > max_peak_to_rms
    except Exception:
        return False


def _is_low_quality_transcript(transcript: str) -> bool:
    text = (transcript or "").strip().lower()
    if not text:
        return True
    if len(text) <= 2:
        return True
    return bool(re.match(r"^(um+|uh+|hmm+|mm+|ah+|ok|okay|yes|yeah|no|hello|hi)$", text))


def _candidate_score(text: str, phrase_hints: list[str], confidence: float | None = None) -> float:
    lower = (text or "").lower().strip()
    if not lower:
        return -1.0

    conf = float(confidence) if confidence is not None else 0.55
    score = 0.10 + min(max(conf, 0.0), 1.0) * 0.65
    score += min(len(lower), 120) / 1100.0

    # Penalize transcripts that look like noisy syllables/repeats.
    if re.search(r"(.)\1{4,}", lower):
        score -= 0.18

    for keyword, boost in ASR_KEYWORD_BOOST.items():
        if keyword in lower:
            score += boost

    for hint in phrase_hints:
        if hint and hint in lower:
            score += 0.10

    # Prefer transcriptions with meaningful words over tiny fragments.
    word_count = len([w for w in lower.split() if w])
    score += min(word_count, 12) * 0.03
    return score


def _extract_google_candidates(raw_result) -> list[tuple[str, float | None]]:
    if not isinstance(raw_result, dict):
        return []
    alternatives = raw_result.get("alternative", [])
    if not alternatives:
        return []

    collected: list[tuple[str, float | None]] = []
    for alt in alternatives:
        transcript = (alt.get("transcript") or "").strip()
        if transcript:
            confidence = alt.get("confidence")
            try:
                confidence_val = float(confidence) if confidence is not None else None
            except Exception:
                confidence_val = None
            collected.append((transcript, confidence_val))
    return collected


def _select_best_candidate(candidates: list[tuple[str, float | None]], phrase_hints: list[str]) -> tuple[str, float, float | None]:
    if not candidates:
        return "", -1.0, None
    deduped: list[tuple[str, float | None]] = []
    seen: set[str] = set()
    for item, confidence in candidates:
        normalized = item.lower().strip()
        if normalized and normalized not in seen:
            seen.add(normalized)
            deduped.append((item, confidence))

    if not deduped:
        return "", -1.0, None

    best_text, best_confidence = max(
        deduped,
        key=lambda pair: _candidate_score(pair[0], phrase_hints, pair[1]),
    )
    best_score = _candidate_score(best_text, phrase_hints, best_confidence)
    return best_text, best_score, best_confidence


def _recognize_best_result(recognizer: sr.Recognizer, audio: sr.AudioData, language: str, phrase_hints: list[str]) -> tuple[str, float, float | None]:
    _asr_debug(f"Trying language={language}")
    try:
        raw_result = recognizer.recognize_google(audio, show_all=True, language=language)
    except Exception:
        raw_result = None

    candidates = _extract_google_candidates(raw_result)
    if not candidates:
        try:
            plain = recognizer.recognize_google(audio, language=language)
            if plain:
                candidates = [(plain, None)]
        except Exception:
            return "", -1.0, None

    best_candidate, best_score, best_confidence = _select_best_candidate(candidates, phrase_hints)
    if not best_candidate:
        _asr_debug(f"No candidate selected for language={language}")
        return "", -1.0, None

    _asr_debug(f"Candidate ({language}): {best_candidate} | score={best_score:.3f} conf={best_confidence}")
    return best_candidate, best_score, best_confidence


def _calibrate_ambient_noise(r: sr.Recognizer, source) -> None:
    rounds = _env_int("JARVIS_NOISE_CALIBRATION_ROUNDS", 1, minimum=1, maximum=6)
    duration = _env_float("JARVIS_NOISE_SAMPLE_SEC", 0.6, minimum=0.25, maximum=3.0)
    thresholds: list[int] = []
    for _ in range(rounds):
        try:
            r.adjust_for_ambient_noise(source, duration=duration)
            thresholds.append(int(r.energy_threshold))
        except Exception:
            continue

    if thresholds:
        thresholds.sort()
        median = thresholds[len(thresholds) // 2]
        r.energy_threshold = min(1200, max(100, int(median * 1.1)))

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
        lower = transcript.lower()
        for keyword, boost in ASR_KEYWORD_BOOST.items():
            if keyword in lower:
                score += boost
        if score > best_score:
            best_score = score
            best_text = transcript

    return best_text


def _normalize_recognition_text(text: str) -> str:
    """Fix frequent speech-recognition confusions for command phrases."""
    if not text:
        return ""

    normalized = text.lower().strip()

    # Convert spaced spelling to canonical forms.
    normalized = re.sub(r"\bp\s*d\s*f\b", "pdf", normalized)
    normalized = re.sub(r"\bdoc\s*x\b", "docx", normalized)
    normalized = re.sub(r"\bm\s*k\s*v\b", "mkv", normalized)
    normalized = re.sub(r"\bm\s*p\s*4\b", "mp4", normalized)
    normalized = re.sub(r"\btkv\b", "mkv", normalized)
    normalized = re.sub(r"\bdot\s+(mkv|mp4|avi|mov|wmv|webm|pdf|docx|txt)\b", r".\1", normalized)
    normalized = re.sub(rf"\b{DOT_MARKER_PATTERN}\s+({FILE_EXTENSIONS})\b", r".\1", normalized)

    # Handle phrases like "open resume short pdf" / "resume dot pdf".
    normalized = re.sub(
        rf"\b([a-z0-9][a-z0-9 _-]{{1,80}}?)\s+{DOT_MARKER_PATTERN}\s+({FILE_EXTENSIONS})\b",
        lambda m: f"{m.group(1).strip()}.{m.group(2)}",
        normalized,
    )

    # Keep dotted filenames tightly formatted.
    normalized = re.sub(r"\s*\.\s*", ".", normalized)

    # Common media-title confusion in ASR: cold <-> gold.
    if re.search(r"\b(play|open)\b", normalized):
        normalized = re.sub(r"\bgold\s+storage\b", "cold storage", normalized)
        normalized = re.sub(r"\bgoal\s+storage\b", "cold storage", normalized)

    has_file_intent = any(hint in normalized for hint in ASR_INTENT_HINTS)
    asks_for_document = any(w in normalized for w in (" file", "document", "doc ", ".pdf", "pdf "))
    mentions_media = any(w in normalized for w in ("movie", "song", "music", "youtube", "folder"))
    if has_file_intent and asks_for_document and not mentions_media:
        # Typical confusion: "pdf" can be misheard as "video" in noisy environments.
        normalized = re.sub(r"\bvideo\b", "pdf", normalized)
        normalized = re.sub(r"\bvideos\b", "pdfs", normalized)

    normalized = _apply_adaptive_replacements(normalized)

    return normalized


def _get_asr_languages() -> list[str]:
    raw = os.environ.get("JARVIS_ASR_LANGUAGES", "en-IN,en-US,ta-IN")
    langs = [x.strip() for x in raw.split(",") if x.strip()]
    return langs or ["en-IN", "en-US", "ta-IN"]


def _get_asr_language_order() -> list[str]:
    langs = _get_asr_languages()
    ordered: list[str] = []
    if PRIMARY_ASR_LANGUAGE and PRIMARY_ASR_LANGUAGE in langs:
        ordered.append(PRIMARY_ASR_LANGUAGE)

    for preferred in ("en-US", "ta-IN"):
        if preferred in langs and preferred not in ordered:
            ordered.append(preferred)

    for lang in langs:
        if lang not in ordered:
            ordered.append(lang)
    return ordered


def _detect_text_language(text: str) -> str:
    sample = (text or "").strip().lower()
    if not sample:
        return "en"

    if re.search(r"[\u0B80-\u0BFF]", sample):
        return "ta"

    tamil_romanized_markers = (
        "vanakkam", "tamizh", "tamil", "ungal", "ennai", "epadi", "enna", "nandri"
    )
    if any(marker in sample for marker in tamil_romanized_markers):
        return "ta"

    return "en"


def _looks_like_english_command_candidate(text: str) -> bool:
    t = (text or "").strip().lower()
    if not t:
        return False

    if not re.search(r"[a-z]", t):
        return False

    command_words = (
        "open", "download", "downloads", "folder", "file", "read", "play",
        "search", "find", "delete", "create", "launch", "desktop", "documents",
        "music", "video", "picture", "photos", "whatsapp", "youtube",
    )
    return any(w in t for w in command_words)


def _select_voice_for_language(tts_engine, language: str) -> None:
    try:
        voices = tts_engine.getProperty("voices")
    except Exception:
        return

    if language == "ta":
        for voice in voices:
            details = " ".join(
                [
                    str(getattr(voice, "name", "")),
                    str(getattr(voice, "id", "")),
                    str(getattr(voice, "languages", "")),
                ]
            ).lower()
            if "ta" in details or "tamil" in details:
                try:
                    tts_engine.setProperty("voice", voice.id)
                    return
                except Exception:
                    continue
    else:
        for voice in voices:
            name = str(getattr(voice, "name", "")).lower()
            if "male" in name or "david" in name or "mark" in name:
                try:
                    tts_engine.setProperty("voice", voice.id)
                    return
                except Exception:
                    continue


def _speak_with_gtts(text: str, language: str) -> bool:
    lang_code = "ta" if language == "ta" else "en"
    try:
        from gtts import gTTS
        from playsound import playsound
    except Exception:
        return False

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp:
            temp_path = tmp.name
        gTTS(text=text, lang=lang_code, slow=False).save(temp_path)
        playsound(temp_path)
        return True
    except Exception:
        return False
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

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


def _emit_runtime_status(status: str) -> None:
    try:
        from gui.app import set_runtime_status
        set_runtime_status(status)
    except Exception:
        pass


def speak(text, language: str | None = None):
    global is_speaking
    text = _clean_for_speech(text)
    lang = language or _detect_text_language(text)

    # Print first so user sees it even if audio fails
    print(f"JARVIS: {text}")
    _emit_runtime_status("Speaking...")

    # Set flag to True before speaking
    is_speaking = True

    try:
        # Use gTTS path for Tamil (cloud voice, works even when Tamil SAPI voice is unavailable).
        prefer_gtts = os.environ.get("JARVIS_USE_GTTS", "true").lower() in {"1", "true", "yes"}
        if prefer_gtts and lang == "ta":
            if _speak_with_gtts(text, lang):
                return

        # On Windows, use native SAPI first for English because pyttsx3 can be silent on some setups.
        if sys.platform == "win32" and lang == "en":
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
                _select_voice_for_language(tts_engine, lang)
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
                    _select_voice_for_language(tts_engine, lang)
                    tts_engine.say(text)
                    tts_engine.runAndWait()
            except Exception as e2:
                print(f"TTS Retry Error: {e2}")

    finally:
        # Ensure flag is reset to False even if errors occur
        is_speaking = False
        _emit_runtime_status("Ready")

def listen():
    global is_speaking
    # if system is speaking, don't listen
    if is_speaking:
        return "none"

    r = sr.Recognizer()
    r.dynamic_energy_threshold = True
    r.dynamic_energy_adjustment_damping = 0.18
    r.dynamic_energy_ratio = 1.15
    r.energy_threshold = _env_int("JARVIS_ENERGY_THRESHOLD", 210, minimum=80, maximum=1400)
    r.pause_threshold = 0.70
    r.phrase_threshold = 0.22
    r.non_speaking_duration = 0.30
    r.operation_timeout = 10

    mic_sample_rate = _env_int("JARVIS_MIC_SAMPLE_RATE", 0, minimum=0, maximum=48000)
    mic_chunk_size = _env_int("JARVIS_MIC_CHUNK_SIZE", 1024, minimum=256, maximum=4096)
    phrase_hints = _get_phrase_hints()

    mic_kwargs = {"chunk_size": mic_chunk_size}
    if mic_sample_rate > 0:
        mic_kwargs["sample_rate"] = mic_sample_rate

    with sr.Microphone(**mic_kwargs) as source:
        print("Listening...")
        _emit_runtime_status("Listening...")
        try:
            _calibrate_ambient_noise(r, source)
            # Keep threshold in a practical range after ambient calibration.
            r.energy_threshold = min(1200, max(100, int(r.energy_threshold)))
        except Exception:
            pass

        dynamic_min_rms = max(40, int(r.energy_threshold * 0.25))

        # Retry with progressively larger listen windows for noisy/slow speech.
        listen_profiles = [
            {"timeout": 5, "phrase_time_limit": 10},
            {"timeout": 7, "phrase_time_limit": 14},
            {"timeout": 9, "phrase_time_limit": 18},
        ]

        for profile in listen_profiles:
            try:
                audio = r.listen(
                    source,
                    timeout=profile["timeout"],
                    phrase_time_limit=profile["phrase_time_limit"],
                )
                print("Recognizing...")
                _emit_runtime_status("Recognizing...")

                if _is_weak_signal(audio, min_rms_override=dynamic_min_rms):
                    # Skip very low-energy clips to avoid random noisy transcriptions.
                    continue

                if _is_probable_noise(audio):
                    continue

                language_order = _get_asr_language_order()
                _asr_debug(f"Language order: {', '.join(language_order)}")
                evaluation_languages = list(language_order)
                if "en" not in evaluation_languages:
                    evaluation_languages.append("en")

                best_candidate = ""
                best_score = -1.0
                best_language = ""
                best_en_candidate = ""
                best_en_score = -1.0
                best_en_language = ""

                for lang in evaluation_languages:
                    candidate, score, _conf = _recognize_best_result(r, audio, lang, phrase_hints)
                    if candidate and score > best_score:
                        best_candidate = candidate
                        best_score = score
                        best_language = lang

                    if lang.lower().startswith("en") and candidate and score > best_en_score:
                        best_en_candidate = candidate
                        best_en_score = score
                        best_en_language = lang

                # Avoid Tamil transliteration false-positives for English file/system commands.
                if (
                    best_language.lower().startswith("ta")
                    and best_en_candidate
                    and _looks_like_english_command_candidate(best_en_candidate)
                    and best_en_score >= (best_score - 0.14)
                ):
                    _asr_debug(
                        f"Preferring English command candidate ({best_en_language}) over Tamil fallback"
                    )
                    best_candidate = best_en_candidate
                    best_score = best_en_score
                    best_language = best_en_language

                if best_candidate and not _is_low_quality_transcript(best_candidate) and best_score >= 0.24:
                    normalized = _normalize_recognition_text(best_candidate)
                    _asr_debug(f"Accepted ({best_language}) score={best_score:.3f}: {normalized}")
                    _learn_from_transcript(normalized, phrase_hints)
                    _emit_runtime_status("Ready")
                    return normalized

                # Optional offline fallback when PocketSphinx is installed.
                if os.environ.get("JARVIS_ENABLE_SPHINX_FALLBACK", "false").lower() in {"1", "true", "yes"}:
                    try:
                        offline = r.recognize_sphinx(audio)
                        if offline:
                            normalized = _normalize_recognition_text(offline)
                            _learn_from_transcript(normalized, phrase_hints)
                            _emit_runtime_status("Ready")
                            return normalized
                    except Exception:
                        pass

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
