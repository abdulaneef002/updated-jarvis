import os
import sys
import argparse
import threading 
import time
import json
import re
import difflib
from dotenv import load_dotenv
from core.voice import speak, listen
from core.registry import SkillRegistry
from core.engine import JarvisEngine
from gui.app import run_gui as run_gui_app, set_runtime_status

# Load Env
load_dotenv()

if not os.environ.get("GROQ_API_KEY"):
    print("Error: GROQ_API_KEY not found.")
    sys.exit(1)


def _contains_wake_word(text: str) -> bool:
    t = (text or "").strip().lower()
    if not t:
        return False

    normalized = re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", t)
    wake_variants = {
        "jarvis", "jarviss", "jarves", "jervis", "jarvice", "jarvish",
        "ஜார்விஸ்", "ஜார்விச்", "ஜார் இஸ்", "ஜர்விஸ்",
    }

    if normalized in {re.sub(r"\s+", "", w.lower()) for w in wake_variants}:
        return True

    # Token-level fuzzy match for ASR outputs like "jar vise" / "ஜார் இஸ்".
    tokens = re.findall(r"[a-z0-9\u0B80-\u0BFF]+", t)
    for token in tokens:
        token_norm = re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", token.lower())
        if not token_norm:
            continue
        for variant in wake_variants:
            v = re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", variant.lower())
            if token_norm == v:
                return True
            if difflib.SequenceMatcher(None, token_norm, v).ratio() >= 0.82:
                return True

    # Phrase-level fallback for spaced variants.
    for variant in wake_variants:
        compact_variant = re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", variant.lower())
        if compact_variant and compact_variant in normalized:
            return True

    return False


def _looks_like_wake_attempt(text: str) -> bool:
    t = (text or "").strip().lower()
    if not t:
        return False

    compact = re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", t)
    if not compact:
        return False

    # Trigger only for short utterances to avoid accidental activation mid-sentence.
    word_count = len(re.findall(r"[a-z0-9\u0B80-\u0BFF]+", t))
    if word_count > 3:
        return False

    candidates = [
        "jarvis", "jarviss", "jarves", "jervis", "jarvice", "jaris", "jarvish",
        "ஜார்விஸ்", "ஜார்விச்", "ஜர்விஸ்",
    ]
    compact_candidates = [re.sub(r"[^a-z0-9\u0B80-\u0BFF]+", "", c.lower()) for c in candidates]

    best_ratio = 0.0
    for cand in compact_candidates:
        if not cand:
            continue
        ratio = difflib.SequenceMatcher(None, compact, cand).ratio()
        if ratio > best_ratio:
            best_ratio = ratio

    return best_ratio >= 0.74

def jarvis_loop(pause_event, registry, args):
    """
    Main loop for JARVIS, running in a separate thread.
    Checks pause_event to determine if it should listen/process.
    """
    # Initialize Engine
    jarvis = JarvisEngine(registry)
    voice_active = bool(args.text)

    if args.text:
        speak("Jarvis Online. Ready for your command.")
        set_runtime_status("Ready")
    else:
        print("JARVIS: Standby mode. Say 'jarvis' to activate.")
        set_runtime_status("Standby (say jarvis)")

    while True:
        # Check for pause
        if pause_event.is_set():
            time.sleep(0.5)
            continue

        if args.text:
            try:
                user_query = input("YOU: ")
            except EOFError:
                break
        else:
            set_runtime_status("Listening...")
            user_query = listen()
            
        # Double check pause after listening (in case paused during listen)
        if pause_event.is_set():
            continue

        normalized_query = user_query.lower().strip()

        if normalized_query == "none" or not normalized_query: continue
        if "quit" in normalized_query: 
            print("Shutting down JARVIS loop...")
            # We can't easily kill the main thread (GUI) from here, 
            # but we can stop this loop. The user will have to close the GUI.
            speak("Shutting down.")
            break
        
        if args.text:
            clean_query = normalized_query
        else:
            if not voice_active:
                if not _contains_wake_word(normalized_query) and not _looks_like_wake_attempt(normalized_query):
                    print(f"Ignored (waiting for wake word): {user_query}")
                    set_runtime_status("Standby (say jarvis)")
                    continue

                voice_active = True
                speak("Jarvis is ready for you.")
                set_runtime_status("Ready")
                clean_query = re.sub(r"\bjar\s*vis\b|\bjarvis\b|ஜார்விஸ்|ஜார்\s*இஸ்", "", normalized_query, flags=re.IGNORECASE).strip()
                if not clean_query:
                    continue
            else:
                clean_query = re.sub(r"\bjar\s*vis\b|\bjarvis\b|ஜார்விஸ்|ஜார்\s*இஸ்", "", normalized_query, flags=re.IGNORECASE).strip()
                if not clean_query:
                    continue
        
        try:
            print(f"Thinking: {clean_query}")
            set_runtime_status("Thinking...")
            response = jarvis.run_conversation(clean_query)
            
            # Check pause before speaking response
            if pause_event.is_set():
                continue

            if response:
                spoken_response = response
                try:
                    parsed = json.loads(response)
                    spoken_response = parsed.get("message", response)
                except Exception:
                    pass

                # Always speak, and speak() also prints the same line so user gets text + voice.
                speak(spoken_response)
                set_runtime_status("Ready")
            else:
                speak("I heard you, but I could not generate a response. Please try again.")
                set_runtime_status("Ready")
        except Exception as e:
            print(f"Main Loop Error: {e}")
            speak("System error. Please try again.")
            set_runtime_status("Ready")

def main():
    parser = argparse.ArgumentParser(description="JARVIS AI Assistant")
    parser.add_argument("--text", action="store_true", help="Run in text mode (no voice I/O)")
    args = parser.parse_args()

    # 1. Setup Pause Event
    # Event is SET when PAUSED, CLEARED when RUNNING
    pause_event = threading.Event()
    context = {"pause_event": pause_event}

    # 2. Initialize Registry and Load Skills
    registry = SkillRegistry()
    skills_dir = os.path.join(os.path.dirname(__file__), "skills")
    registry.load_skills(skills_dir, context=context)
    
    # 3. Start JARVIS Loop in Background Thread
    # Daemon thread so it dies when GUI closes
    t = threading.Thread(target=jarvis_loop, args=(pause_event, registry, args), daemon=True)
    t.start()
    
    # 4. Start GUI in Main Thread (Required for PyQt)
    # This will block until the window is closed
    run_gui_app(pause_event)

if __name__ == "__main__":
    main()