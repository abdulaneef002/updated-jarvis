import os
import sys
import argparse
import threading 
import time
import json
import re
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


def _ready_prompt(language: str) -> str:
    return "உங்கள் கட்டளைக்காக தயார்." if language == "ta" else "Ready for your command."


def _error_prompt(language: str) -> str:
    return "மன்னிக்கவும், ஒரு பிழை ஏற்பட்டது. மீண்டும் முயற்சிக்கவும்." if language == "ta" else "System error. Please try again."


def _no_response_prompt(language: str) -> str:
    return "நான் கேட்டேன், ஆனால் பதில் உருவாக்க முடியவில்லை. தயவுசெய்து மீண்டும் முயற்சிக்கவும்." if language == "ta" else "I heard you, but I could not generate a response. Please try again."

def jarvis_loop(pause_event, registry, args):
    """
    Main loop for JARVIS, running in a separate thread.
    Checks pause_event to determine if it should listen/process.
    """
    # Initialize Engine
    jarvis = JarvisEngine(registry)

    if args.text:
        speak("Jarvis Online. Ready for your command.")
        set_runtime_status("Ready")
    else:
        speak("Jarvis Online. Listening for your commands.")
        set_runtime_status("Ready")

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
            quit_lang = getattr(jarvis, "last_user_language", "en")
            speak("நிறுத்துகிறேன்." if quit_lang == "ta" else "Shutting down.", language=quit_lang)
            break
        
        if args.text:
            clean_query = normalized_query
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

                reply_lang = getattr(jarvis, "last_user_language", "en")

                # Always speak, and speak() also prints the same line so user gets text + voice.
                speak(spoken_response, language=reply_lang)
                if not args.text:
                    speak(_ready_prompt(reply_lang), language=reply_lang)
                set_runtime_status("Ready")
            else:
                reply_lang = getattr(jarvis, "last_user_language", "en")
                speak(_no_response_prompt(reply_lang), language=reply_lang)
                if not args.text:
                    speak(_ready_prompt(reply_lang), language=reply_lang)
                set_runtime_status("Ready")
        except Exception as e:
            print(f"Main Loop Error: {e}")
            reply_lang = getattr(jarvis, "last_user_language", "en")
            speak(_error_prompt(reply_lang), language=reply_lang)
            if not args.text:
                speak(_ready_prompt(reply_lang), language=reply_lang)
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