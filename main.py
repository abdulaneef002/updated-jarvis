import os
import sys
import argparse
import threading 
import time
import json
from dotenv import load_dotenv
from core.voice import speak, listen
from core.registry import SkillRegistry
from core.engine import JarvisEngine
from gui.app import run_gui as run_gui_app

# Load Env
load_dotenv()

if not os.environ.get("GROQ_API_KEY"):
    print("Error: GROQ_API_KEY not found.")
    sys.exit(1)

def jarvis_loop(pause_event, registry, args):
    """
    Main loop for JARVIS, running in a separate thread.
    Checks pause_event to determine if it should listen/process.
    """
    # Initialize Engine
    jarvis = JarvisEngine(registry)

    speak("Jarvis Online. Ready for your command.")

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
            always_listen = os.environ.get("ALWAYS_LISTEN", "true").lower() in {"1", "true", "yes"}
            if always_listen:
                clean_query = normalized_query
            else:
                # Wake word / Command filtering Logic (voice mode only)
                direct_commands = [
                    "open", "volume", "search", "create", "write", "read", "make", "delete",
                    "shutdown", "restart", "rename", "move", "copy", "run", "execute",
                    "brightness", "wifi", "bluetooth",
                    "who", "what", "when", "where", "how", "why", "thank", "hello"
                ]
                
                is_direct = any(cmd in normalized_query for cmd in direct_commands)
                
                if "jarvis" not in normalized_query and not is_direct:
                    print(f"Ignored: {user_query}")
                    continue
                    
                clean_query = normalized_query.replace("jarvis", "").strip()
        
        try:
            print(f"Thinking: {clean_query}")
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
            else:
                speak("I heard you, but I could not generate a response. Please try again.")
        except Exception as e:
            print(f"Main Loop Error: {e}")
            speak("System error. Please try again.")

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