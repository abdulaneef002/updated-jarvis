import os
import re
import json
import difflib
import shutil
import subprocess
import webbrowser
import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional


class SystemController:
    def __init__(self) -> None:
        self.os_name = self._detect_os()
        self.pending_action: Optional[Dict[str, Any]] = None
        self.pending_file_choices: Optional[Dict[str, Any]] = None

    def _normalize_folder_alias(self, folder_name: str) -> str:
        text = (folder_name or "").strip().lower()
        if not text:
            return ""

        text = re.sub(r"^the\s+", "", text)
        text = re.sub(r"\s+folder$", "", text)
        text = re.sub(r"\s+directory$", "", text)
        text = text.strip()

        alias_map = {
            "download": "downloads",
            "document": "documents",
            "video": "videos",
            "picture": "pictures",
            "photo": "pictures",
            "desktops": "desktop",
            "homes": "home",
            "my downloads": "downloads",
            "my documents": "documents",
            "my videos": "videos",
            "my music": "music",
            "my pictures": "pictures",
            "my photos": "pictures",
        }
        return alias_map.get(text, text)

    def _detect_os(self) -> str:
        if os.name == "nt":
            return "windows"
        if os.uname().sysname.lower() == "darwin":
            return "macos"
        return "linux"

    def handle_command(self, command_text: str) -> Dict[str, Any]:
        command = (command_text or "").strip()
        if not command:
            return self._response("unknown", "none", "failed", "No command received.")

        if self.pending_action:
            return self._handle_confirmation(command)

        if self.pending_file_choices:
            pending_choice_result = self._handle_pending_file_choice(command)
            if pending_choice_result:
                return pending_choice_result

        parsed = self._parse_command(command)
        if parsed["status"] == "failed":
            return parsed

        if parsed.get("dangerous", False):
            self.pending_action = parsed
            return self._response(
                parsed["intent"],
                "pending_confirmation",
                "confirmation_required",
                parsed["confirmation_message"],
            )

        return self._execute(parsed)

    def _handle_pending_file_choice(self, command: str) -> Optional[Dict[str, Any]]:
        if not self.pending_file_choices:
            return None

        normalized = command.lower().strip()
        if normalized in {"cancel", "stop", "nevermind", "no"}:
            self.pending_file_choices = None
            return self._response("open_file", "cancelled", "failed", "Selection cancelled.")

        idx = self._parse_choice_index(normalized)
        if idx is None:
            # If user started a fresh command, drop stale pending choices and continue normal parse flow.
            if re.match(r"^(open|launch|play|search|find|create|delete|rename|move|copy|run|execute)\b", normalized):
                self.pending_file_choices = None
                return None

            choices = self.pending_file_choices.get("matches", [])
            limit = len(choices)
            return self._response(
                "open_file",
                "clarification_required",
                "failed",
                f"Please say open first one, open second one, or open number 1 to {limit}.",
            )

        choices = self.pending_file_choices.get("matches", [])
        if not choices:
            self.pending_file_choices = None
            return self._response("open_file", "open_file", "failed", "No pending file choices found.")

        if idx < 1 or idx > len(choices):
            return self._response(
                "open_file",
                "clarification_required",
                "failed",
                f"Please choose a number between 1 and {len(choices)}.",
            )

        target = choices[idx - 1]
        self.pending_file_choices = None
        try:
            os.startfile(str(target))
            return self._response("open_file", "open_file", "success", f"Opened {target.name} from location {idx}.")
        except Exception as exc:
            return self._response("open_file", "open_file", "failed", f"Found {target.name} but could not open it: {exc}")

    def _parse_choice_index(self, normalized: str) -> Optional[int]:
        number_match = re.search(r"(?:number\s+)?(\d{1,2})", normalized)
        if number_match and re.search(r"\b(open|choose|select|pick|play|launch)\b", normalized):
            return int(number_match.group(1))

        ordinal_map = {
            "first": 1,
            "second": 2,
            "third": 3,
            "fourth": 4,
            "fifth": 5,
            "sixth": 6,
            "seventh": 7,
            "eighth": 8,
            "ninth": 9,
            "tenth": 10,
        }
        for word, idx in ordinal_map.items():
            if re.search(rf"\b{word}\b", normalized) and re.search(r"\b(open|choose|select|pick|play|launch)\b", normalized):
                return idx
        return None

    def _handle_confirmation(self, command: str) -> Dict[str, Any]:
        normalized = command.lower().strip()
        if normalized in {"yes", "y", "confirm", "proceed", "do it"}:
            parsed = self.pending_action
            self.pending_action = None
            if not parsed:
                return self._response("unknown", "none", "failed", "No pending action found.")
            return self._execute(parsed)

        if normalized in {"no", "n", "cancel", "stop"}:
            cancelled_intent = self.pending_action["intent"] if self.pending_action else "unknown"
            self.pending_action = None
            return self._response(cancelled_intent, "cancelled", "failed", "Action cancelled.")

        return self._response(
            self.pending_action.get("intent", "unknown"),
            "pending_confirmation",
            "confirmation_required",
            "Please reply with 'yes' to confirm or 'no' to cancel.",
        )

    def _parse_command(self, command: str) -> Dict[str, Any]:
        lowered = command.lower()

        open_file_in_folder = re.match(
            r"^(?:open|play|launch)\s+(.+?)\s+(?:in|from|inside)\s+(?:the\s+)?(.+?)\s+folder$",
            command,
            flags=re.IGNORECASE,
        )
        if open_file_in_folder:
            file_query = open_file_in_folder.group(1).strip().strip("\"'")
            folder_name = open_file_in_folder.group(2).strip().strip("\"'")
            return {
                "intent": "open_file_in_folder",
                "action": "open_file_in_folder",
                "status": "success",
                "dangerous": False,
                "params": {"query": file_query, "folder_name": folder_name},
            }

        open_folder_then_file = re.match(
            r"^(?:open|go\s+to)\s+(?:the\s+)?(.+?)\s+folder\s+and\s+(?:open|play)\s+(.+)$",
            command,
            flags=re.IGNORECASE,
        )
        if open_folder_then_file:
            folder_name = open_folder_then_file.group(1).strip().strip("\"'")
            file_query = open_folder_then_file.group(2).strip().strip("\"'")
            return {
                "intent": "open_file_in_folder",
                "action": "open_file_in_folder",
                "status": "success",
                "dangerous": False,
                "params": {"query": file_query, "folder_name": folder_name},
            }

        play_media_from_folder = re.match(
            r"^(?:play|open)\s+(?:a\s+)?(?:movie|video|song|audio)\s+(?:in|from|inside)\s+(?:the\s+)?(.+?)\s+folder$",
            command,
            flags=re.IGNORECASE,
        )
        if play_media_from_folder:
            folder_name = play_media_from_folder.group(1).strip().strip("\"'")
            return {
                "intent": "open_file_in_folder",
                "action": "open_file_in_folder",
                "status": "success",
                "dangerous": False,
                "params": {"query": "movie", "folder_name": folder_name},
            }

        play_match = re.match(r"^(?:play|open\s+and\s+play)\s+(.+)$", command, flags=re.IGNORECASE)
        if play_match:
            target = play_match.group(1).strip().strip("\"'")
            if "telegram desktop" in target.lower() or "telegram" in target.lower():
                return {
                    "intent": "play_media_in_folder",
                    "action": "play_media_in_folder",
                    "status": "success",
                    "dangerous": False,
                    "params": {"folder_hint": "telegram desktop", "query": ""},
                }
            return {
                "intent": "open_file",
                "action": "open_file",
                "status": "success",
                "dangerous": False,
                "params": {"query": target, "search_mode": "media"},
            }

        # Handle common ASR typo where "open" is recognized as "pen"
        if lowered.startswith("pen "):
            command = "open " + command[4:]
            lowered = command.lower()

        youtube_play_match = re.match(r"^(?:youtube\s+and\s+play|play\s+on\s+youtube)\s+(.+)$", command, flags=re.IGNORECASE)
        if youtube_play_match:
            query = youtube_play_match.group(1).strip().strip("\"'")
            url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
            return {
                "intent": "open_website",
                "action": "open_website",
                "status": "success",
                "dangerous": False,
                "params": {"url": url},
            }

        # Conversational quick replies
        if re.search(r"\b(can\s+you\s+hear\s+me|do\s+you\s+hear\s+me)\b", lowered):
            return {
                "intent": "conversation_reply",
                "action": "conversation_reply",
                "status": "success",
                "dangerous": False,
                "params": {"message": "Yes, I can hear you clearly."},
            }

        if re.search(r"\b(can\s+you\s+speak|do\s+you\s+speak|can\s+you\s+talk|do\s+you\s+talk)\b", lowered):
            return {
                "intent": "conversation_reply",
                "action": "conversation_reply",
                "status": "success",
                "dangerous": False,
                "params": {"message": "Yes, I can speak. I am speaking with you now."},
            }

        if re.search(r"\b(are\s+you\s+there|jarvis\s+are\s+you\s+there)\b", lowered):
            return {
                "intent": "conversation_reply",
                "action": "conversation_reply",
                "status": "success",
                "dangerous": False,
                "params": {"message": "Yes, I am here and ready."},
            }

        learn_correction = re.match(
            r"^(?:learn|remember)\s+correction\s*:?\s*(.+?)\s+(?:means|as)\s+(.+)$",
            command,
            flags=re.IGNORECASE,
        )
        if learn_correction:
            wrong_phrase = learn_correction.group(1).strip().strip("\"'")
            right_phrase = learn_correction.group(2).strip().strip("\"'")
            if wrong_phrase and right_phrase:
                return {
                    "intent": "learn_asr_correction",
                    "action": "learn_asr_correction",
                    "status": "success",
                    "dangerous": False,
                    "params": {"wrong": wrong_phrase, "right": right_phrase},
                }

        forget_correction = re.match(
            r"^(?:forget|remove|delete)\s+correction\s*:?\s*(.+)$",
            command,
            flags=re.IGNORECASE,
        )
        if forget_correction:
            wrong_phrase = forget_correction.group(1).strip().strip("\"'")
            if wrong_phrase:
                return {
                    "intent": "forget_asr_correction",
                    "action": "forget_asr_correction",
                    "status": "success",
                    "dangerous": False,
                    "params": {"wrong": wrong_phrase},
                }

        shell_match = re.match(r"^(?:run(?:\s+command)?|execute)\s+(.+)$", command, flags=re.IGNORECASE)
        if shell_match:
            shell_command = shell_match.group(1).strip()
            if not shell_command:
                return self._response(
                    "run_system_command",
                    "clarification_required",
                    "failed",
                    "Please provide the command to run.",
                )

            if not self._is_allowed_shell_command(shell_command):
                return self._response(
                    "run_system_command",
                    "blocked",
                    "failed",
                    "Command is outside the safe allowlist.",
                )

            is_risky = self._is_risky_shell_command(shell_command)
            parsed = {
                "intent": "run_system_command",
                "action": "run_system_command",
                "status": "success",
                "dangerous": is_risky,
                "params": {"command": shell_command},
            }
            if is_risky:
                parsed["confirmation_message"] = f"This command may change system state. Confirm execution: {shell_command}?"
            return parsed

        volume_match = re.search(r"(?:set\s+)?volume\s+(?:to\s+)?(\d{1,3})", lowered)
        if volume_match:
            return {
                "intent": "set_volume",
                "action": "set_volume",
                "status": "success",
                "dangerous": True,
                "confirmation_message": "This changes system configuration. Confirm volume update?",
                "params": {"level": int(volume_match.group(1))},
            }

        brightness_match = re.search(r"(?:set\s+)?brightness\s+(?:to\s+)?(\d{1,3})", lowered)
        if brightness_match:
            return {
                "intent": "set_brightness",
                "action": "set_brightness",
                "status": "success",
                "dangerous": True,
                "confirmation_message": "This changes system configuration. Confirm brightness update?",
                "params": {"level": int(brightness_match.group(1))},
            }

        wifi_on = re.search(r"(?:turn|switch)\s+(?:on|enable)\s+(?:the\s+)?wifi", lowered)
        wifi_off = re.search(r"(?:turn|switch)\s+(?:off|disable)\s+(?:the\s+)?wifi", lowered)
        if wifi_on or wifi_off:
            enable = wifi_on is not None
            return {
                "intent": "toggle_wifi",
                "action": "toggle_wifi",
                "status": "success",
                "dangerous": True,
                "confirmation_message": f"This changes system configuration. Confirm Wi-Fi {'enable' if enable else 'disable'}?",
                "params": {"enable": enable},
            }

        bt_on = re.search(r"(?:turn|switch)\s+(?:on|enable)\s+(?:the\s+)?bluetooth", lowered)
        bt_off = re.search(r"(?:turn|switch)\s+(?:off|disable)\s+(?:the\s+)?bluetooth", lowered)
        if bt_on or bt_off:
            enable = bt_on is not None
            return {
                "intent": "toggle_bluetooth",
                "action": "toggle_bluetooth",
                "status": "success",
                "dangerous": True,
                "confirmation_message": f"This changes system configuration. Confirm Bluetooth {'enable' if enable else 'disable'}?",
                "params": {"enable": enable},
            }

        rename_match = re.search(r"rename\s+file\s+(.+?)\s+to\s+(.+)$", command, flags=re.IGNORECASE)
        if rename_match:
            source_name = rename_match.group(1).strip().strip("\"'")
            target_name = rename_match.group(2).strip().strip("\"'")
            return {
                "intent": "rename_file",
                "action": "rename_file",
                "status": "success",
                "dangerous": False,
                "params": {"source_name": source_name, "target_name": target_name},
            }

        move_match = re.search(r"move\s+file\s+(.+?)\s+to\s+(.+)$", command, flags=re.IGNORECASE)
        if move_match:
            source_name = move_match.group(1).strip().strip("\"'")
            destination = move_match.group(2).strip().strip("\"'")
            return {
                "intent": "move_file",
                "action": "move_file",
                "status": "success",
                "dangerous": False,
                "params": {"source_name": source_name, "destination": destination},
            }

        copy_match = re.search(r"copy\s+file\s+(.+?)\s+to\s+(.+)$", command, flags=re.IGNORECASE)
        if copy_match:
            source_name = copy_match.group(1).strip().strip("\"'")
            destination = copy_match.group(2).strip().strip("\"'")
            return {
                "intent": "copy_file",
                "action": "copy_file",
                "status": "success",
                "dangerous": False,
                "params": {"source_name": source_name, "destination": destination},
            }

        open_match = re.match(r"^(open|launch)\s+(.+)$", command, flags=re.IGNORECASE)
        if open_match:
            target = open_match.group(2).strip()
            target_lower = target.lower()

            open_file_phrase = re.match(r"^file\s+(.+)$", target, flags=re.IGNORECASE)
            if open_file_phrase:
                explicit_file_query = open_file_phrase.group(1).strip().strip("\"'")
                return {
                    "intent": "open_file",
                    "action": "open_file",
                    "status": "success",
                    "dangerous": False,
                    "params": {"query": explicit_file_query},
                }

            # Support folder opening commands like "open desktop folder"
            folder_open_match = re.match(r"^(desktop|documents|document|downloads|download|videos|video|music|pictures|picture|photos|photo)\s+folder$", target_lower)
            if folder_open_match:
                return {
                    "intent": "open_folder",
                    "action": "open_folder",
                    "status": "success",
                    "dangerous": False,
                    "params": {"folder_name": self._normalize_folder_alias(folder_open_match.group(1))},
                }

            # Clean location qualifiers from app commands (e.g. "whatsapp in desktop")
            target_lower = re.sub(r"\s+(?:in|on|from)\s+(?:the\s+)?(?:desktop|home|home page)$", "", target_lower).strip()
            target = target_lower

            # Example: "open telegram desktop and play movie"
            if "telegram desktop" in target_lower and re.search(r"\b(play|movie|video)\b", target_lower):
                return {
                    "intent": "play_media_in_folder",
                    "action": "play_media_in_folder",
                    "status": "success",
                    "dangerous": False,
                    "params": {"folder_hint": "telegram desktop", "query": ""},
                }
            
            if target_lower in {"whatsapp web", "web whatsapp"}:
                return {
                    "intent": "open_website",
                    "action": "open_website",
                    "status": "success",
                    "dangerous": False,
                    "params": {"url": "https://web.whatsapp.com"},
                }

            if target_lower == "whatsapp" or "whatsapp" in target_lower:
                return {
                    "intent": "open_application",
                    "action": "open_application",
                    "status": "success",
                    "dangerous": False,
                    "params": {"app_name": "WhatsApp"},
                }
            
            # Common websites/apps with correct known URLs
            KNOWN_SITE_URLS = {
                "youtube": "https://www.youtube.com",
                "google": "https://www.google.com",
                "gmail": "https://mail.google.com",
                "google mail": "https://mail.google.com",
                "facebook": "https://www.facebook.com",
                "instagram": "https://www.instagram.com",
                "twitter": "https://www.twitter.com",
                "x": "https://www.x.com",
                "linkedin": "https://www.linkedin.com",
                "github": "https://www.github.com",
                "reddit": "https://www.reddit.com",
                "stackoverflow": "https://www.stackoverflow.com",
                "amazon": "https://www.amazon.com",
                "wikipedia": "https://www.wikipedia.org",
                "netflix": "https://www.netflix.com",
                "spotify": "https://open.spotify.com",
                "twitch": "https://www.twitch.tv",
                "discord": "https://discord.com/app",
                "zoom": "https://zoom.us",
                "chatgpt": "https://chat.openai.com",
            }

            # Apps that prefer desktop installation but can fall back to web
            DESKTOP_FIRST_APPS = {
                "telegram": {
                    "paths": [
                        Path(os.environ.get("APPDATA", "")) / "Telegram Desktop" / "Telegram.exe",
                        Path(os.environ.get("LOCALAPPDATA", "")) / "Telegram Desktop" / "Telegram.exe",
                    ],
                    "web_fallback": "https://web.telegram.org",
                },
            }

            if target_lower in DESKTOP_FIRST_APPS or "telegram" in target_lower:
                # Route to open_application so _open_application can handle desktop-first logic
                return {
                    "intent": "open_application",
                    "action": "open_application",
                    "status": "success",
                    "dangerous": False,
                    "params": {"app_name": "telegram" if "telegram" in target_lower else target},
                }

            # Handle natural commands like "open youtube and play <song>"
            yt_inline_play = re.match(r"^youtube\s+and\s+play\s+(.+)$", target_lower)
            if yt_inline_play:
                query = yt_inline_play.group(1).strip()
                url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
                return {
                    "intent": "open_website",
                    "action": "open_website",
                    "status": "success",
                    "dangerous": False,
                    "params": {"url": url},
                }

            if target_lower in KNOWN_SITE_URLS or target.startswith("http://") or target.startswith("https://") or " website" in lowered:
                if target_lower in KNOWN_SITE_URLS:
                    url = KNOWN_SITE_URLS[target_lower]
                else:
                    url = self._normalize_url(target.replace(" website", "").strip())
                return {
                    "intent": "open_website",
                    "action": "open_website",
                    "status": "success",
                    "dangerous": False,
                    "params": {"url": url},
                }

            app_like_targets = {
                "telegram", "whatsapp", "spotify", "vlc", "chrome", "edge", "firefox",
                "notepad", "calculator", "explorer", "word", "excel", "powerpoint",
                "media player", "windows media player", "code", "vs code", "github desktop",
            }

            if target_lower in app_like_targets:
                return {
                    "intent": "open_application",
                    "action": "open_application",
                    "status": "success",
                    "dangerous": False,
                    "params": {"app_name": target},
                }

            # Default to file search for plain names so commands like "open jananayagan" work.
            if self._looks_like_file_query(target) or target_lower:
                return {
                    "intent": "open_file",
                    "action": "open_file",
                    "status": "success",
                    "dangerous": False,
                    "params": {"query": target, "search_mode": "document"},
                }
            
            return {
                "intent": "open_application",
                "action": "open_application",
                "status": "success",
                "dangerous": False,
                "params": {"app_name": target},
            }

        # Play media commands
        if re.search(r"(?:create|make)\s+(?:a\s+)?(?:new\s+)?folder(?:\s+(?:on|in)\s+(desktop|home\s?page|home))?$", lowered):
            return {
                "intent": "create_folder",
                "action": "create_folder",
                "status": "success",
                "dangerous": False,
                "params": {"folder_name": "New Folder", "location": "desktop"},
            }

        # Create folder commands (named folder)
        create_named_folder = re.search(r"(?:create|make)\s+(?:a\s+)?folder\s+named\s+(.+?)(?:\s+(?:on|in)\s+(desktop|home\s?page|home))?$", command, flags=re.IGNORECASE)
        if create_named_folder:
            folder_name = create_named_folder.group(1).strip().strip("\"'")
            location = (create_named_folder.group(2) or "desktop").lower()
            location = "desktop" if location in {"desktop", "home page", "home"} else "desktop"
            return {
                "intent": "create_folder",
                "action": "create_folder",
                "status": "success",
                "dangerous": False,
                "params": {"folder_name": folder_name or "New Folder", "location": location},
            }

        # Create text file commands
        create_text_file = re.search(r"(?:(?:create|make)\s+(?:a\s+)?(?:new\s+)?)?text\s+file(?:\s+(?:on|in)\s+(desktop|home\s?page|home))?", lowered)
        if create_text_file:
            location = (create_text_file.group(1) or "desktop").lower()
            location = "desktop" if location in {"desktop", "home page", "home"} else "desktop"
            return {
                "intent": "create_text_file",
                "action": "create_text_file",
                "status": "success",
                "dangerous": False,
                "params": {"location": location, "filename": "New Text File.txt"},
            }

        create_folder = re.search(r"create\s+folder\s+named\s+(.+?)\s+on\s+desktop$", command, flags=re.IGNORECASE)
        if create_folder:
            folder_name = create_folder.group(1).strip().strip("\"'")
            return {
                "intent": "create_folder",
                "action": "create_folder",
                "status": "success",
                "dangerous": False,
                "params": {"folder_name": folder_name, "location": "desktop"},
            }

        delete_file = re.search(r"delete\s+file\s+(.+)$", command, flags=re.IGNORECASE)
        if delete_file:
            filename = delete_file.group(1).strip().strip("\"'")
            return {
                "intent": "delete_file",
                "action": "delete_file",
                "status": "success",
                "dangerous": True,
                "confirmation_message": f"Are you sure you want to delete {filename}?",
                "params": {"filename": filename},
            }

        search_file = re.search(r"search\s+for\s+(.+?)\s+file$", command, flags=re.IGNORECASE)
        if search_file:
            query = search_file.group(1).strip().strip("\"'")
            return {
                "intent": "search_file",
                "action": "search_file",
                "status": "success",
                "dangerous": False,
                "params": {"query": query},
            }

        if re.search(r"\bshutdown\b", lowered):
            return {
                "intent": "shutdown_system",
                "action": "shutdown_system",
                "status": "success",
                "dangerous": True,
                "confirmation_message": "Are you sure you want to shut down the system?",
                "params": {},
            }

        if re.search(r"\brestart\b", lowered):
            return {
                "intent": "restart_system",
                "action": "restart_system",
                "status": "success",
                "dangerous": True,
                "confirmation_message": "Are you sure you want to restart the system?",
                "params": {},
            }

        if re.search(r"\b(uninstall|format|factory reset)\b", lowered):
            return {
                "intent": "system_configuration_change",
                "action": "system_configuration_change",
                "status": "success",
                "dangerous": True,
                "confirmation_message": "This is a sensitive operation. Confirm to continue?",
                "params": {"raw_command": command},
            }

        system_keywords = [
            "open", "launch", "create", "delete", "remove", "rename", "move",
            "copy", "search", "find", "shutdown", "restart", "run", "execute",
            "volume", "brightness", "wifi", "bluetooth", "uninstall", "format",
        ]
        if any(keyword in lowered for keyword in system_keywords):
            return self._response(
                "unknown",
                "clarification_required",
                "failed",
                "Command unclear. Please specify exact action, target, and location.",
            )

        return self._response(
            "unknown",
            "not_system_command",
            "failed",
            "Not a system-control command.",
        )

    def _execute(self, parsed: Dict[str, Any]) -> Dict[str, Any]:
        intent = parsed["intent"]
        params = parsed.get("params", {})

        try:
            if intent == "conversation_reply":
                return self._response("conversation_reply", "conversation_reply", "success", params["message"])
            if intent == "learn_asr_correction":
                return self._learn_asr_correction(params["wrong"], params["right"])
            if intent == "forget_asr_correction":
                return self._forget_asr_correction(params["wrong"])
            if intent == "open_application":
                return self._open_application(params["app_name"])
            if intent == "create_folder":
                return self._create_folder(params["folder_name"], params["location"])
            if intent == "create_text_file":
                return self._create_text_file(params.get("location", "desktop"), params.get("filename", "New Text File.txt"))
            if intent == "open_folder":
                return self._open_folder(params["folder_name"])
            if intent == "play_media_in_folder":
                return self._play_media_in_folder(params.get("folder_hint", "telegram desktop"), params.get("query", ""))
            if intent == "open_file":
                return self._open_file(params["query"], params.get("search_mode", "auto"))
            if intent == "open_file_in_folder":
                return self._open_file_in_folder(params["query"], params["folder_name"])
            if intent == "delete_file":
                return self._delete_file(params["filename"])
            if intent == "search_file":
                return self._search_file(params["query"])
            if intent == "rename_file":
                return self._rename_file(params["source_name"], params["target_name"])
            if intent == "move_file":
                return self._move_file(params["source_name"], params["destination"])
            if intent == "copy_file":
                return self._copy_file(params["source_name"], params["destination"])
            if intent == "shutdown_system":
                return self._shutdown_system()
            if intent == "restart_system":
                return self._restart_system()
            if intent == "open_website":
                return self._open_website(params["url"])
            if intent == "toggle_wifi":
                return self._toggle_wifi(params["enable"])
            if intent == "toggle_bluetooth":
                return self._toggle_bluetooth(params["enable"])
            if intent == "set_volume":
                return self._set_volume(params["level"])
            if intent == "set_brightness":
                return self._set_brightness(params["level"])
            if intent == "run_system_command":
                return self._run_system_command(params["command"])
            if intent == "system_configuration_change":
                return self._response(
                    intent,
                    "blocked",
                    "failed",
                    "Sensitive system configuration changes require explicit dedicated handlers.",
                )
        except Exception as exc:
            return self._response(intent, parsed.get("action", "execute"), "failed", str(exc))

        return self._response(intent, "none", "failed", "Unsupported intent.")

    def _open_application(self, app_name: str) -> Dict[str, Any]:
        app_lower = app_name.lower().strip()

        if app_lower == "whatsapp":
            desktop_opened = self._try_open_whatsapp_desktop()
            if desktop_opened:
                return self._response("open_application", "open_application", "success", "Opened WhatsApp desktop app.")
            return self._response(
                "open_application",
                "open_application",
                "failed",
                "Could not find WhatsApp desktop app path on this PC. Say 'open whatsapp web' to use browser, or share the install path so I can add it.",
            )

        # Apps with known installation paths - try desktop app first, then web fallback
        KNOWN_APP_PATHS = {
            "whatsapp": {
                "paths": [Path(os.environ.get("LOCALAPPDATA", "")) / "WhatsApp" / "WhatsApp.exe"],
                "web": "https://web.whatsapp.com",
            },
            "telegram": {
                "paths": [
                    Path(os.environ.get("APPDATA", "")) / "Telegram Desktop" / "Telegram.exe",
                    Path(os.environ.get("LOCALAPPDATA", "")) / "Telegram Desktop" / "Telegram.exe",
                ],
                "web": "https://web.telegram.org",
            },
            "spotify": {
                "paths": [
                    Path(os.environ.get("APPDATA", "")) / "Spotify" / "Spotify.exe",
                    Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WindowsApps" / "Spotify.exe",
                ],
                "web": "https://open.spotify.com",
            },
        }

        if app_lower in KNOWN_APP_PATHS:
            info = KNOWN_APP_PATHS[app_lower]
            for p in info["paths"]:
                if p.exists():
                    subprocess.Popen([str(p)], shell=False)
                    return self._response("open_application", "open_application", "success", f"Opened {app_name}.")
            webbrowser.open(info["web"])
            return self._response("open_application", "open_application", "success", f"{app_name} app not found. Opened web version instead.")

        executable = shutil.which(app_name) or shutil.which(app_lower)
        if executable:
            subprocess.Popen([executable], shell=False)
            return self._response("open_application", "open_application", "success", f"Opened {app_name}.")

        if self.os_name == "windows":
            discovered = self._discover_windows_app_executable(app_name)
            if discovered:
                result = subprocess.run(["powershell", "-Command", f"Start-Process -FilePath '{discovered}'"], capture_output=True, text=True)
                if result.returncode == 0:
                    return self._response("open_application", "open_application", "success", f"Opened {app_name}.")

            result = subprocess.run(["powershell", "-Command", f"Start-Process '{app_name}'"], capture_output=True, text=True)
            if result.returncode == 0:
                return self._response("open_application", "open_application", "success", f"Opened {app_name}.")

            # Final fallback: search for it as a local file (movie, doc, music, etc.)
            file_matches = self._find_files(app_name, limit=1)
            if file_matches:
                found_path = file_matches[0]
                try:
                    os.startfile(str(found_path))
                    return self._response("open_application", "open_application", "success", f"Opened file: {found_path.name}.")
                except Exception as e:
                    return self._response("open_application", "open_application", "failed", f"Found {found_path.name} but could not open it: {e}")

            return self._response("open_application", "open_application", "failed", f"Could not find '{app_name}' as an app or file on this computer.")

        return self._response("open_application", "open_application", "failed", f"{app_name} is not installed on this system.")

    def _create_folder(self, folder_name: str, location: str) -> Dict[str, Any]:
        if location.lower() != "desktop":
            return self._response("create_folder", "create_folder", "failed", "Only Desktop location is currently supported.")

        base = Path.home() / "Desktop"
        target = self._unique_path(base, folder_name, is_file=False)
        target.mkdir(parents=True, exist_ok=True)
        return self._response("create_folder", "create_folder", "success", f"Created folder {target.name} on Desktop.")

    def _create_text_file(self, location: str, filename: str) -> Dict[str, Any]:
        base = Path.home() / "Desktop" if location.lower() in {"desktop", "home", "home page"} else Path.home()
        target = self._unique_path(base, filename, is_file=True)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("", encoding="utf-8")
        return self._response("create_text_file", "create_text_file", "success", f"Created text file {target.name} on Desktop.")

    def _open_file(self, query: str, search_mode: str = "auto") -> Dict[str, Any]:
        matches = self._find_files(query, limit=30, search_mode=search_mode)
        if not matches:
            return self._response("open_file", "open_file", "failed", f"I could not find {query} on this computer.")

        exact_same_name = self._same_name_matches(matches, query)
        if len(exact_same_name) > 1:
            joined = self._format_numbered_paths(exact_same_name)
            self.pending_file_choices = {"matches": exact_same_name, "query": query}
            return self._response(
                "open_file",
                "open_file",
                "failed",
                f"I found multiple files named {exact_same_name[0].name}.\n{joined}\nSay open first one, open second one, or open number 1.",
            )

        if len(exact_same_name) == 1:
            target = exact_same_name[0]
        else:
            target = self._pick_best_match(matches, query, search_mode=search_mode)

        try:
            os.startfile(str(target))
            return self._response("open_file", "open_file", "success", f"Opened {target.name}.")
        except Exception as exc:
            return self._response("open_file", "open_file", "failed", f"Found {target.name} but could not open it: {exc}")

    def _open_file_in_folder(self, query: str, folder_name: str) -> Dict[str, Any]:
        target_folder = self._resolve_folder_name(folder_name)
        if not target_folder or not target_folder.exists() or not target_folder.is_dir():
            return self._response("open_file_in_folder", "open_file_in_folder", "failed", f"I could not find the folder {folder_name}.")

        file_matches: List[Path] = []
        folder_matches: List[Path] = []
        query_lower = query.lower().strip()
        normalized_query = self._normalize_media_query(query)
        normalized_tokens = [t for t in normalized_query.split() if len(t) >= 2]
        for current_root, dirs, files in os.walk(target_folder, topdown=True):
            dirs[:] = [d for d in dirs if d.lower() not in {".git", "__pycache__", "node_modules"}]

            for dirname in dirs:
                if query_lower in dirname.lower():
                    folder_matches.append(Path(current_root) / dirname)
                    if len(folder_matches) >= 10:
                        break

            for filename in files:
                filename_lower = filename.lower()
                stem_lower = Path(filename).stem.lower()
                token_hit = normalized_tokens and all(t in stem_lower for t in normalized_tokens)
                if query_lower in filename_lower or normalized_query in stem_lower or token_hit:
                    file_matches.append(Path(current_root) / filename)
                    if len(file_matches) >= 15:
                        break
            if len(file_matches) >= 15 and len(folder_matches) >= 10:
                break

        # Folder-first: if the requested item is a folder inside the folder, open it.
        if folder_matches:
            chosen_folder = self._pick_best_match(folder_matches, query)
            try:
                os.startfile(str(chosen_folder))
                return self._response(
                    "open_file_in_folder",
                    "open_file_in_folder",
                    "success",
                    f"Opened folder {chosen_folder.name} inside {folder_name}.",
                )
            except Exception as exc:
                return self._response(
                    "open_file_in_folder",
                    "open_file_in_folder",
                    "failed",
                    f"Found folder {chosen_folder.name} in {folder_name}, but could not open it: {exc}",
                )

        if not file_matches:
            # If query is generic (movie/video/song), pick the most recent media file in the folder.
            generic_media_words = {
                "movie", "video", "song", "audio", "music", "file", "latest movie", "latest video",
                "pdf", "image", "photo", "document", "doc", "presentation", "ppt", "spreadsheet", "excel",
            }
            if query_lower in generic_media_words:
                media_matches: List[Path] = []
                media_exts = {
                    ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".m4v", ".webm", ".mp3", ".wav", ".m4a",
                    ".pdf", ".txt", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
                    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp",
                }
                for current_root, dirs, files in os.walk(target_folder, topdown=True):
                    dirs[:] = [d for d in dirs if d.lower() not in {".git", "__pycache__", "node_modules"}]
                    for filename in files:
                        if Path(filename).suffix.lower() in media_exts:
                            media_matches.append(Path(current_root) / filename)
                if media_matches:
                    target = max(media_matches, key=lambda p: p.stat().st_mtime)
                    try:
                        os.startfile(str(target))
                        return self._response(
                            "open_file_in_folder",
                            "open_file_in_folder",
                            "success",
                            f"Opened {target.name} from {folder_name} folder.",
                        )
                    except Exception as exc:
                        return self._response(
                            "open_file_in_folder",
                            "open_file_in_folder",
                            "failed",
                            f"Found {target.name} in {folder_name}, but could not open it: {exc}",
                        )

            return self._response("open_file_in_folder", "open_file_in_folder", "failed", f"I found {folder_name}, but no item matched {query}.")

        exact_same_name = self._same_name_matches(file_matches, query)
        if len(exact_same_name) > 1:
            joined = self._format_numbered_paths(exact_same_name)
            self.pending_file_choices = {"matches": exact_same_name, "query": query}
            return self._response(
                "open_file_in_folder",
                "open_file_in_folder",
                "failed",
                f"I found multiple files named {exact_same_name[0].name} in {folder_name}.\n{joined}\nSay open first one, open second one, or open number 1.",
            )

        target = exact_same_name[0] if len(exact_same_name) == 1 else self._pick_best_match(file_matches, query, search_mode="media")
        try:
            os.startfile(str(target))
            return self._response(
                "open_file_in_folder",
                "open_file_in_folder",
                "success",
                f"Opened {target.name} from {folder_name} folder.",
            )
        except Exception as exc:
            return self._response(
                "open_file_in_folder",
                "open_file_in_folder",
                "failed",
                f"Found {target.name} in {folder_name}, but could not open it: {exc}",
            )

    def _open_folder(self, folder_name: str) -> Dict[str, Any]:
        target = self._resolve_folder_name(folder_name) or Path(folder_name)
        if not target.exists():
            return self._response("open_folder", "open_folder", "failed", f"I could not find the folder {folder_name}.")
        try:
            os.startfile(str(target))
            return self._response("open_folder", "open_folder", "success", f"Opened {folder_name} folder.")
        except Exception as exc:
            return self._response("open_folder", "open_folder", "failed", f"Found {folder_name} but could not open it: {exc}")

    def _play_media_in_folder(self, folder_hint: str, query: str = "") -> Dict[str, Any]:
        folders = self._find_folders(folder_hint, limit=5)
        if not folders and folder_hint.lower() == "telegram desktop":
            # Common Telegram locations
            folders = [
                Path(os.environ.get("APPDATA", "")) / "Telegram Desktop",
                Path(os.environ.get("LOCALAPPDATA", "")) / "Telegram Desktop",
            ]
            folders = [f for f in folders if f.exists()]

        if not folders:
            return self._response("play_media_in_folder", "play_media_in_folder", "failed", f"I could not find the folder {folder_hint}.")

        media_files: List[Path] = []
        for folder in folders:
            try:
                for ext in ("*.mp4", "*.mkv", "*.avi", "*.mov", "*.wmv", "*.m4v", "*.webm"):
                    media_files.extend(folder.rglob(ext))
            except Exception:
                continue

        if not media_files:
            return self._response("play_media_in_folder", "play_media_in_folder", "failed", f"I found {folder_hint}, but there are no videos in it.")

        # If a query exists, try name match first
        chosen: Optional[Path] = None
        if query:
            q = query.lower().strip()
            for item in media_files:
                if q in item.name.lower():
                    chosen = item
                    break

        # fallback to latest modified video
        if not chosen:
            chosen = max(media_files, key=lambda p: p.stat().st_mtime)

        try:
            os.startfile(str(chosen))
            return self._response("play_media_in_folder", "play_media_in_folder", "success", f"Playing {chosen.name} from {folder_hint}.")
        except Exception as exc:
            return self._response("play_media_in_folder", "play_media_in_folder", "failed", f"I found {chosen.name} but could not play it: {exc}")

    def _delete_file(self, filename: str) -> Dict[str, Any]:
        matches = self._find_files(filename, limit=3)
        if not matches:
            return self._response("delete_file", "delete_file", "failed", f"File not found: {filename}")
        if len(matches) > 1:
            joined = "; ".join(str(path) for path in matches)
            return self._response("delete_file", "delete_file", "failed", f"Multiple files matched: {joined}")

        target = matches[0]
        target.unlink()
        return self._response("delete_file", "delete_file", "success", f"Deleted {target}")

    def _search_file(self, query: str) -> Dict[str, Any]:
        matches = self._find_files(query, limit=10)
        if not matches:
            return self._response("search_file", "search_file", "failed", f"No files found for {query}")

        joined = "\n".join(str(path) for path in matches)
        return self._response("search_file", "search_file", "success", f"Found files:\n{joined}")

    def _rename_file(self, source_name: str, target_name: str) -> Dict[str, Any]:
        matches = self._find_files(source_name, limit=3)
        if not matches:
            return self._response("rename_file", "rename_file", "failed", f"File not found: {source_name}")
        if len(matches) > 1:
            joined = "; ".join(str(path) for path in matches)
            return self._response("rename_file", "rename_file", "failed", f"Multiple files matched: {joined}")

        source = matches[0]
        target = source.with_name(target_name)
        source.rename(target)
        return self._response("rename_file", "rename_file", "success", f"Renamed to {target}")

    def _move_file(self, source_name: str, destination: str) -> Dict[str, Any]:
        matches = self._find_files(source_name, limit=3)
        if not matches:
            return self._response("move_file", "move_file", "failed", f"File not found: {source_name}")
        if len(matches) > 1:
            joined = "; ".join(str(path) for path in matches)
            return self._response("move_file", "move_file", "failed", f"Multiple files matched: {joined}")

        source = matches[0]
        destination_path = self._resolve_destination(destination)
        destination_path.mkdir(parents=True, exist_ok=True)
        moved_path = shutil.move(str(source), str(destination_path / source.name))
        return self._response("move_file", "move_file", "success", f"Moved to {moved_path}")

    def _copy_file(self, source_name: str, destination: str) -> Dict[str, Any]:
        matches = self._find_files(source_name, limit=3)
        if not matches:
            return self._response("copy_file", "copy_file", "failed", f"File not found: {source_name}")
        if len(matches) > 1:
            joined = "; ".join(str(path) for path in matches)
            return self._response("copy_file", "copy_file", "failed", f"Multiple files matched: {joined}")

        source = matches[0]
        destination_path = self._resolve_destination(destination)
        destination_path.mkdir(parents=True, exist_ok=True)
        copied_path = shutil.copy2(source, destination_path / source.name)
        return self._response("copy_file", "copy_file", "success", f"Copied to {copied_path}")

    def _find_files(self, query: str, limit: int = 10, search_mode: str = "auto") -> List[Path]:
        scored_results: List[tuple[int, Path]] = []
        query_lower = (query or "").lower().strip()
        normalized_query = self._normalize_media_query(query_lower)
        query_tokens = [t for t in normalized_query.split() if len(t) >= 2]
        query_ext = Path(query_lower).suffix.lower()
        media_exts = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".m4v", ".webm", ".mp3", ".wav", ".m4a"}
        document_exts = {".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".ppt", ".pptx", ".xls", ".xlsx"}
        media_intent = bool(re.search(r"\b(movie|video|song|audio|music|mkv|mp4|avi|mov|wmv|webm)\b", normalized_query)) or query_ext in media_exts
        document_intent = search_mode == "document" or bool(re.search(r"\b(pdf|doc|docx|txt|word|document|ppt|pptx|excel|sheet|presentation)\b", normalized_query)) or query_ext in document_exts

        # Stage 1: user folders only for low-latency lookup; Stage 2: full drives if needed.
        roots = self._get_search_roots(include_drives=False)
        skip_dirs = {
            "windows", "system32", "syswow64", "$recycle.bin", "program files", "program files (x86)",
            "programdata", "appdata", "node_modules", ".git", "__pycache__"
        }

        def scan_roots(scan_roots: List[Path]) -> None:
            for root in scan_roots:
                if not root.exists():
                    continue
                for current_root, dirs, files in os.walk(root, topdown=True):
                    dirs[:] = [d for d in dirs if d.lower() not in skip_dirs and not d.startswith(".")]
                    for filename in files:
                        name_lower = filename.lower()
                        stem_lower = Path(filename).stem.lower()
                        ext_lower = Path(filename).suffix.lower()

                        score = 0
                        if ext_lower in media_exts:
                            score += 35
                        if ext_lower in document_exts:
                            score += 20
                        if query_lower and query_lower == name_lower:
                            score += 300
                        elif query_lower and query_lower in name_lower:
                            score += 140

                        if normalized_query:
                            if normalized_query == stem_lower:
                                score += 220
                            elif normalized_query in stem_lower:
                                score += 120

                        if query_tokens:
                            overlap = sum(1 for token in query_tokens if token in stem_lower)
                            score += overlap * 24

                        if query_ext and ext_lower == query_ext:
                            score += 40

                        if search_mode == "media":
                            if ext_lower in media_exts:
                                score += 60
                            elif ext_lower in document_exts:
                                score -= 15
                            else:
                                score -= 25
                        elif document_intent:
                            if ext_lower in document_exts:
                                score += 55
                            elif ext_lower in media_exts:
                                score -= 15
                            else:
                                score -= 10
                        elif media_intent and ext_lower not in media_exts:
                            score -= 30

                        if normalized_query and len(normalized_query) >= 4:
                            ratio = difflib.SequenceMatcher(None, normalized_query, stem_lower).ratio()
                            if ratio >= 0.66:
                                score += int(ratio * 80)

                        if score > 0:
                            if (search_mode == "media" or media_intent) and score < 30:
                                continue
                            scored_results.append((score, Path(current_root) / filename))

        scan_roots(roots)
        if not scored_results:
            scan_roots(self._get_search_roots(include_drives=True))

        if not scored_results:
            return []

        scored_results.sort(key=lambda item: item[0], reverse=True)
        deduped: List[Path] = []
        seen: set[str] = set()
        for _, path in scored_results:
            key = str(path).lower()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(path)
            if len(deduped) >= limit:
                break
        return deduped

    def _get_search_roots(self, include_drives: bool = True) -> List[Path]:
        roots: List[Path] = [
            Path.home(),
            Path.home() / "Desktop",
            Path.home() / "Documents",
            Path.home() / "Downloads",
            Path.home() / "Videos",
            Path.home() / "Music",
            Path.home() / "Pictures",
        ]

        if include_drives and self.os_name == "windows":
            for letter in ["C", "D", "E", "F", "G"]:
                drive = Path(f"{letter}:\\")
                if drive.exists():
                    roots.append(drive)

        deduped: List[Path] = []
        seen: set[str] = set()
        for root in roots:
            key = str(root).lower()
            if key not in seen:
                seen.add(key)
                deduped.append(root)
        return deduped

    def _find_folders(self, query: str, limit: int = 10) -> List[Path]:
        results: List[Path] = []
        roots = [Path.home(), Path(os.environ.get("APPDATA", "")), Path(os.environ.get("LOCALAPPDATA", ""))]
        query_lower = query.lower().strip()
        seen: set[str] = set()

        for root in roots:
            if not root.exists():
                continue
            for current_root, dirs, _ in os.walk(root, topdown=True):
                for dirname in dirs:
                    if query_lower in dirname.lower():
                        p = Path(current_root) / dirname
                        key = str(p).lower()
                        if key in seen:
                            continue
                        seen.add(key)
                        results.append(p)
                        if len(results) >= limit:
                            return results
        return results

    def _pick_best_match(self, matches: List[Path], query: str, search_mode: str = "auto") -> Path:
        media_exts = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".m4v", ".webm", ".mp3", ".wav", ".m4a"}
        document_exts = {".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".ppt", ".pptx", ".xls", ".xlsx"}
        query_lower = query.lower()
        normalized_query = self._normalize_media_query(query)
        query_tokens = [t for t in normalized_query.split() if len(t) >= 2]
        scored: List[tuple[int, Path]] = []
        for path in matches:
            score = 0
            stem_lower = path.stem.lower()
            suffix = path.suffix.lower()
            if suffix in media_exts:
                score += 4
            if suffix in document_exts:
                score += 3
            if query_lower == stem_lower:
                score += 6
            elif query_lower in stem_lower:
                score += 3
            if normalized_query and normalized_query == stem_lower:
                score += 7
            elif normalized_query and normalized_query in stem_lower:
                score += 4
            if query_tokens:
                overlap = sum(1 for token in query_tokens if token in stem_lower)
                score += overlap * 2
            if search_mode == "media" and suffix in media_exts:
                score += 5
            if search_mode == "document" and suffix in document_exts:
                score += 5
            scored.append((score, path))
        scored.sort(key=lambda item: item[0], reverse=True)
        return scored[0][1]

    def _pick_best_media_match(self, matches: List[Path], query: str) -> Path:
        return self._pick_best_match(matches, query, search_mode="media")

    def _same_name_matches(self, matches: List[Path], query: str) -> List[Path]:
        normalized_query = (query or "").strip().strip("\"'").lower()
        if not normalized_query:
            return []

        query_path = Path(normalized_query)
        has_ext = bool(query_path.suffix)
        if has_ext:
            return [p for p in matches if p.name.lower() == query_path.name.lower()]
        return [p for p in matches if p.stem.lower() == query_path.stem.lower()]

    def _normalize_media_query(self, query: str) -> str:
        text = (query or "").strip().lower()
        if not text:
            return ""

        # Common ASR extension confusions and spoken forms.
        text = re.sub(r"\btkv\b", "mkv", text)
        text = re.sub(r"\bm\s*k\s*v\b", "mkv", text)
        text = re.sub(r"\bm\s*p\s*4\b", "mp4", text)
        text = re.sub(r"\bdot\s+(mkv|mp4|avi|mov|wmv|webm)\b", r".\1", text)

        # Remove common filler words often spoken with media commands.
        filler_pattern = (
            r"\b(open|play|movie|video|vlc|file|please|the|a|an|in|from|inside|folder)\b"
        )
        text = re.sub(filler_pattern, " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def _format_numbered_paths(self, paths: List[Path]) -> str:
        lines = [f"{idx}. {path}" for idx, path in enumerate(paths, start=1)]
        return "\n".join(lines)

    def _looks_like_file_query(self, text: str) -> bool:
        target = (text or "").strip().lower()
        if not target:
            return False

        # Has extension-like suffix.
        if re.search(r"\.[a-z0-9]{2,5}$", target):
            return True

        file_hint_words = {
            "file", "folder", "pdf", "image", "photo", "movie", "video", "song", "audio",
            "document", "doc", "ppt", "excel", "spreadsheet", "txt",
        }
        if any(word in target for word in file_hint_words):
            return True

        return False

    def _unique_path(self, base: Path, name: str, is_file: bool) -> Path:
        candidate = base / name
        if not candidate.exists():
            return candidate

        stem = candidate.stem if is_file else candidate.name
        suffix = candidate.suffix if is_file else ""
        for i in range(2, 1000):
            trial_name = f"{stem} ({i}){suffix}"
            trial = base / trial_name
            if not trial.exists():
                return trial
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        return base / f"{stem}_{timestamp}{suffix}"

    def _resolve_destination(self, destination: str) -> Path:
        destination_lower = destination.strip().lower()
        if destination_lower == "desktop":
            return Path.home() / "Desktop"
        if destination_lower == "documents":
            return Path.home() / "Documents"
        if destination_lower == "downloads":
            return Path.home() / "Downloads"

        expanded = os.path.expandvars(os.path.expanduser(destination.strip()))
        return Path(expanded)

    def _resolve_folder_name(self, folder_name: str) -> Optional[Path]:
        name = self._normalize_folder_alias(folder_name)
        mapping = {
            "desktop": Path.home() / "Desktop",
            "documents": Path.home() / "Documents",
            "downloads": Path.home() / "Downloads",
            "videos": Path.home() / "Videos",
            "movies": Path.home() / "Videos",
            "music": Path.home() / "Music",
            "pictures": Path.home() / "Pictures",
            "photos": Path.home() / "Pictures",
            "home": Path.home(),
        }
        if name in mapping:
            return mapping[name]

        maybe_path = Path(os.path.expandvars(os.path.expanduser(folder_name.strip())))
        if maybe_path.exists() and maybe_path.is_dir():
            return maybe_path

        found = self._find_folders(name, limit=1)
        return found[0] if found else None

    def _try_open_whatsapp_desktop(self) -> bool:
        whatsapp_candidates = [
            Path(os.environ.get("LOCALAPPDATA", "")) / "WhatsApp" / "WhatsApp.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "WhatsApp" / "WhatsApp.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WindowsApps" / "WhatsApp.exe",
            Path(os.environ.get("ProgramFiles", "")) / "WhatsApp" / "WhatsApp.exe",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "WhatsApp" / "WhatsApp.exe",
        ]

        for candidate in whatsapp_candidates:
            if candidate.exists():
                try:
                    subprocess.Popen([str(candidate)], shell=False)
                    return True
                except Exception:
                    pass

        shortcut_candidates = [
            Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "WhatsApp.lnk",
            Path(os.environ.get("PROGRAMDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "WhatsApp.lnk",
            Path.home() / "Desktop" / "WhatsApp.lnk",
        ]

        for shortcut in shortcut_candidates:
            if shortcut.exists():
                try:
                    os.startfile(str(shortcut))
                    return True
                except Exception:
                    pass

        for programs_root in [
            Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs",
            Path(os.environ.get("PROGRAMDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs",
        ]:
            if not programs_root.exists():
                continue
            try:
                for shortcut in programs_root.rglob("*WhatsApp*.lnk"):
                    try:
                        os.startfile(str(shortcut))
                        return True
                    except Exception:
                        continue
            except Exception:
                continue

        protocol_attempt = subprocess.run(
            ["powershell", "-Command", "Start-Process 'whatsapp:'"],
            capture_output=True,
            text=True,
        )
        if protocol_attempt.returncode == 0:
            return True

        known_app_ids = [
            "5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App",
            "WhatsApp.WhatsAppDesktop_cv1g1gvanyjgm!App",
        ]
        for app_id in known_app_ids:
            appid_attempt = subprocess.run(
                ["powershell", "-Command", f"Start-Process 'shell:AppsFolder\\{app_id}'"],
                capture_output=True,
                text=True,
            )
            if appid_attempt.returncode == 0:
                return True

        appsfolder_attempt = subprocess.run(
            [
                "powershell",
                "-Command",
                "$app = Get-StartApps | Where-Object { $_.Name -like '*WhatsApp*' } | Select-Object -First 1; "
                "if ($app) { Start-Process ('shell:AppsFolder\\' + $app.AppID); exit 0 } else { exit 1 }",
            ],
            capture_output=True,
            text=True,
        )
        return appsfolder_attempt.returncode == 0

    def _toggle_wifi(self, enable: bool) -> Dict[str, Any]:
        if self.os_name != "windows":
            return self._response("toggle_wifi", "toggle_wifi", "failed", "Wi-Fi toggle is currently implemented for Windows only.")

        state = "ENABLED" if enable else "DISABLED"
        result = subprocess.run(
            ["netsh", "interface", "set", "interface", "name=Wi-Fi", f"admin={state}"],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            return self._response("toggle_wifi", "toggle_wifi", "failed", (result.stderr or result.stdout or "Failed to toggle Wi-Fi.").strip())

        return self._response("toggle_wifi", "toggle_wifi", "success", f"Wi-Fi {'enabled' if enable else 'disabled'}.")

    def _toggle_bluetooth(self, enable: bool) -> Dict[str, Any]:
        if self.os_name != "windows":
            return self._response("toggle_bluetooth", "toggle_bluetooth", "failed", "Bluetooth toggle is currently implemented for Windows only.")

        action = "Enable-PnpDevice" if enable else "Disable-PnpDevice"
        command = (
            "$devices = Get-PnpDevice -Class Bluetooth -Status OK -ErrorAction SilentlyContinue; "
            "if (-not $devices) { Write-Output 'No active Bluetooth device found.'; exit 1 }; "
            f"$devices | ForEach-Object {{ {action} -InstanceId $_.InstanceId -Confirm:$false -ErrorAction Stop }}"
        )
        result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
        if result.returncode != 0:
            return self._response(
                "toggle_bluetooth",
                "toggle_bluetooth",
                "failed",
                (result.stderr or result.stdout or "Failed to toggle Bluetooth. Administrator permissions may be required.").strip(),
            )

        return self._response("toggle_bluetooth", "toggle_bluetooth", "success", f"Bluetooth {'enabled' if enable else 'disabled'}.")

    def _set_volume(self, level: int) -> Dict[str, Any]:
        clamped = max(0, min(100, int(level)))
        if self.os_name == "windows":
            command = (
                "$v = New-Object -ComObject WScript.Shell; "
                f"1..50 | ForEach-Object {{ $v.SendKeys([char]174) }}; "
                f"1..{max(0, round(clamped / 2))} | ForEach-Object {{ $v.SendKeys([char]175) }}"
            )
            result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
            if result.returncode != 0:
                return self._response("set_volume", "set_volume", "failed", (result.stderr or result.stdout or "Failed to set volume.").strip())
            return self._response("set_volume", "set_volume", "success", f"Volume set to {clamped}.")

        if self.os_name == "macos":
            result = subprocess.run(["osascript", "-e", f"set volume output volume {clamped}"], capture_output=True, text=True)
            if result.returncode != 0:
                return self._response("set_volume", "set_volume", "failed", (result.stderr or result.stdout or "Failed to set volume.").strip())
            return self._response("set_volume", "set_volume", "success", f"Volume set to {clamped}.")

        return self._response("set_volume", "set_volume", "failed", "Volume control is not implemented for this OS.")

    def _set_brightness(self, level: int) -> Dict[str, Any]:
        clamped = max(0, min(100, int(level)))
        if self.os_name == "windows":
            command = (
                "Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightnessMethods | "
                f"ForEach-Object {{ $_.WmiSetBrightness(1,{clamped}) }}"
            )
            result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
            if result.returncode != 0:
                return self._response("set_brightness", "set_brightness", "failed", (result.stderr or result.stdout or "Failed to set brightness.").strip())
            return self._response("set_brightness", "set_brightness", "success", f"Brightness set to {clamped}.")

        return self._response("set_brightness", "set_brightness", "failed", "Brightness control is not implemented for this OS.")

    def _discover_windows_app_executable(self, app_name: str) -> Optional[str]:
        app_lower = app_name.lower().strip()
        candidate_exe_names = [
            f"{app_lower}.exe",
            f"{app_lower.replace(' ', '')}.exe",
        ]

        common_roots = [
            Path(os.environ.get("ProgramFiles", "")),
            Path(os.environ.get("ProgramFiles(x86)", "")),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs",
            Path(os.environ.get("LOCALAPPDATA", "")),
            Path(os.environ.get("APPDATA", "")),
        ]

        for root in common_roots:
            if not root.exists():
                continue
            for exe_name in candidate_exe_names:
                direct = root / app_name / exe_name
                if direct.exists():
                    return str(direct)

            try:
                for candidate in root.rglob("*.exe"):
                    stem = candidate.stem.lower()
                    if app_lower == stem or app_lower in stem:
                        return str(candidate)
            except Exception:
                continue

        return None

    def _run_system_command(self, command: str) -> Dict[str, Any]:
        if not self._is_allowed_shell_command(command):
            return self._response("run_system_command", "blocked", "failed", "Command is outside the safe allowlist.")

        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        output = (result.stdout or result.stderr or "").strip()

        if result.returncode != 0:
            message = output or f"Command failed with exit code {result.returncode}."
            return self._response("run_system_command", "run_system_command", "failed", message)

        return self._response("run_system_command", "run_system_command", "success", output or "Command executed successfully.")

    def _is_allowed_shell_command(self, command: str) -> bool:
        normalized = command.strip().lower()
        if not normalized:
            return False

        first_token = re.split(r"\s+", normalized, maxsplit=1)[0]
        allowlist = {
            "dir", "echo", "where", "whoami", "hostname", "ipconfig", "systeminfo",
            "tasklist", "ping", "netstat", "type", "findstr", "wmic", "powershell",
        }

        if first_token not in allowlist:
            return False

        blocked_patterns = [
            r"\b(del|erase|rd|rmdir|format|shutdown|restart|reboot|sc\s+delete|reg\s+delete)\b",
            r"\b(remove-item|set-itemproperty|new-itemproperty|stop-computer|restart-computer)\b",
            r"[;&]{2}|\|\|",
            r">\s*[^\s]+",
        ]
        return not any(re.search(pattern, normalized) for pattern in blocked_patterns)

    def _is_risky_shell_command(self, command: str) -> bool:
        normalized = command.strip().lower()
        risky_patterns = [
            r"\b(wmic|powershell)\b",
            r"\b(netsh|sc|reg)\b",
            r"\|",
            r"\b(taskkill)\b",
        ]
        return any(re.search(pattern, normalized) for pattern in risky_patterns)

    def _shutdown_system(self) -> Dict[str, Any]:
        if self.os_name == "windows":
            subprocess.Popen(["shutdown", "/s", "/t", "0"], shell=False)
        elif self.os_name == "macos":
            subprocess.Popen(["shutdown", "-h", "now"], shell=False)
        else:
            subprocess.Popen(["shutdown", "-h", "now"], shell=False)
        return self._response("shutdown_system", "shutdown_system", "success", "System shutdown initiated.")

    def _restart_system(self) -> Dict[str, Any]:
        if self.os_name == "windows":
            subprocess.Popen(["shutdown", "/r", "/t", "0"], shell=False)
        elif self.os_name == "macos":
            subprocess.Popen(["shutdown", "-r", "now"], shell=False)
        else:
            subprocess.Popen(["shutdown", "-r", "now"], shell=False)
        return self._response("restart_system", "restart_system", "success", "System restart initiated.")

    def _open_website(self, url: str) -> Dict[str, Any]:
        webbrowser.open(url)
        # Map known URLs to friendly spoken names
        URL_FRIENDLY_NAMES = {
            "mail.google.com": "Gmail",
            "open.spotify.com": "Spotify",
            "web.telegram.org": "Telegram",
            "web.whatsapp.com": "WhatsApp",
            "discord.com": "Discord",
            "chat.openai.com": "ChatGPT",
            "zoom.us": "Zoom",
        }
        host = url.replace("https://", "").replace("http://", "").split("/")[0]
        friendly = URL_FRIENDLY_NAMES.get(host)
        if not friendly:
            # Get the second-to-last part of the domain e.g. youtube from www.youtube.com
            parts = host.replace("www.", "").split(".")
            friendly = parts[-2] if len(parts) >= 2 else parts[0]
        return self._response("open_website", "open_website", "success", f"Opened {friendly.capitalize()} in your browser.")

    def _normalize_url(self, target: str) -> str:
        normalized = target.strip()
        if normalized.startswith("http://") or normalized.startswith("https://"):
            return normalized
        if "." in normalized:
            return f"https://{normalized}"
        return f"https://www.{normalized}.com"

    def _asr_adaptation_path(self) -> Path:
        custom = os.environ.get("JARVIS_ASR_ADAPT_FILE", "").strip()
        if custom:
            return Path(os.path.expandvars(os.path.expanduser(custom)))
        return Path(__file__).resolve().parent.parent / "asr_adaptation.json"

    def _load_asr_adaptation(self) -> Dict[str, Any]:
        path = self._asr_adaptation_path()
        try:
            if path.exists():
                payload = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(payload, dict):
                    payload.setdefault("phrase_hints", [])
                    payload.setdefault("replacements", {})
                    return payload
        except Exception:
            pass
        return {"phrase_hints": [], "replacements": {}}

    def _save_asr_adaptation(self, payload: Dict[str, Any]) -> None:
        path = self._asr_adaptation_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")

    def _learn_asr_correction(self, wrong: str, right: str) -> Dict[str, Any]:
        wrong_norm = wrong.lower().strip()
        right_norm = right.lower().strip()
        if not wrong_norm or not right_norm:
            return self._response("learn_asr_correction", "learn_asr_correction", "failed", "Please provide both wrong and correct phrases.")

        payload = self._load_asr_adaptation()
        replacements = payload.get("replacements", {})
        if not isinstance(replacements, dict):
            replacements = {}
        replacements[wrong_norm] = right_norm
        payload["replacements"] = replacements

        phrase_hints = payload.get("phrase_hints", [])
        if not isinstance(phrase_hints, list):
            phrase_hints = []
        for token in [wrong_norm, right_norm]:
            if token and token not in phrase_hints:
                phrase_hints.append(token)
        payload["phrase_hints"] = phrase_hints[:300]

        try:
            self._save_asr_adaptation(payload)
            return self._response(
                "learn_asr_correction",
                "learn_asr_correction",
                "success",
                f"Learned correction: {wrong_norm} means {right_norm}.",
            )
        except Exception as exc:
            return self._response("learn_asr_correction", "learn_asr_correction", "failed", f"Could not save correction: {exc}")

    def _forget_asr_correction(self, wrong: str) -> Dict[str, Any]:
        wrong_norm = wrong.lower().strip()
        if not wrong_norm:
            return self._response("forget_asr_correction", "forget_asr_correction", "failed", "Please provide the correction phrase to remove.")

        payload = self._load_asr_adaptation()
        replacements = payload.get("replacements", {})
        if not isinstance(replacements, dict):
            replacements = {}

        if wrong_norm not in replacements:
            return self._response("forget_asr_correction", "forget_asr_correction", "failed", f"No saved correction found for {wrong_norm}.")

        replacements.pop(wrong_norm, None)
        payload["replacements"] = replacements
        try:
            self._save_asr_adaptation(payload)
            return self._response("forget_asr_correction", "forget_asr_correction", "success", f"Removed correction for {wrong_norm}.")
        except Exception as exc:
            return self._response("forget_asr_correction", "forget_asr_correction", "failed", f"Could not remove correction: {exc}")

    def _response(self, intent: str, action: str, status: str, message: str) -> Dict[str, Any]:
        return {
            "intent": intent,
            "action": action,
            "status": status,
            "message": message,
        }
