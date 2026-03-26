import os
import json
import subprocess
import webbrowser
import sys
from typing import List, Dict, Any, Callable
from core.skill import Skill

# Map of common app names to their Windows executable paths or commands
APP_MAP = {
    "telegram": [
        os.path.join(os.environ.get("APPDATA", ""), "Telegram Desktop", "Telegram.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Telegram Desktop", "Telegram.exe"),
    ],
    "chrome": [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ],
    "google chrome": [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    ],
    "firefox": [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
    ],
    "vlc": [
        r"C:\Program Files\VideoLAN\VLC\vlc.exe",
        r"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe",
    ],
    "notepad": ["notepad.exe"],
    "calculator": ["calc.exe"],
    "wordpad": ["wordpad.exe"],
    "paint": ["mspaint.exe"],
    "word": [r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"],
    "excel": [r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"],
    "powerpoint": [r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"],
    "task manager": ["taskmgr.exe"],
    "file explorer": ["explorer.exe"],
    "explorer": ["explorer.exe"],
    "spotify": [
        os.path.join(os.environ.get("APPDATA", ""), "Spotify", "Spotify.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "WindowsApps", "Spotify.exe"),
    ],
    "discord": [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Discord", "Update.exe"),
        os.path.join(os.environ.get("APPDATA", ""), "Discord", "Discord.exe"),
    ],
    "whatsapp": [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "WhatsApp", "WhatsApp.exe"),
    ],
    "edge": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ],
    "microsoft edge": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ],
    "cmd": ["cmd.exe"],
    "command prompt": ["cmd.exe"],
    "powershell": ["powershell.exe"],
    "snipping tool": ["SnippingTool.exe"],
    "settings": ["ms-settings:"],
    "control panel": ["control.exe"],
}

# Apps that should open as a website in the browser
WEB_APPS = {
    "gmail": "https://mail.google.com",
    "google mail": "https://mail.google.com",
    "youtube": "https://www.youtube.com",
    "google": "https://www.google.com",
    "facebook": "https://www.facebook.com",
    "instagram": "https://www.instagram.com",
    "twitter": "https://www.twitter.com",
    "x": "https://www.x.com",
    "reddit": "https://www.reddit.com",
    "netflix": "https://www.netflix.com",
    "amazon": "https://www.amazon.com",
    "github": "https://www.github.com",
    "telegram web": "https://web.telegram.org",
}


class SystemSkill(Skill):
    @property
    def name(self) -> str:
        return "system_skill"

    def get_tools(self) -> List[Dict[str, Any]]:
        return [
            {
                "type": "function",
                "function": {
                    "name": "set_volume",
                    "description": "Set system volume level (0-100)",
                    "parameters": { "type": "object", "properties": { "level": {"type": "integer"} }, "required": ["level"] }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "open_app",
                    "description": "Open an application, program, or website on the computer. Use this for apps like Telegram, Chrome, Gmail, Spotify, VLC, Notepad, Calculator, etc.",
                    "parameters": { "type": "object", "properties": { "app_name": {"type": "string", "description": "Name of the app to open, e.g. 'telegram', 'gmail', 'chrome', 'calculator'"} }, "required": ["app_name"] }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "shutdown_computer",
                    "description": "Shutdown or restart the computer",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "action": {"type": "string", "enum": ["shutdown", "restart"], "description": "Whether to shutdown or restart"}
                        },
                        "required": ["action"]
                    }
                }
            }
        ]

    def get_functions(self) -> Dict[str, Callable]:
        return {
            "set_volume": self.set_volume,
            "open_app": self.open_app,
            "shutdown_computer": self.shutdown_computer,
        }

    def set_volume(self, level):
        try:
            level = max(0, min(100, int(level)))
            # Use PowerShell to set volume on Windows
            script = (
                f"$vol = {level} / 100.0; "
                "$obj = New-Object -ComObject WScript.Shell; "
                "Add-Type -TypeDefinition @'\n"
                "using System.Runtime.InteropServices;\n"
                "[Guid(\"5CDF2C82-841E-4546-9722-0CF74078229A\"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]\n"
                "interface IAudioEndpointVolume { }\n"
                "'@; "
                f"$wshShell = New-Object -ComObject WScript.Shell; "
                f"$vol_steps = [math]::Round($vol * 65535); "
            )
            # Simpler: use nircmd if available, else PowerShell audio API
            result = subprocess.run(
                ["powershell", "-Command",
                 f"$obj = New-Object -ComObject Shell.Application; "
                 f"(New-Object -ComObject WScript.Shell).SendKeys([char]0xAD)"],
                capture_output=True, timeout=5
            )
            # Use nircmd as a fallback approach via os system
            os.system(f'powershell -c "$vol={level}/100; $sink=(Get-AudioDevice -Playback).Index; Set-AudioDevice -PlaybackVolume $vol" 2>nul')
            return json.dumps({"status": "success", "message": f"Volume set to {level} percent."})
        except Exception as e:
            return json.dumps({"error": str(e)})

    def open_app(self, app_name):
        app_lower = app_name.lower().strip()

        # Check if it's a web app first
        for key, url in WEB_APPS.items():
            if key in app_lower or app_lower in key:
                try:
                    webbrowser.open(url)
                    return json.dumps({"status": "success", "message": f"Opened {app_name} in your browser."})
                except Exception as e:
                    return json.dumps({"error": str(e)})

        # Check known app map
        for key, paths in APP_MAP.items():
            if key in app_lower or app_lower in key:
                for path in paths:
                    # Handle ms-settings: style URIs
                    if path.startswith("ms-"):
                        try:
                            os.startfile(path)
                            return json.dumps({"status": "success", "message": f"Opened {app_name}."})
                        except:
                            continue
                    if os.path.exists(path):
                        try:
                            subprocess.Popen([path], shell=False)
                            return json.dumps({"status": "success", "message": f"Opened {app_name}."})
                        except Exception as e:
                            continue
                    elif path.endswith(".exe") and not os.sep in path:
                        # Simple exe name like notepad.exe - just run it
                        try:
                            subprocess.Popen(path, shell=True)
                            return json.dumps({"status": "success", "message": f"Opened {app_name}."})
                        except Exception:
                            continue

        # Final fallback: try to run using Windows 'start' command
        try:
            subprocess.Popen(f'start "" "{app_name}"', shell=True)
            return json.dumps({"status": "success", "message": f"Attempting to open {app_name}."})
        except Exception as e:
            return json.dumps({"error": f"Could not open {app_name}. Please make sure it is installed. {str(e)}"})

    def shutdown_computer(self, action):
        try:
            if action == "shutdown":
                os.system("shutdown /s /t 10")
                return json.dumps({"status": "success", "message": "Shutting down in 10 seconds."})
            elif action == "restart":
                os.system("shutdown /r /t 10")
                return json.dumps({"status": "success", "message": "Restarting in 10 seconds."})
        except Exception as e:
            return json.dumps({"error": str(e)})
