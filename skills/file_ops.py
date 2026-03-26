import os
import json
import subprocess
import glob
from typing import List, Dict, Any, Callable
from core.skill import Skill

# Common search locations on Windows
SEARCH_ROOTS = [
    os.path.join(os.path.expanduser("~"), "Desktop"),
    os.path.join(os.path.expanduser("~"), "Documents"),
    os.path.join(os.path.expanduser("~"), "Downloads"),
    os.path.join(os.path.expanduser("~"), "Videos"),
    os.path.join(os.path.expanduser("~"), "Music"),
    os.path.join(os.path.expanduser("~"), "Pictures"),
    os.path.join(os.path.expanduser("~"), "Movies"),
    "C:\\",
    "D:\\",
    "E:\\",
]

VIDEO_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".mpg", ".mpeg"}
AUDIO_EXTENSIONS = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a"}
DOC_EXTENSIONS   = {".pdf", ".docx", ".doc", ".txt", ".xlsx", ".xls", ".pptx", ".ppt", ".csv"}


def _find_file(filename: str, search_dirs=None, max_depth=5) -> str | None:
    """Search for a file by name (partial match) across common directories."""
    if search_dirs is None:
        search_dirs = SEARCH_ROOTS
    filename_lower = filename.lower()
    for root_dir in search_dirs:
        if not os.path.exists(root_dir):
            continue
        for dirpath, dirnames, filenames in os.walk(root_dir):
            # Limit recursion depth
            depth = dirpath[len(root_dir):].count(os.sep)
            if depth >= max_depth:
                dirnames.clear()
                continue
            # Skip hidden/system dirs
            dirnames[:] = [d for d in dirnames if not d.startswith('.') and d not in 
                           ('$Recycle.Bin', 'Windows', 'System32', 'SysWOW64', 'Program Files', 
                            'Program Files (x86)', '__pycache__', 'node_modules', '.git')]
            for fname in filenames:
                if filename_lower in fname.lower():
                    return os.path.join(dirpath, fname)
    return None


def _open_file(filepath: str) -> dict:
    """Open a file using the default Windows application."""
    try:
        os.startfile(filepath)
        return {"status": "success", "message": f"Opened {os.path.basename(filepath)}."}
    except Exception as e:
        try:
            subprocess.Popen(["explorer", filepath])
            return {"status": "success", "message": f"Opened {os.path.basename(filepath)}."}
        except Exception as e2:
            return {"error": str(e2)}


class FileSkill(Skill):
    @property
    def name(self) -> str:
        return "file_skill"

    def get_tools(self) -> List[Dict[str, Any]]:
        return [
            {
                "type": "function",
                "function": {
                    "name": "manage_file",
                    "description": "Create, read, write, or append to text files on the Desktop.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "action": {"type": "string", "enum": ["read", "write", "create", "append"]},
                            "filename": {"type": "string"},
                            "content": {"type": "string"}
                        },
                        "required": ["action", "filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "open_local_file",
                    "description": "Find and open any local file on the computer including movies, videos, music, documents, or any file by name. Searches across Desktop, Documents, Downloads, Videos, Music, Pictures and drives.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {
                                "type": "string",
                                "description": "Name or partial name of the file to search for and open, e.g. 'Avengers', 'homework', 'my song'"
                            },
                            "file_type": {
                                "type": "string",
                                "enum": ["any", "video", "audio", "document", "image"],
                                "description": "Type of file to search for. Default is 'any'."
                            }
                        },
                        "required": ["filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "open_folder",
                    "description": "Open a folder or directory on the computer in File Explorer, e.g. 'Videos', 'Documents', 'Downloads', 'Desktop', 'Music', 'Pictures' or any folder path.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "folder_name": {
                                "type": "string",
                                "description": "Name of the folder to open, e.g. 'Videos', 'Desktop', 'Downloads', or a full path"
                            }
                        },
                        "required": ["folder_name"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "list_folder_contents",
                    "description": "List the files and folders inside a directory, e.g. list what movies are in the Videos folder.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "folder_name": {
                                "type": "string",
                                "description": "Name of the folder to list, e.g. 'Videos', 'Desktop', 'Downloads', 'Movies'"
                            }
                        },
                        "required": ["folder_name"]
                    }
                }
            }
        ]

    def get_functions(self) -> Dict[str, Callable]:
        return {
            "manage_file": self.manage_file,
            "open_local_file": self.open_local_file,
            "open_folder": self.open_folder,
            "list_folder_contents": self.list_folder_contents,
        }

    def manage_file(self, action: str, filename: str, content: str = ""):
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            filepath = os.path.join(desktop_path, filename)

            if action == "read":
                if os.path.exists(filepath):
                    with open(filepath, 'r') as f:
                        data = f.read()
                    return json.dumps({"status": "success", "content": data})
                else:
                    return json.dumps({"error": "File not found on Desktop."})

            elif action in ["write", "create"]:
                with open(filepath, 'w') as f:
                    f.write(content)
                return json.dumps({"status": "success", "message": f"Created {filename} on Desktop."})

            elif action == "append":
                with open(filepath, 'a') as f:
                    f.write("\n" + content)
                return json.dumps({"status": "success", "message": f"Updated {filename}."})

        except Exception as e:
            return json.dumps({"error": str(e)})

    def open_local_file(self, filename: str, file_type: str = "any"):
        try:
            # Determine which extensions to search for
            ext_filter = None
            if file_type == "video":
                ext_filter = VIDEO_EXTENSIONS
            elif file_type == "audio":
                ext_filter = AUDIO_EXTENSIONS
            elif file_type == "document":
                ext_filter = DOC_EXTENSIONS
            elif file_type == "image":
                ext_filter = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff"}

            filename_lower = filename.lower()

            # Search through common dirs
            for root_dir in SEARCH_ROOTS:
                if not os.path.exists(root_dir):
                    continue
                for dirpath, dirnames, filenames in os.walk(root_dir):
                    depth = dirpath[len(root_dir):].count(os.sep)
                    if depth >= 6:
                        dirnames.clear()
                        continue
                    dirnames[:] = [d for d in dirnames if not d.startswith('.') and d not in
                                   ('$Recycle.Bin', 'Windows', 'System32', 'SysWOW64',
                                    'Program Files', 'Program Files (x86)', '__pycache__',
                                    'node_modules', '.git')]
                    for fname in filenames:
                        fname_lower = fname.lower()
                        ext = os.path.splitext(fname)[1].lower()
                        if filename_lower in fname_lower:
                            if ext_filter is None or ext in ext_filter:
                                full_path = os.path.join(dirpath, fname)
                                result = _open_file(full_path)
                                return json.dumps(result)

            return json.dumps({"error": f"Could not find any file matching '{filename}' on your computer. Please check the name and try again."})

        except Exception as e:
            return json.dumps({"error": str(e)})

    def open_folder(self, folder_name: str):
        try:
            home = os.path.expanduser("~")
            known_folders = {
                "desktop": os.path.join(home, "Desktop"),
                "documents": os.path.join(home, "Documents"),
                "downloads": os.path.join(home, "Downloads"),
                "videos": os.path.join(home, "Videos"),
                "movies": os.path.join(home, "Videos"),
                "music": os.path.join(home, "Music"),
                "pictures": os.path.join(home, "Pictures"),
                "photos": os.path.join(home, "Pictures"),
                "home": home,
            }
            folder_lower = folder_name.lower().strip()
            target = known_folders.get(folder_lower)

            if not target:
                # Try to find the folder
                for root_dir in [home, "C:\\", "D:\\"]:
                    if not os.path.exists(root_dir):
                        continue
                    for dirpath, dirnames, _ in os.walk(root_dir):
                        depth = dirpath[len(root_dir):].count(os.sep)
                        if depth >= 4:
                            dirnames.clear()
                            continue
                        for d in dirnames:
                            if folder_lower in d.lower():
                                target = os.path.join(dirpath, d)
                                break
                        if target:
                            break
                    if target:
                        break

            if target and os.path.exists(target):
                subprocess.Popen(["explorer", target])
                return json.dumps({"status": "success", "message": f"Opened {folder_name} folder."})
            else:
                return json.dumps({"error": f"Folder '{folder_name}' not found."})
        except Exception as e:
            return json.dumps({"error": str(e)})

    def list_folder_contents(self, folder_name: str):
        try:
            home = os.path.expanduser("~")
            known_folders = {
                "desktop": os.path.join(home, "Desktop"),
                "documents": os.path.join(home, "Documents"),
                "downloads": os.path.join(home, "Downloads"),
                "videos": os.path.join(home, "Videos"),
                "movies": os.path.join(home, "Videos"),
                "music": os.path.join(home, "Music"),
                "pictures": os.path.join(home, "Pictures"),
            }
            folder_lower = folder_name.lower().strip()
            target = known_folders.get(folder_lower, os.path.join(home, folder_name))

            if not os.path.exists(target):
                return json.dumps({"error": f"Folder '{folder_name}' not found."})

            items = os.listdir(target)
            files = [f for f in items if os.path.isfile(os.path.join(target, f))]
            folders = [f for f in items if os.path.isdir(os.path.join(target, f))]

            return json.dumps({
                "status": "success",
                "folder": target,
                "files": files[:30],
                "folders": folders[:20],
                "message": f"Found {len(files)} files and {len(folders)} folders in {folder_name}."
            })
        except Exception as e:
            return json.dumps({"error": str(e)})

