import os
import json
import platform
import subprocess
from datetime import datetime
from typing import List, Dict, Any, Callable
from core.skill import Skill

class ScreenshotSkill(Skill):
    """Skill for taking screenshots on Windows/macOS/Linux."""
    
    def __init__(self):
        # Default screenshot directory
        self.screenshot_dir = os.path.expanduser("~/Desktop/JARVIC_Screenshots")
        # Create directory if it doesn't exist
        os.makedirs(self.screenshot_dir, exist_ok=True)
    
    @property
    def name(self) -> str:
        return "screenshot_skill"

    def get_tools(self) -> List[Dict[str, Any]]:
        return [
            {
                "type": "function",
                "function": {
                    "name": "take_screenshot",
                    "description": "Take a screenshot of the entire screen and save it to a file. Returns the path to the saved screenshot.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {
                                "type": "string",
                                "description": "Optional custom filename for the screenshot (without extension). If not provided, uses timestamp."
                            }
                        },
                        "required": []
                    }
                }
            }
        ]

    def get_functions(self) -> Dict[str, Callable]:
        return {
            "take_screenshot": self.take_screenshot
        }

    def take_screenshot(self, filename: str = None) -> str:
        """
        Take a screenshot using OS-specific methods.
        
        Args:
            filename: Optional custom filename (without extension)
            
        Returns:
            JSON string with status and filepath
        """
        try:
            # Generate filename if not provided
            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"screenshot_{timestamp}"
            
            # Ensure .png extension
            if not filename.endswith('.png'):
                filename += '.png'
            
            filepath = os.path.join(self.screenshot_dir, filename)
            
            # Use native screenshot tooling per OS.
            result = self._capture_with_platform_tool(filepath)
            
            if result == 0 and os.path.exists(filepath):
                return json.dumps({
                    "status": "success",
                    "message": f"Screenshot saved successfully",
                    "path": filepath
                })
            else:
                return json.dumps({
                    "status": "error",
                    "message": "Failed to capture screenshot"
                })
                
        except Exception as e:
            return json.dumps({
                "status": "error",
                "message": f"Screenshot error: {str(e)}"
            })

    def _capture_with_platform_tool(self, filepath: str) -> int:
        system = platform.system().lower()

        if system == "windows":
            ps_command = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "Add-Type -AssemblyName System.Drawing; "
                "$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; "
                "$bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height; "
                "$gfx = [System.Drawing.Graphics]::FromImage($bmp); "
                "$gfx.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size); "
                f"$bmp.Save('{filepath.replace("'", "''")}'); "
                "$gfx.Dispose(); $bmp.Dispose();"
            )
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_command],
                capture_output=True,
                text=True,
            )
            return completed.returncode

        if system == "darwin":
            completed = subprocess.run(["screencapture", "-x", filepath], capture_output=True, text=True)
            return completed.returncode

        # Linux fallback (requires gnome-screenshot on most distros)
        completed = subprocess.run(["gnome-screenshot", "-f", filepath], capture_output=True, text=True)
        return completed.returncode
