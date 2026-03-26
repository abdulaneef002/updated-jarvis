import os
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

class WhatsAppDriver:
    _instance = None
    _driver = None

    @classmethod
    def get_driver(cls):
        # Check if existing driver is valid
        if cls._driver is not None:
            try:
                # This will raise an exception if the window is closed/crashed
                _ = cls._driver.current_url
            except Exception:
                print("Driver found but unresponsive. cleanup...")
                try:
                    cls._driver.quit()
                except Exception:
                    pass
                cls._driver = None
                
        if cls._driver is None:
            cls._driver = cls._init_driver()
        return cls._driver

    @staticmethod
    def _init_driver():
        try:
            if os.name == "nt":
                return WhatsAppDriver._init_chrome()
            return WhatsAppDriver._init_safari()
        except Exception as e:
            print(f"Failed to initialize WhatsApp driver: {e}")
            raise e

    @staticmethod
    def _init_chrome():
        print("Initializing Chrome Driver...")
        options = Options()
        options.add_argument("--start-maximized")

        profile_dir = Path.home() / ".jarvis_whatsapp_profile"
        options.add_argument(f"--user-data-dir={profile_dir}")

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver

    @staticmethod
    def _init_safari():
        print("Initializing Safari Driver...")
        # Safari doesn't need webdriver_manager, it's built-in on macOS.
        # You just need to enable 'Allow Remote Automation' in Safari > Develop menu.
        driver = webdriver.Safari()
        driver.maximize_window()
        return driver
