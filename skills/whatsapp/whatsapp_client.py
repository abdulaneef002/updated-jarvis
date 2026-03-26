import time
import urllib.parse
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from .driver import WhatsAppDriver

class WhatsAppClient:
    def __init__(self):
        # We don't store driver here to avoid stale references
        pass

    def send_message(self, identifier, message):
        """
        Sends a message to a WhatsApp contact by name or phone number.
        Uses search bar when possible; falls back to URL for phone numbers.
        """
        from selenium.webdriver.common.keys import Keys
        
        try:
            # Get fresh driver instance
            driver = WhatsAppDriver.get_driver()
            wait = WebDriverWait(driver, 20)
            
            # 0. Determine if identifier looks like a phone number
            is_phone_number = identifier.strip().lstrip("+").isdigit()

            # 1. Check if we are already on WhatsApp Web and loaded
            try:
                if "web.whatsapp.com" in driver.current_url:
                    search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')

                    print("WhatsApp already open. Using Search Bar for speed...")
                    search_box.clear()
                    search_box.send_keys(identifier)
                    search_box.send_keys(Keys.ENTER)

                    time.sleep(1)
                else:
                    raise Exception("Not on WhatsApp Web")
            except Exception:
                print("Opening WhatsApp Web...")
                if is_phone_number:
                    encoded_message = urllib.parse.quote(message)
                    url = f"https://web.whatsapp.com/send?phone={identifier}&text={encoded_message}"
                    driver.get(url)
                else:
                    driver.get("https://web.whatsapp.com")

            # 2. Ensure chat opens for name-based search
            if not is_phone_number:
                try:
                    search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
                    search_box.clear()
                    search_box.send_keys(identifier)
                    search_box.send_keys(Keys.ENTER)
                    time.sleep(1)
                except TimeoutException:
                    return "WhatsApp Web is not ready. Scan the QR code to sign in first."

            # 3. Wait for Message Box
            print("Waiting for chat to load...")
            textbox = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
            
            # Type message (if we used search, text isn't pre-filled)
            # If we used URL, it might be pre-filled, but it's safer to just type it again or check value.
            # actually, URL ?text=... pre-fills it. Search doesn't.
            # So let's just type it to be sure.
            
            # Check if text is present? 
            # Simplified: Just type it. If it was pre-filled via URL, we might duplicate? 
            # If we navigated via URL with text, we don't need to type.
            # Let's clean the box just in case, or simpler:
            
            # OPTIMIZATION:
            # If we used Search, we MUST type.
            # If we used URL, we MIGHT not need to, but typing is safer.
            
            textbox.send_keys(message)
            time.sleep(0.5)
            
            # 3. Click Send or Hit Enter
            print("Chat loaded. Attempting to send...")
            try:
                # Primary method: Click the send button (icon)
                send_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
                send_button.click()
            except Exception as e:
                print(f"Send button click failed. Trying ENTER key...")
                textbox.send_keys(Keys.ENTER)
                
            # 4. Wait a bit for send to process
            time.sleep(1) 
            
            return f"Message sent to {identifier}"

        except Exception as e:
            print(f"Error in send_message: {e}")
            return f"Failed to send message: {e}"
