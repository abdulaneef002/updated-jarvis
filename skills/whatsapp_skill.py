from core.skill import Skill
from skills.whatsapp.whatsapp_client import WhatsAppClient

class WhatsappSkill(Skill):
    """
    Skill for sending WhatsApp messages using Selenium and a local contact list.
    """
    
    def __init__(self):
        self.client = None # Lazy load the client

    @property
    def name(self):
        return "whatsapp_skill"
        
    def _get_client(self):
        if not self.client:
            self.client = WhatsAppClient()
        return self.client

    def get_tools(self):
        return [
            {
                "type": "function",
                "function": {
                    "name": "send_whatsapp_message",
                    "description": "Send a WhatsApp message to a specific person by name.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "name": {
                                "type": "string",
                                "description": "The name of the contact (e.g., 'Dad', 'Mom')."
                            },
                            "message": {
                                "type": "string",
                                "description": "The message to send."
                            }
                        },
                        "required": ["name", "message"],
                    },
                },
            }
        ]

    def get_functions(self):
        return {
            "send_whatsapp_message": self.send_whatsapp_message
        }

    def send_whatsapp_message(self, name, message):
        """
        Sends a WhatsApp message to a contact by name or phone number.
        """
        # Send message via client using name or number directly
        try:
            client = self._get_client()
            result = client.send_message(name, message)
            return result
        except Exception as e:
            return f"Error sending message: {e}"
