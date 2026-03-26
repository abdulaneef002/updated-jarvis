import webbrowser
import json
import urllib.parse
import requests
from typing import List, Dict, Any, Callable
from core.skill import Skill

QUICK_SITES = {
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
    "telegram": "https://web.telegram.org",
    "whatsapp": "https://web.whatsapp.com",
    "spotify": "https://open.spotify.com",
    "linkedin": "https://www.linkedin.com",
    "wikipedia": "https://www.wikipedia.org",
    "stackoverflow": "https://www.stackoverflow.com",
    "chatgpt": "https://chat.openai.com",
}

class WebSkill(Skill):
    @property
    def name(self) -> str:
        return "web_skill"

    def get_tools(self) -> List[Dict[str, Any]]:
        return [
            {
                "type": "function",
                "function": {
                    "name": "google_search",
                    "description": "Search Google for any query or topic",
                    "parameters": { "type": "object", "properties": { "search_term": {"type": "string"} }, "required": ["search_term"] }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "open_website",
                    "description": "Open a website by name or URL in the browser. Works for Gmail, YouTube, Telegram, WhatsApp, Netflix, Reddit, GitHub, Spotify, Instagram, Facebook, Twitter, and any other website.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "site_name": {"type": "string", "description": "Name of the website or full URL, e.g. 'gmail', 'youtube', 'telegram', 'https://example.com'"}
                        },
                        "required": ["site_name"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "youtube_search",
                    "description": "Search YouTube for a video and open the results",
                    "parameters": { "type": "object", "properties": { "query": {"type": "string"} }, "required": ["query"] }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "web_lookup",
                    "description": "Fetch a direct factual answer from the web for information questions, without opening a browser tab.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "query": {"type": "string", "description": "The factual question to answer, e.g. 'who is the prime minister of india'"}
                        },
                        "required": ["query"]
                    }
                }
            }
        ]

    def get_functions(self) -> Dict[str, Callable]:
        return {
            "google_search": self.google_search,
            "open_website": self.open_website,
            "youtube_search": self.youtube_search,
            "web_lookup": self.web_lookup,
        }

    def google_search(self, search_term):
        try:
            webbrowser.open(f"https://www.google.com/search?q={search_term}")
            return json.dumps({"status": "success", "message": f"Searching Google for {search_term}."})
        except Exception as e:
            return json.dumps({"error": str(e)})

    def open_website(self, site_name: str = ""):
        try:
            if not site_name:
                return json.dumps({"error": "Please tell me which website to open."})
            site_lower = site_name.lower().strip()
            # Check quick site map
            for key, url in QUICK_SITES.items():
                if key in site_lower or site_lower in key:
                    webbrowser.open(url)
                    return json.dumps({"status": "success", "message": f"Opened {key} in your browser."})
            # If it looks like a URL already
            if site_name.startswith("http://") or site_name.startswith("https://"):
                webbrowser.open(site_name)
                return json.dumps({"status": "success", "message": f"Opened {site_name} in your browser."})
            # Guess URL from name
            url = f"https://www.{site_lower.replace(' ', '')}.com"
            webbrowser.open(url)
            return json.dumps({"status": "success", "message": f"Opened {site_name} in your browser."})
        except Exception as e:
            return json.dumps({"error": str(e)})

    def youtube_search(self, query):
        try:
            webbrowser.open(f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}")
            return json.dumps({"status": "success", "message": f"Searching YouTube for {query}."})
        except Exception as e:
            return json.dumps({"error": str(e)})

    def web_lookup(self, query: str):
        query = (query or "").strip()
        if not query:
            return json.dumps({"status": "failed", "message": "Please provide a question to look up."})

        headers = {"User-Agent": "JARVIS/1.0"}

        # 1) Try Wikipedia search + summary for concise factual answers.
        try:
            search_resp = requests.get(
                "https://en.wikipedia.org/w/api.php",
                params={
                    "action": "query",
                    "list": "search",
                    "srsearch": query,
                    "format": "json",
                    "utf8": 1,
                    "srlimit": 1,
                },
                headers=headers,
                timeout=6,
            )
            if search_resp.ok:
                search_data = search_resp.json()
                hits = search_data.get("query", {}).get("search", [])
                if hits:
                    title = hits[0].get("title", "").strip()
                    if title:
                        summary_url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{urllib.parse.quote(title)}"
                        summary_resp = requests.get(summary_url, headers=headers, timeout=6)
                        if summary_resp.ok:
                            summary_data = summary_resp.json()
                            extract = (summary_data.get("extract") or "").strip()
                            if extract:
                                return json.dumps({
                                    "status": "success",
                                    "answer": extract,
                                    "source": f"Wikipedia: {title}",
                                })
        except Exception:
            pass

        # 2) Fallback to DuckDuckGo instant answers.
        try:
            ddg_resp = requests.get(
                "https://api.duckduckgo.com/",
                params={"q": query, "format": "json", "no_redirect": 1, "skip_disambig": 1},
                headers=headers,
                timeout=6,
            )
            if ddg_resp.ok:
                ddg_data = ddg_resp.json()
                abstract = (ddg_data.get("AbstractText") or "").strip()
                if abstract:
                    source = ddg_data.get("AbstractSource") or "DuckDuckGo"
                    return json.dumps({"status": "success", "answer": abstract, "source": source})

                related = ddg_data.get("RelatedTopics") or []
                for item in related:
                    text = (item.get("Text") or "").strip() if isinstance(item, dict) else ""
                    if text:
                        return json.dumps({"status": "success", "answer": text, "source": "DuckDuckGo"})
        except Exception:
            pass

        return json.dumps({
            "status": "failed",
            "message": "I could not fetch a reliable web answer right now. Please try again.",
        })

