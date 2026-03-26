import os
import json
import re
import inspect
from datetime import date
from groq import Groq
from core.registry import SkillRegistry
from core.system_controller import SystemController

class JarvisEngine:
    def __init__(self, registry: SkillRegistry):
        self.registry = registry
        self.controller = SystemController()
        self.client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
        self.model_name = "llama-3.3-70b-versatile"
        
        today = date.today().strftime("%B %d, %Y")
        self.system_instruction = (
            f"You are Jarvis, a voice-controlled AI assistant running on Windows. "
            f"Today's date is {today}. Use this to answer any time-related questions accurately. "
            "Your responses will be SPOKEN ALOUD by a text-to-speech engine, so always respond in plain conversational English. "
            "NEVER use markdown formatting (no **, ##, -, *, backticks, bullet points). "
            "Never output JSON in normal responses. "
            "If the user asks for code, provide short runnable code directly. "
            "For factual questions, rely on the web context provided to you when available, otherwise answer from your knowledge. "
            "After using a tool, give a SHORT spoken confirmation, e.g. 'Done, I opened Telegram for you.' or 'I found the file and opened it.' "
            "If a tool returns an error, explain it simply in one sentence. "
            "Keep all responses under 2 sentences. Be natural and conversational like a voice assistant."
        )

    def _normalize_prompt_for_routing(self, text: str) -> str:
        q = (text or "").strip()
        # Remove common prefixes from speech/UI intermediates.
        q = re.sub(r"^\s*\d+\s*[\).:-]\s*", "", q, flags=re.IGNORECASE)
        q = re.sub(r"^\s*(thinking|thought|query|question)\s*:\s*", "", q, flags=re.IGNORECASE)
        return q.strip()

    def _format_retry_wait(self, error_str: str) -> str | None:
        # Example fragment: "Please try again in 7m0.768s"
        match = re.search(r"try again in\s+(\d+)m([\d.]+)s", error_str, flags=re.IGNORECASE)
        if not match:
            return None
        try:
            minutes = int(match.group(1))
            seconds = int(float(match.group(2)))
            if minutes <= 0:
                return f"{seconds} seconds"
            if seconds <= 0:
                return f"{minutes} minutes"
            return f"{minutes} minutes and {seconds} seconds"
        except Exception:
            return None

    def _is_coding_request(self, text: str) -> bool:
        q = self._normalize_prompt_for_routing(text).lower()
        if not q:
            return False
        code_keywords = (
            "code", "python", "java", "javascript", "c++", "c#", "program", "script", "function", "algorithm"
        )
        request_words = ("write", "give", "create", "make", "show")
        return any(k in q for k in code_keywords) and any(w in q for w in request_words)

    def _is_time_sensitive_question(self, text: str) -> bool:
        q = self._normalize_prompt_for_routing(text).lower()
        if not q:
            return False
        if re.search(r"\b(202[3-9]|203\d)\b", q):
            return True
        time_words = (
            "current", "currently", "latest", "today", "now", "this year", "recent",
            "who is the", "who are the", "what is the", "which is the",
            "president", "prime minister", "ceo", "captain", "captain of",
            "champion", "winner", "won", "score", "rank", "ranked", "richest",
            "most popular", "best", "top",
        )
        return any(w in q for w in time_words)

    def _is_factual_question(self, text: str) -> bool:
        q = self._normalize_prompt_for_routing(text).lower()
        if not q:
            return False
        if not re.match(r"^(who|what|when|where|which|whom|whose|define|tell\s+me\s+about)\b", q):
            return False
        # If user is asking for an action, let command/tool path handle it.
        action_words = ("open", "play", "launch", "search", "create", "delete", "run", "execute")
        return not any(word in q for word in action_words)

    def _direct_answer_without_tools(self, user_prompt: str) -> str:
        try:
            today = date.today().strftime("%B %d, %Y")
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            f"Today's date is {today}. "
                            "Answer the user directly in one or two short sentences using your up-to-date knowledge. "
                            "Do not call tools. Do not output JSON. "
                            "If the question asks 'who is', start with the person's name first. "
                            "If there are multiple formats (for example cricket formats), state that clearly in one sentence."
                        ),
                    },
                    {"role": "user", "content": user_prompt},
                ],
                max_tokens=120,
            )
            return (response.choices[0].message.content or "").strip() or "I could not find that right now."
        except Exception:
            return "I could not fetch that answer right now. Please try again."

    def _direct_code_without_tools(self, user_prompt: str) -> str:
        q = (user_prompt or "").strip().lower()

        # Fast deterministic response for the common request.
        if "python" in q and ("add" in q or "sum" in q) and ("two" in q or "2" in q):
            return "Here is simple Python code: a = 1; b = 2; print(a + b)"

        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a coding assistant. Provide a short, runnable code answer. "
                            "If the user request is ambiguous, provide a simple beginner example in the requested language. "
                            "Do not call tools. Do not output JSON. "
                            "Do not use markdown or code fences; return plain text only."
                        ),
                    },
                    {"role": "user", "content": user_prompt},
                ],
                max_tokens=220,
            )
            return (response.choices[0].message.content or "").strip() or "I could not generate the code right now."
        except Exception:
            return "I could not generate the code right now. Please try again."

    def _try_web_lookup_answer(self, user_prompt: str) -> str | None:
        web_lookup = self.registry.get_function("web_lookup")
        if not web_lookup:
            return None
        try:
            raw = web_lookup(query=user_prompt)
            payload = json.loads(raw) if isinstance(raw, str) else raw
            if isinstance(payload, dict) and payload.get("status") == "success":
                answer = (payload.get("answer") or payload.get("message") or "").strip()
                if not answer:
                    return None

                # If user asks for winner/champion and lookup text does not contain that signal,
                # avoid speaking an unrelated tournament description.
                q = (user_prompt or "").lower()
                asks_winner = any(k in q for k in ("won", "winner", "champion"))
                if asks_winner and not any(k in answer.lower() for k in ("won", "winner", "champion")):
                    return None

                # Keep answers short and speech-friendly.
                short = answer.split(".")[0].strip()
                if short:
                    answer = short + "."
                if len(answer) > 220:
                    answer = answer[:217].rstrip() + "..."

                source = (payload.get("source") or "").strip()
                if source:
                    return f"{answer} Source: {source}."
                return answer
        except Exception:
            return None
        return None

    def _build_recovered_tool_args(self, func_name: str, raw_args: str, user_prompt: str, function_to_call) -> dict:
        try:
            parsed = json.loads(raw_args) if raw_args else {}
            if not isinstance(parsed, dict):
                parsed = {}
        except Exception:
            parsed = {}

        # Tool-specific fallbacks for common failed generations.
        lowered_prompt = self._normalize_prompt_for_routing(user_prompt)
        if func_name == "google_search" and "search_term" not in parsed:
            parsed["search_term"] = lowered_prompt
        elif func_name == "open_website" and "site_name" not in parsed:
            parsed["site_name"] = lowered_prompt
        elif func_name == "youtube_search" and "query" not in parsed:
            parsed["query"] = lowered_prompt
        elif func_name == "web_lookup" and "query" not in parsed:
            parsed["query"] = lowered_prompt

        # Generic fallback: fill required string params from the user prompt.
        try:
            sig = inspect.signature(function_to_call)
            for name, param in sig.parameters.items():
                if name in parsed:
                    continue
                if param.default is inspect._empty:
                    parsed[name] = lowered_prompt
        except Exception:
            pass

        return parsed

    def run_conversation(self, user_prompt: str) -> str:
        normalized_prompt = self._normalize_prompt_for_routing(user_prompt)
        controller_result = self.controller.handle_command(normalized_prompt)
        if controller_result.get("intent") != "unknown" or controller_result.get("action") == "clarification_required":
            return json.dumps(controller_result)

        if self._is_coding_request(normalized_prompt):
            return self._direct_code_without_tools(normalized_prompt)

        if self._is_factual_question(normalized_prompt):
            q = (normalized_prompt or "").lower()
            asks_winner = any(k in q for k in ("won", "winner", "champion"))

            # Always try live web first for ALL factual questions so answers are up-to-date.
            looked_up = self._try_web_lookup_answer(normalized_prompt)
            if looked_up:
                return looked_up

            if asks_winner and self._is_time_sensitive_question(normalized_prompt):
                return "I could not verify the winner from reliable sources right now. Please ask again in a little while."

            # Fall back to model knowledge with today's date already injected in system prompt.
            direct = self._direct_answer_without_tools(normalized_prompt)
            if direct and "could not" not in direct.lower() and "try again" not in direct.lower():
                return direct

            return "I could not find a reliable answer right now. Please try again."

        messages = [
            {"role": "system", "content": self.system_instruction},
            {"role": "user", "content": user_prompt}
        ]
        
        try:
            tools_schema = self.registry.get_tools_schema()
            # If no tools are loaded, don't pass tools argument (or pass empty? Groq might handle it)
            # Better to pass None if empty to avoid api error if specific models dislike empty tool lists?
            # Actually, let's pass it if it exists.
            
            completion_kwargs = {
                "model": self.model_name,
                "messages": messages,
                "max_tokens": 200
            }
            
            if tools_schema:
                completion_kwargs["tools"] = tools_schema
                completion_kwargs["tool_choice"] = "auto"
            
            response = self.client.chat.completions.create(**completion_kwargs)
        except Exception as e:
            # Handle tool_use_failed error from Groq
            error_str = str(e)
            if "tool_use_failed" in error_str and "failed_generation" in error_str:
                try:
                    # Extract failed generation from error message (it's inside the dict string)
                    # Pattern 1: <function=NAME{ARGS}</function> - with arguments
                    match = re.search(r"<function=(\w+)(?:.*?)(?=\{)(\{.*?\})<\/function>", error_str)
                    if match:
                        func_name = match.group(1)
                        func_args_str = match.group(2)
                        print(f"DEBUG: Recovered failed tool call: {func_name} with {func_args_str}")
                    else:
                        # Pattern 2: <function=NAME...anything...</function> - no valid JSON, assume no args
                        match = re.search(r"<function=(\w+).*?<\/function>", error_str)
                        if match:
                            func_name = match.group(1)
                            func_args_str = "{}"
                            print(f"DEBUG: Recovered failed tool call: {func_name} with empty args")
                    
                    if match:
                        function_to_call = self.registry.get_function(func_name)
                        if function_to_call:
                            try:
                                args = self._build_recovered_tool_args(func_name, func_args_str, user_prompt, function_to_call)
                                res = function_to_call(**args)
                                return str(res) # Return result directly as if it was the answer
                            except Exception as exec_e:
                                print(f"Recovered tool execution failed: {exec_e}")
                except Exception as parse_e:
                    print(f"Failed to recover tool call: {parse_e}")

                # Final fallback: answer directly without tools instead of failing the conversation.
                return self._direct_answer_without_tools(user_prompt)

            # Handle API rate limits with a clear spoken message.
            if "rate_limit_exceeded" in error_str or "Rate limit reached" in error_str:
                wait_text = self._format_retry_wait(error_str)
                if wait_text:
                    return f"I have hit the Groq token limit. Please try again in about {wait_text}."
                return "I have hit the Groq token limit right now. Please try again in a few minutes."

            print(f"Groq API Error: {e}")
            return "I am having trouble connecting to the brain, sir."

        response_message = response.choices[0].message
        tool_calls = response_message.tool_calls

        # CASE 1: AI wants to use a tool (Action)
        if tool_calls:
            print("DEBUG: Executing Tool...")
            messages.append(response_message)

            for tool_call in tool_calls:
                function_name = tool_call.function.name
                print(f"DEBUG: AI attempting to call: {function_name}")
                
                function_to_call = self.registry.get_function(function_name)
                
                if not function_to_call:
                    res = "Error: Tool not found."
                    print(f"DEBUG: Tool {function_name} not found in registry.")
                else:
                    try:
                        function_args = json.loads(tool_call.function.arguments)
                        print(f"DEBUG: Tool arguments: {function_args}")
                        
                        if function_args is None:
                            function_args = {}
                            
                        res = function_to_call(**function_args)
                        print(f"DEBUG: Tool Output: {str(res)[:100]}...") # Truncate for readability
                    except Exception as e:
                        res = f"Error executing tool: {e}"
                        print(f"DEBUG: Tool Execution Error: {e}")

                
                messages.append(
                    {
                        "tool_call_id": tool_call.id,
                        "role": "tool",
                        "name": function_name,
                        "content": str(res),
                    }
                )
            
            # Get final spoken response after tool runs
            # Remove tools arg for second call or keep it? Usually keep it in case it needs to chain.
            # But for simplicity let's just complete.
            second_response = self.client.chat.completions.create(
                model=self.model_name,
                messages=messages
            )
            return second_response.choices[0].message.content
        
        # CASE 2: AI wants to chat
        else:
            return response_message.content
