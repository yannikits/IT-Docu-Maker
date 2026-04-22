"""
OpenAI / ChatGPT Provider
"""

from ai_providers.base import AIProvider


class OpenAIProvider(AIProvider):
    def __init__(self, api_key: str, model: str = "gpt-4o"):
        if not api_key:
            raise ValueError("OpenAI API-Key fehlt. Bitte in config.ini eintragen.")
        self.api_key = api_key
        self.model = model

    def generate_document(self, description, title, fmt, template_id,
                          chapters, aushang, refs) -> str:
        try:
            from openai import OpenAI
        except ImportError:
            raise RuntimeError(
                "openai-Paket nicht installiert. Bitte 'pip install openai' ausführen."
            )
        client = OpenAI(api_key=self.api_key)
        prompt = self._build_prompt(description, title, fmt, template_id, chapters, aushang, refs)
        response = client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=4000,
        )
        text = response.choices[0].message.content or ""
        text = text.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])
        return text.strip()
