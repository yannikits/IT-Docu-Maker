"""
Anthropic Claude Provider
"""

from ai_providers.base import AIProvider


class AnthropicProvider(AIProvider):
    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6"):
        if not api_key:
            raise ValueError("Anthropic API-Key fehlt. Bitte in config.ini eintragen.")
        self.api_key = api_key
        self.model = model

    def generate_document(self, description, title, fmt, template_id,
                          chapters, aushang, refs) -> str:
        try:
            import anthropic
        except ImportError:
            raise RuntimeError(
                "anthropic-Paket nicht installiert. Bitte 'pip install anthropic' ausführen."
            )
        client = anthropic.Anthropic(api_key=self.api_key)
        prompt = self._build_prompt(description, title, fmt, template_id, chapters, aushang, refs)
        message = client.messages.create(
            model=self.model,
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}],
        )
        text = message.content[0].text or ""
        text = text.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])
        return text.strip()
