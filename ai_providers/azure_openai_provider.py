"""
Azure OpenAI Provider
"""

from ai_providers.base import AIProvider


class AzureOpenAIProvider(AIProvider):
    def __init__(self, api_key, endpoint, deployment, api_version="2024-02-01"):
        if not api_key or not endpoint:
            raise ValueError("Azure API-Key und Endpoint fehlen. Bitte in config.ini eintragen.")
        self.api_key = api_key
        self.endpoint = endpoint
        self.deployment = deployment
        self.api_version = api_version

    def generate_document(self, description, title, fmt, template_id,
                          chapters, aushang, refs) -> str:
        try:
            from openai import AzureOpenAI
        except ImportError:
            raise RuntimeError(
                "openai-Paket nicht installiert. Bitte 'pip install openai' ausführen."
            )
        client = AzureOpenAI(
            api_key=self.api_key,
            azure_endpoint=self.endpoint,
            api_version=self.api_version,
        )
        prompt = self._build_prompt(description, title, fmt, template_id, chapters, aushang, refs)
        response = client.chat.completions.create(
            model=self.deployment,
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
