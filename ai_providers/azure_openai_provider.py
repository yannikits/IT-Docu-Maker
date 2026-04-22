"""
Azure OpenAI Provider
"""
import time
from ai_providers.base import AIProvider


class AzureOpenAIProvider(AIProvider):
    def __init__(self, api_key, endpoint, deployment, api_version="2024-02-01"):
        if not api_key or not endpoint:
            raise ValueError("Azure API-Key und Endpoint fehlen. Bitte in config.ini eintragen.")
        self.api_key     = api_key
        self.endpoint    = endpoint
        self.deployment  = deployment
        self.api_version = api_version

    def generate_document(self, description, title, fmt, template_id,
                          chapters, aushang, refs) -> str:
        try:
            from openai import AzureOpenAI, APIStatusError, APITimeoutError
        except ImportError:
            raise RuntimeError(
                "openai-Paket nicht installiert. Bitte 'pip install openai' ausfuehren."
            )
        client = AzureOpenAI(
            api_key=self.api_key,
            azure_endpoint=self.endpoint,
            api_version=self.api_version,
            timeout=300,
        )
        prompt = self._build_prompt(description, title, fmt, template_id, chapters, aushang, refs)

        last_exc = None
        for attempt in range(3):
            try:
                response = client.chat.completions.create(
                    model=self.deployment,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=4000,
                )
                text = response.choices[0].message.content or ""
                break
            except (APIStatusError, APITimeoutError) as e:
                last_exc = e
                if attempt < 2:
                    time.sleep(2 ** attempt * 2)
                    continue
                raise
        else:
            raise last_exc

        text = text.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])
        return text.strip()
