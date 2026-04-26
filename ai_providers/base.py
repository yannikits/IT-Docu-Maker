"""
KI-Provider Abstraktion
Alle Provider implementieren dieselbe Schnittstelle.
"""

from abc import ABC, abstractmethod


class AIProvider(ABC):
    @abstractmethod
    def generate_document(
        self,
        description: str,
        title: str,
        fmt: str,
        template_id: str,
        chapters: list,
        aushang: bool,
        refs: list,
    ) -> str:
        raise NotImplementedError

    def _build_prompt(
        self, description, title, fmt, template_id, chapters, aushang, refs
    ) -> str:
        fmt_label = {
            "word":  "Word-Dokument (IT-Dokumentation / Schritt-für-Schritt-Anleitung)",
            "excel": "Excel-Tabelle",
            "ppt":   "PowerPoint-Präsentation",
        }.get(fmt, "IT-Dokumentation")

        tpl_label = {
            "intern":        "Internes Dokument",
            "extern":        "Externes Dokument",
            "kunde":         "Kundenanleitung",
            "netzwerk":      "Netzwerkdokumentation",
            "praesentation": "Präsentation",
        }.get(template_id, template_id)

        return f"""Du bist ein technischer Redakteur für IT-Dokumentationen.

Erstelle ein vollständiges {fmt_label} vom Typ \"{tpl_label}\" auf Deutsch im Markdown-Format.

Dokumenttitel: {title}

Die folgende Aufzeichnung enthält alle aufgezeichneten Klicks, Tastatureingaben und Notizen 
die ein Benutzer während der Durchführung eines IT-Prozesses gemacht hat.
Erstelle daraus eine professionelle, gut lesbare Schritt-für-Schritt-Anleitung.

Aufzeichnung:
{description}

Formatierungsregeln:
- Dokumenttitel als # Überschrift
- Kapitel als ## N. Kapitelname (orientiere dich an den ABSCHNITT-Markierungen)
- Unterkapitel als ### N.M Unterkapitelname
- Einzelne Schritte als nummerierte Liste mit präziser Beschreibung
- Hinweise und Warnungen als > ⚠️ Text
- Kein Abschnitt "Hinweise zur Vorlage"
- Keine Bilder, Screenshots oder Bildplatzhalter einfügen – Screenshots werden automatisch als Anhang angehängt
- Nur Markdown ausgeben, keine Erklärungen davor oder danach
- Wenn keine ABSCHNITT-Markierungen vorhanden, sinnvolle Kapitel aus dem Ablauf ableiten"""


def get_provider(config: dict) -> "AIProvider":
    provider_name = config.get("provider", "openai").lower()
    if provider_name == "openai":
        from ai_providers.openai_provider import OpenAIProvider
        return OpenAIProvider(
            api_key=config.get("openai_api_key", ""),
            model=config.get("openai_model", "gpt-4o"),
        )
    elif provider_name == "anthropic":
        from ai_providers.anthropic_provider import AnthropicProvider
        return AnthropicProvider(
            api_key=config.get("anthropic_api_key", ""),
            model=config.get("anthropic_model", "claude-sonnet-4-6"),
        )
    elif provider_name == "azure_openai":
        from ai_providers.azure_openai_provider import AzureOpenAIProvider
        return AzureOpenAIProvider(
            api_key=config.get("azure_api_key", ""),
            endpoint=config.get("azure_endpoint", ""),
            deployment=config.get("azure_deployment", "gpt-4o"),
            api_version=config.get("azure_api_version", "2024-02-01"),
        )
    else:
        raise ValueError(f"Unbekannter KI-Provider: {provider_name}")
