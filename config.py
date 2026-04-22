"""
Konfigurationsverwaltung für IT-Docu-Maker.
Liest config.ini aus dem Projektordner.
Umgebungsvariablen überschreiben config.ini (für Server-Deployments).
"""

import os
import configparser

_CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.ini")
_config = configparser.ConfigParser()
_config.read(_CONFIG_FILE, encoding="utf-8")


def get(section: str, key: str, fallback: str = "") -> str:
    env_key = f"ITDM_{section.upper()}_{key.upper()}"
    env_val = os.environ.get(env_key)
    if env_val:
        return env_val
    return _config.get(section, key, fallback=fallback)


def get_ai_config() -> dict:
    return {
        "provider":          get("ai", "provider", "openai"),
        "openai_api_key":    get("ai", "openai_api_key", ""),
        "openai_model":      get("ai", "openai_model", "gpt-4o"),
        "anthropic_api_key": get("ai", "anthropic_api_key", ""),
        "anthropic_model":   get("ai", "anthropic_model", "claude-sonnet-4-6"),
        "azure_api_key":     get("ai", "azure_api_key", ""),
        "azure_endpoint":    get("ai", "azure_endpoint", ""),
        "azure_deployment":  get("ai", "azure_deployment", "gpt-4o"),
        "azure_api_version": get("ai", "azure_api_version", "2024-02-01"),
        "enabled":           get("ai", "enabled", "false").lower() == "true",
    }


def ai_enabled() -> bool:
    cfg = get_ai_config()
    if not cfg["enabled"]:
        return False
    provider = cfg["provider"]
    if provider == "openai" and not cfg["openai_api_key"]:
        return False
    if provider == "anthropic" and not cfg["anthropic_api_key"]:
        return False
    if provider == "azure_openai" and not (cfg["azure_api_key"] and cfg["azure_endpoint"]):
        return False
    return True
