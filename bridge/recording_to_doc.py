"""
Brücke: Aufzeichnung → Dokumentdaten
"""

from datetime import datetime
from typing import List, Any

MAX_AI_DESCRIPTION_CHARS = 40_000  # ~10 000 tokens, safe for all providers


def _group_into_sections(events) -> list:
    sections = []
    current  = {"heading": None, "steps": [], "auto": []}

    for ev in events:
        if ev.action_type in ("start", "stop"):
            continue
        if ev.action_type == "section":
            if current["steps"] or current["auto"] or current["heading"]:
                sections.append(current)
            current = {"heading": ev.description, "steps": [], "auto": []}
        elif ev.action_type == "step":
            current["steps"].append(ev)
        elif ev.action_type in ("click", "scroll"):
            current["auto"].append(ev)

    if current["steps"] or current["auto"] or current["heading"]:
        sections.append(current)

    return sections


def build_recording_markdown(events, include_screenshots: bool = True) -> str:
    sections = _group_into_sections(events)
    if not sections:
        return ""

    if len(sections) == 1 and sections[0]["heading"] is None:
        sections[0]["heading"] = "Vorgehen"

    lines: List[str] = []

    for idx, sec in enumerate(sections, 1):
        heading = sec["heading"] or f"Abschnitt {idx}"
        lines.append(f"## {idx}. {heading}\n")

        content_events = sec["steps"] if sec["steps"] else sec["auto"]

        for snum, ev in enumerate(content_events, 1):
            if ev.action_type == "step":
                lines.append(f"{snum}. {ev.description}")
            else:
                ts    = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
                label = {"click": "Klick", "scroll": "Scrollen"}.get(ev.action_type, "Aktion")
                lines.append(f"{snum}. **{label}** ({ts}): {ev.description}")

            if include_screenshots and ev.screenshot_b64:
                alt = ev.description[:60].replace('"', "'")
                lines.append(
                    f"\n![{alt}](data:image/jpeg;base64,{ev.screenshot_b64})\n"
                )

        lines.append("")

    return "\n".join(lines)


def build_ai_description(events, title: str) -> str:
    """
    Erstellt eine rein textuelle Beschreibung der Aufzeichnung für den KI-Provider.
    Keine Base64-Bilder. Auto-Events werden nur verwendet wenn keine manuellen Schritte vorhanden.
    """
    lines = [
        "Folgende Bildschirmaufzeichnung soll als professionelle IT-Dokumentation aufbereitet werden.",
        "",
        f"Titel:         {title}",
        f"Aufgenommen:   {datetime.now().strftime('%d.%m.%Y')}",
        "",
        "Aufgezeichnete Schritte (in Reihenfolge):",
        "",
    ]

    sections = _group_into_sections(events)

    for sec in sections:
        if sec["heading"]:
            lines.append(f"\n=== ABSCHNITT: {sec['heading']} ===")

        # Only use auto-events as fallback when no manual steps exist in this section
        content_events = sec["steps"] if sec["steps"] else sec["auto"]

        for ev in content_events:
            if ev.action_type == "step":
                ss = " [Screenshot vorhanden]" if ev.screenshot_b64 else ""
                lines.append(f"  SCHRITT: {ev.description}{ss}")
            else:
                ts    = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
                label = {"click": "Klick", "scroll": "Scroll"}.get(ev.action_type, "Aktion")
                lines.append(f"  ({label} {ts}) {ev.description}")

    result = "\n".join(lines)

    if len(result) > MAX_AI_DESCRIPTION_CHARS:
        result = result[:MAX_AI_DESCRIPTION_CHARS]
        result += "\n\n[Aufzeichnung gekürzt – zu viele Schritte für die API-Anfrage.]"

    return result


def inject_screenshots_into_markdown(ai_markdown: str, events) -> str:
    screenshots = [
        (ev, i + 1)
        for i, ev in enumerate(events)
        if ev.screenshot_b64 and ev.action_type not in ("start", "stop")
    ]
    if not screenshots:
        return ai_markdown

    lines = [ai_markdown, "", "## Anhang: Screenshots der Aufzeichnung", ""]
    for ev, idx in screenshots:
        ts   = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
        desc = ev.description[:80].replace('"', "'")
        lines.append(f"### Screenshot {idx} – {ts}: {ev.description[:60]}")
        lines.append(
            f"\n![{desc}](data:image/jpeg;base64,{ev.screenshot_b64})\n"
        )
    return "\n".join(lines)


def events_to_doc_data(
    events, title: str, template_id: str, fmt: str,
    markdown_content: str = "",
) -> dict:
    parts = title.split(" - ", 1) if " - " in title else [title, ""]
    return {
        "titleSubject":    parts[0].strip(),
        "titleTopic":      parts[1].strip() if len(parts) > 1 else "",
        "template":        template_id,
        "format":          fmt,
        "markdownContent": markdown_content,
        "chapters":        [],
        "refs":            [],
        "aushang":         False,
    }


def recording_to_doc_data_no_ai(
    events, title: str, template_id: str, fmt: str
) -> dict:
    markdown = build_recording_markdown(events, include_screenshots=True)
    return events_to_doc_data(events, title, template_id, fmt, markdown)
