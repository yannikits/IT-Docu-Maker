"""
Brücke: Aufzeichnung → Dokumentdaten

Event-Typen (Priorität):
  Primär:
    section  – Kapitelmarker        (description = Kapitelname)
    step     – Manueller Schritt    (description = Notiztext, screenshot_b64 = optional)
  Sekundär (Auto-Tracking, Fallback wenn keine Schritte vorhanden):
    click, scroll  – automatisch aufgezeichnete Aktionen

Dokumentstruktur (ohne KI):
  ## 1. Abschnittsname
    1. Schritt-Text
       [Bild]
    2. Schritt-Text
       ...
  ## 2. Nächster Abschnitt
    ...

Wenn keine section-Events vorhanden: alle Schritte in einem Abschnitt "Vorgehen".
Wenn keine step-Events vorhanden: Auto-Tracking-Events als Schritte verwenden.
"""

from datetime import datetime
from typing import List, Any


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _group_into_sections(events) -> list:
    """
    Gruppiert Events in Abschnitte.
    Rückgabe: Liste von {"heading": str|None, "steps": [...], "auto": [...]}
    """
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
        # key-Events werden ignoriert

    if current["steps"] or current["auto"] or current["heading"]:
        sections.append(current)

    return sections


# ---------------------------------------------------------------------------
# Markdown-Generierung (ohne KI)
# ---------------------------------------------------------------------------

def build_recording_markdown(events, include_screenshots: bool = True) -> str:
    """
    Konvertiert Aufzeichnung direkt in Markdown.

    Logik:
    • Manuelle Schritte (step) sind Primärinhalt.
    • Wenn kein manueller Schritt in einem Abschnitt: Auto-Events als Fallback.
    • Abschnitt-Marker (section) werden zu ## Überschriften.
    • Kein Abschnitt-Marker → Standardabschnitt »Vorgehen«.
    """
    sections = _group_into_sections(events)
    if not sections:
        return ""

    # Kein expliziter Abschnitt → Standardname
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
                # Auto-Tracking-Fallback
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


# ---------------------------------------------------------------------------
# Text-Beschreibung für KI-Provider
# ---------------------------------------------------------------------------

def build_ai_description(events, title: str) -> str:
    """
    Erstellt eine rein textuelle Beschreibung der Aufzeichnung für den KI-Provider.
    Keine Base64-Bilder (zu groß für API-Anfragen).
    Die KI soll daraus einen professionellen Dokumentationstext erzeugen.
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
            lines.append(f"\n=== ABSCHNITT: {sec["heading"]} ===")

        for ev in sec["steps"]:
            ss = " [Screenshot vorhanden]" if ev.screenshot_b64 else ""
            lines.append(f"  SCHRITT: {ev.description}{ss}")

        for ev in sec["auto"]:
            ts    = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
            label = {"click": "Klick", "scroll": "Scroll"}.get(ev.action_type, "Aktion")
            lines.append(f"  ({label} {ts}) {ev.description}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Screenshots in KI-Markdown einbetten
# ---------------------------------------------------------------------------

def inject_screenshots_into_markdown(ai_markdown: str, events) -> str:
    """
    Hängt alle Screenshots als Anhang ans KI-generierte Markdown an.
    (Die KI erhält keine Bilder – sie werden nachträglich eingebettet.)
    """
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


# ---------------------------------------------------------------------------
# Daten-Dict für Generator-Funktionen
# ---------------------------------------------------------------------------

def events_to_doc_data(
    events, title: str, template_id: str, fmt: str,
    markdown_content: str = "",
) -> dict:
    """Erstellt das data-Dict für generate_word / generate_excel / generate_pptx."""
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
    """Komplette Konvertierung ohne KI: Aufzeichnung → Markdown → doc-Datenstruktur."""
    markdown = build_recording_markdown(events, include_screenshots=True)
    return events_to_doc_data(events, title, template_id, fmt, markdown)
