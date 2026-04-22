"""
Brücke zwischen IT-Docu-Assistant (Aufzeichnung) und WBI-Docu-Assist (Dokumentgenerierung).

Zwei Modi:
  ohne KI  – recording_to_doc_data_no_ai()  → Aufzeichnung direkt als strukturiertes Markdown
  mit KI   – build_ai_description()          → Text-Beschreibung für KI-Provider
             inject_screenshots_into_markdown() → Screenshots als Anhang einbetten
"""

from datetime import datetime
from typing import List


def build_recording_markdown(events, include_screenshots: bool = True) -> str:
    """
    Konvertiert ActionEvent-Liste direkt in Markdown ohne KI.
    Notizen (F10) werden zu Kapitelüberschriften (## N. Name).
    Screenshots werden als eingebettete Base64-Bilder eingefügt.
    """
    sections = []
    current = {"heading": None, "events": []}

    for ev in events:
        if ev.action_type in ("start", "stop"):
            continue
        if ev.action_type == "note":
            if current["events"] or current["heading"]:
                sections.append(current)
            current = {"heading": ev.note, "events": []}
        else:
            current["events"].append(ev)

    if current["events"] or current["heading"]:
        sections.append(current)

    if not sections:
        return ""

    # Kein expliziter Abschnitt → alles unter einen Standard-Abschnitt
    if len(sections) == 1 and sections[0]["heading"] is None:
        sections[0]["heading"] = "Aufgezeichnete Schritte"

    lines = []
    for idx, section in enumerate(sections, 1):
        heading = section["heading"] or f"Abschnitt {idx}"
        lines.append(f"## {idx}. {heading}\n")

        step_num = 0
        for ev in section["events"]:
            step_num += 1
            ts = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
            label = {
                "click":      "Klick",
                "scroll":     "Scrollen",
                "key":        "Tastatureingabe",
                "screenshot": "Screenshot",
            }.get(ev.action_type, ev.action_type.capitalize())

            lines.append(f"{step_num}. **{label}** ({ts}): {ev.description}")

            if include_screenshots and ev.screenshot_b64:
                alt = ev.description[:60].replace('"', "'")
                lines.append(f"\n![{alt}](data:image/jpeg;base64,{ev.screenshot_b64})\n")

        lines.append("")

    return "\n".join(lines)


def build_ai_description(events, title: str) -> str:
    """
    Erstellt eine rein textuelle Beschreibung der Aufzeichnung für den KI-Provider.
    Keine Base64-Bilder – nur strukturierter Text.
    """
    lines = [
        "Folgende Bildschirmaufzeichnung soll als professionelle IT-Dokumentation aufbereitet werden.",
        "",
        f"Titel: {title}",
        f"Aufgenommen am: {datetime.now().strftime('%d.%m.%Y')}",
        "",
        "Aufgezeichnete Aktionen (chronologisch):",
        "",
    ]
    step = 0
    for ev in events:
        if ev.action_type in ("start", "stop"):
            continue
        step += 1
        ts    = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
        label = {
            "click":      "Mausklick",
            "scroll":     "Scrollen",
            "key":        "Tastatureingabe",
            "screenshot": "Screenshot",
            "note":       "Neuer Abschnitt",
        }.get(ev.action_type, ev.action_type)

        if ev.action_type == "note":
            lines.append(f"\n=== ABSCHNITT: {ev.note} ===")
        else:
            ss_hint = " [Screenshot vorhanden]" if ev.screenshot_b64 else ""
            lines.append(f"{step:3}. [{label}] ({ts}) {ev.description}{ss_hint}")

    return "\n".join(lines)


def inject_screenshots_into_markdown(ai_markdown: str, events) -> str:
    """
    Hängt alle Screenshots aus der Aufzeichnung als Anhang an das KI-generierte Markdown.
    Screenshots werden in der Reihenfolge ihres Auftretens eingebettet.
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
        lines.append(f"\n![{desc}](data:image/jpeg;base64,{ev.screenshot_b64})\n")

    return "\n".join(lines)


def events_to_doc_data(
    events, title: str, template_id: str, fmt: str,
    markdown_content: str = ""
) -> dict:
    """Erstellt das data-Dict für die Generator-Funktionen (generate_word / generate_excel / generate_pptx)."""
    parts = title.split(" - ", 1) if " - " in title else [title, ""]
    return {
        "titleSubject":   parts[0].strip(),
        "titleTopic":     parts[1].strip() if len(parts) > 1 else "",
        "template":       template_id,
        "format":         fmt,
        "markdownContent": markdown_content,
        "chapters":       [],
        "refs":           [],
        "aushang":        False,
    }


def recording_to_doc_data_no_ai(
    events, title: str, template_id: str, fmt: str
) -> dict:
    """Komplette Konvertierung ohne KI: Aufzeichnung → Markdown → doc-Datenstruktur."""
    markdown = build_recording_markdown(events, include_screenshots=True)
    return events_to_doc_data(events, title, template_id, fmt, markdown)
