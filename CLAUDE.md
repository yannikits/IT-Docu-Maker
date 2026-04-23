# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the app

```bash
pip install -r requirements.txt
python main.py
# or on Windows:
start.bat
```

No build step, no test suite. The app is a standalone Tkinter desktop application.

## Architecture

IT-Docu-Maker merges two existing tools:
- **IT-Docu-Assistant** → `recorder/` (screen capture + event tracking)
- **WBI-Docu-Assist** → `generator/` + `ai_providers/` (document generation from templates)
- **`bridge/`** → new glue layer converting recorder output into generator input

### Data flow

```
User interaction
  → EventTracker (pynput) → list[ActionEvent]
  → bridge/recording_to_doc.py → data dict
  → generator/{word,excel,pptx}_generator.py → BytesIO document
```

### ActionEvent (recorder/event_tracker.py)

The central data unit. Key `action_type` values:
- `start` / `stop` – recording boundaries, ignored by bridge
- `section` – chapter marker; `description` becomes a heading
- `step` – primary content; has `description` (text note) and optional `screenshot_b64` (base64 JPEG)
- `click` / `scroll` – auto-tracked secondary events, used as fallback if no `step` events exist

### Bridge (bridge/recording_to_doc.py)

Converts `list[ActionEvent]` → `data` dict expected by generators.

- `_group_into_sections(events)` – splits events at `section` markers into `{heading, steps, auto}` dicts
- `build_recording_markdown()` – `step` events are primary; `click`/`scroll` are fallback
- `build_ai_description()` – text-only prompt for AI (no base64 images)
- `inject_screenshots_into_markdown(md, events)` – appends screenshots as `## Anhang` after AI-generated text
- `recording_to_doc_data_no_ai()` – full pipeline without AI
- `events_to_doc_data()` – assembles final `data` dict from pre-generated markdown

### Generator input format (`data` dict)

```python
{
    "titleSubject":    str,   # document title
    "titleTopic":      str,
    "markdownContent": str,   # full markdown text
    "chapters":        list,
    "screenshots":     list[{"b64": str, "caption": str}],
    # ... other template-specific keys
}
```

### AI providers (ai_providers/)

All providers implement `AIProvider.generate_document()` from `base.py`. `base._build_prompt()` builds the IT-documentation-specific prompt. Providers: `openai_provider.py`, `anthropic_provider.py`, `azure_openai_provider.py`. Each uses a 300s timeout and up to 3 retries on 5xx errors.

Provider is selected via `config.ini` (`provider = openai|anthropic|azure_openai`). Copy `config.ini.example` → `config.ini` and set `enabled = true` to activate AI export.

### Snipping tool (recorder/snipping.py)

`SnippingTool.capture_area(parent)` takes a full screenshot of all monitors (`mss.monitors[0]`), shows a darkened Tkinter overlay, and returns the selected region as base64 JPEG. HiDPI is handled via `scale_x = img_full.width / mon_w`.

The main window is hidden before snipping (`root.withdraw()` + 220ms delay) and restored afterwards (`root.deiconify()`).

## Templates

Template `.docx`/`.xlsx`/`.pptx` files must be placed in `vorlagen/`. They are not committed to this repo — copy them from `wbi-doku/vorlagen/` in WBI-Docu-Assist. The mapping from template ID to filename is defined in `TEMPLATE_FILES` in `main.py`.

## Key constraints

- Screenshots sent to AI providers are **never included** in the prompt — only the text description. Screenshots are appended as an appendix after AI generation via `inject_screenshots_into_markdown()`.
- The `_snipping_active` flag in `ITDocuMakerApp` prevents re-entrant snipping overlays if F9 is pressed during an open overlay.
- All UI updates from background threads must go through `root.after(0, ...)` — the AI export runs in a `daemon=True` thread.
