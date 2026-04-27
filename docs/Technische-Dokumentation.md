# IT-Docu-Maker – Technische Dokumentation

## Inhaltsverzeichnis

1. [Projektübersicht](#projektübersicht)
2. [Architektur](#architektur)
3. [Verzeichnisstruktur](#verzeichnisstruktur)
4. [Abhängigkeiten](#abhängigkeiten)
5. [Konfiguration](#konfiguration)
6. [Modul: main.py](#modul-mainpy)
7. [Modul: recorder/](#modul-recorder)
   - [event_tracker.py](#event_trackerpy)
   - [screen_capture.py](#screen_capturepy)
   - [snipping.py](#snippingpy)
   - [annotation_editor.py](#annotation_editorpy)
8. [Modul: bridge/](#modul-bridge)
9. [Modul: generator/](#modul-generator)
   - [word_generator.py](#word_generatorpy)
   - [excel_generator.py](#excel_generatorpy)
   - [pptx_generator.py](#pptx_generatorpy)
10. [Modul: ai_providers/](#modul-ai_providers)
11. [Datenstrukturen](#datenstrukturen)
12. [Datenfluss](#datenfluss)
13. [Annotations-Editor: Interna](#annotations-editor-interna)
14. [Template-System](#template-system)
15. [Erweiterung und Beiträge](#erweiterung-und-beiträge)

---

## Projektübersicht

IT-Docu-Maker kombiniert zwei bestehende Projekte:

| Herkunft | Funktion |
|---------|---------|
| **IT-Docu-Assistant** | Aufnahme von Maus-/Tastaturereignissen und Screenshots |
| **WBI-Docu-Assist** | Generierung strukturierter Word-/Excel-/PowerPoint-Dokumente |

Die **Bridge-Schicht** (`bridge/`) koppelt beide Seiten: Sie wandelt rohe `ActionEvent`-Objekte in Markdown um, das die Generatoren verstehen.

**Technologie-Stack:**

| Kategorie | Technologie |
|-----------|------------|
| GUI | Python `tkinter` |
| Screenshot | `mss`, `Pillow` |
| Ereigniserfassung | `pynput` |
| Dokument-Export | `python-docx`, `openpyxl`, `python-pptx` |
| KI-Integration | `openai`, `anthropic` |
| Konfiguration | `configparser` + Umgebungsvariablen |

---

## Architektur

```
┌────────────────────────────────────────────────────────────┐
│  Tkinter GUI  (main.py – ITDocuMakerApp)                   │
│  • Hotkeys F8 / F9 / F10                                   │
│  • Template-Selektor                                       │
│  • Queue-basierter KI-Thread                               │
└──────┬────────────────────────────────────────────────────┘
       │
       ├── recorder/event_tracker.py   (pynput → ActionEvent)
       ├── recorder/screen_capture.py  (mss → Base64 JPEG)
       ├── recorder/snipping.py        (Tkinter-Overlay → Base64 JPEG)
       └── recorder/annotation_editor.py (PIL/Tk → Base64 JPEG)
                              │
                              ▼
              bridge/recording_to_doc.py
              • build_recording_markdown()
              • build_ai_description()
              • inject_screenshots_into_markdown()
              • recording_to_doc_data_no_ai()
                              │
               ┌──────────────┴──────────────┐
               │ (ohne KI)                   │ (mit KI)
               ▼                             ▼
          data-Dict                  ai_providers/
         direkt                  (OpenAI / Anthropic / Azure)
               │                             │
               └──────────────┬──────────────┘
                               ▼
                       generator/
               word_generator / excel_generator / pptx_generator
                               │
                           BytesIO
                               │
                      Speicher-Dialog → Datei
```

---

## Verzeichnisstruktur

```
IT-Docu-Maker/
├── main.py                   # Einstiegspunkt, Tkinter-GUI
├── config.py                 # Konfigurationsmanagement
├── config.ini.example        # Vorlage für KI-Einstellungen
├── requirements.txt          # Python-Abhängigkeiten
├── start.bat                 # Windows-Startskript
├── README.md                 # Projektbeschreibung (Deutsch)
├── CLAUDE.md                 # KI-Hilfskontext für Entwicklung
├── recorder/
│   ├── __init__.py
│   ├── event_tracker.py      # Maus-/Tastaturaufnahme
│   ├── screen_capture.py     # Vollbildschirm-Screenshot
│   ├── snipping.py           # Ausschnitt-Screenshot (Overlay)
│   └── annotation_editor.py  # Bild-Annotations-Editor
├── bridge/
│   ├── __init__.py
│   └── recording_to_doc.py   # ActionEvent → Markdown → Doc-Dict
├── generator/
│   ├── __init__.py
│   ├── word_generator.py     # .docx-Generierung
│   ├── excel_generator.py    # .xlsx-Generierung
│   └── pptx_generator.py     # .pptx-Generierung
├── ai_providers/
│   ├── __init__.py
│   ├── base.py               # Abstrakte Basisklasse + Prompt-Builder
│   ├── openai_provider.py    # OpenAI-GPT-4o-Integration
│   ├── anthropic_provider.py # Anthropic-Claude-Integration
│   └── azure_openai_provider.py # Azure-OpenAI-Integration
├── docs/
│   ├── Anleitung.md          # Benutzeranleitung
│   └── Technische-Dokumentation.md  # diese Datei
└── vorlagen/
    ├── README.md
    ├── Internes_Dokument_Vorlage.docx
    ├── Externes_Dokument_Vorlage.docx
    ├── Kundenanleitung_Vorlage.docx
    ├── Internes_Dokument_Vorlage.xlsx
    ├── Externes_Dokument_Vorlage.xlsx
    ├── Netzwerkdoku_Vorlage.xlsx
    └── Präsentationsvorlage_Vorlage.pptx
```

---

## Abhängigkeiten

| Paket | Version | Verwendung |
|-------|---------|-----------|
| `mss` | ≥ 9.0.0 | Multi-Monitor-Screenshots (plattformübergreifend) |
| `Pillow` | ≥ 10.0.0 | Bildbearbeitung, Annotationen, JPEG-Kodierung |
| `pynput` | ≥ 1.7.0 | Globale Tastatur- und Maus-Listener |
| `python-docx` | ≥ 1.1.0 | Word-Dokument-Generierung |
| `openpyxl` | ≥ 3.1.0 | Excel-Arbeitsmappen-Generierung |
| `python-pptx` | ≥ 0.6.23 | PowerPoint-Präsentationen |
| `lxml` | ≥ 5.0.0 | XML-Verarbeitung für Dokument-Formatting |
| `openai` | ≥ 1.0.0 | OpenAI- und Azure-OpenAI-API-Client |
| `anthropic` | ≥ 0.25.0, < 0.52.0 | Anthropic-Claude-API-Client |

---

## Konfiguration

### config.py

`config.py` kapselt den Zugriff auf `config.ini` und erlaubt Überschreibung durch Umgebungsvariablen.

```python
# Priorität: Umgebungsvariable > config.ini > Fallback-Wert
def get(section: str, key: str, fallback: str = "") -> str

def get_ai_config() -> dict  # Alle KI-Einstellungen als Dict
def ai_enabled() -> bool     # True, wenn KI aktiv UND API-Key vorhanden
```

**Umgebungsvariablen-Schema:** `ITDM_{SECTION}_{KEY}` (Großbuchstaben)

Beispiele:
- `ITDM_AI_ENABLED=true`
- `ITDM_AI_PROVIDER=anthropic`
- `ITDM_AI_ANTHROPIC_API_KEY=sk-ant-...`

### config.ini-Struktur

```ini
[ai]
enabled = false           # true/false
provider = openai         # openai | anthropic | azure

openai_api_key =
openai_model = gpt-4o

anthropic_api_key =
anthropic_model = claude-sonnet-4-6

azure_api_key =
azure_endpoint =
azure_deployment =
azure_api_version = 2024-02-01
```

---

## Modul: main.py

### Klasse `ITDocuMakerApp`

Hauptklasse der Anwendung. Erbt von keiner Tkinter-Klasse direkt – das `root`-Objekt wird im Konstruktor übergeben.

**Initialisierung:**

```python
def __init__(self, root: tk.Tk):
    # 1. Abhängigkeiten prüfen / installieren
    # 2. Tkinter-Fenster aufbauen (_build_ui)
    # 3. EventTracker, ScreenCapture, SnippingTool initialisieren
    # 4. Hotkeys F8/F9/F10 registrieren (pynput GlobalHotKeys)
    # 5. Queue für KI-Thread-Kommunikation erstellen
    # 6. Vorlagen registrieren (TEMPLATE_FILES-Dict)
    self._auto_screenshot_dir: str = ""  # optionaler Auto-Speicherordner
```

**Template-Registrierung:**

```python
TEMPLATE_FILES = {
    "intern":        ("Internes_Dokument_Vorlage.docx",    ..., "word"),
    "extern":        ("Externes_Dokument_Vorlage.docx",    ..., "word"),
    "kunde":         ("Kundenanleitung_Vorlage.docx",      ..., "word"),
    "netzwerk":      ("Netzwerkdoku_Vorlage.xlsx",         ..., "excel"),
    "intern_xl":     ("Internes_Dokument_Vorlage.xlsx",    ..., "excel"),
    "extern_xl":     ("Externes_Dokument_Vorlage.xlsx",    ..., "excel"),
    "praesentation": ("Präsentationsvorlage_Vorlage.pptx", ..., "ppt"),
}
```

**Screenshot-Auto-Speicherung:**

```python
def _choose_auto_screenshot_dir(self):
    # Öffnet einmalig einen Ordner-Auswahldialog
    # Speichert den Pfad in self._auto_screenshot_dir
    # Aktualisiert das Label in der UI (grün bei gesetztem Ordner)

def _do_snip(self, orig_geo):
    # ...nach Session-Speicherung in sessions/{title}/screenshots/:
    if self._auto_screenshot_dir:
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        auto_filename = f"{safe_title}_{ts}.png"
        pil_img.save(os.path.join(self._auto_screenshot_dir, auto_filename))
        # Statusleiste: "Screenshot N gespeichert: X  ·  Auto: Y"
```

Der Ordnerpfad wird nur für die laufende Sitzung im Speicher gehalten (nicht in `config.ini` persistiert).

**KI-Export (asynchron):**

Der Export mit KI läuft in einem Daemon-Thread, um die GUI nicht zu blockieren. Die Kommunikation erfolgt über eine `queue.Queue`:

```
Haupt-Thread:         Daemon-Thread:
_export_with_ai()  →  _ai_worker()
 startet Thread         ruft AI-Provider auf
 poll_queue()           legt Ergebnis in Queue
 ←── erkennt Ergebnis ──
 öffnet Speicher-Dialog
```

`poll_queue()` wird periodisch via `root.after(200, poll_queue)` aufgerufen.

---

## Modul: recorder/

### event_tracker.py

**Klasse `EventTracker`**

Wrappt `pynput`-Listener für Maus und Tastatur. Erstellt `ActionEvent`-Objekte.

```python
class EventTracker:
    def start(self): ...  # startet pynput-Listener
    def stop(self):  ...  # stoppt Listener, gibt ActionEvent-Liste zurück
```

**Throttling:** Mausklicks werden auf min. 0,5 s Abstand gedrosselt, um Doppelklicks zu filtern.

**Tastaturfilter:** Nur Sondertasten (Enter, Tab, Strg, Esc usw.) werden protokolliert; Buchstaben/Zeichen werden ignoriert.

---

### screen_capture.py

**Klasse `ScreenCapture`**

```python
def capture(quality: int = 85) -> str          # Vollbild → Base64-JPEG
def capture_thumbnail(quality=75, max_width=1280) -> str  # Skaliertes Bild
```

Nutzt `mss.mss()` für plattformübergreifende Multi-Monitor-Unterstützung. Thread-sicher via `threading.Lock`.

---

### snipping.py

**Klasse `SnippingTool`**

```python
def capture_area(parent: tk.Tk) -> Optional[str]  # Base64-JPEG oder None
```

Erstellt pro Monitor ein transparentes Tkinter-Overlay (abgedunkelter Hintergrund). Der Nutzer zieht einen Bereich auf; das Overlay wird geschlossen und der ausgewählte Bereich als Base64-JPEG zurückgegeben.

**HiDPI-Skalierung:** Vergleicht tatsächliche Screenshot-Pixelanzahl mit logischen Monitor-Abmessungen und berechnet einen Skalierungsfaktor zur korrekten Koordinatenmappung.

**Eintrittschutz:** `_snipping_active`-Flag in `main.py` verhindert mehrfaches Öffnen des Overlays.

---

### annotation_editor.py

Zentrales Modul für die Bildannotation. Siehe separaten Abschnitt [Annotations-Editor: Interna](#annotations-editor-interna).

---

## Modul: bridge/

### recording_to_doc.py

Entkoppelt die Aufnahmeschicht von der Generierungsschicht. Alle öffentlichen Funktionen erwarten eine `list[ActionEvent]`.

#### Funktionen

```python
def build_recording_markdown(events, include_screenshots=True) -> str
```
Erzeugt rohes Markdown mit eingebetteten Base64-Bildern (Inline-Data-URIs). Gliederung folgt `section`/`step`-Ereignissen.

```python
def build_ai_description(events, title: str) -> str
```
Erzeugt einen **Text-only**-Prompt für das KI-Modell (keine Bilder). Begrenzt auf ~40.000 Zeichen, um Token-Limits einzuhalten.

```python
def inject_screenshots_into_markdown(ai_markdown: str, events) -> str
```
Hängt an das KI-generierte Markdown einen `## Anhang`-Abschnitt mit allen Screenshots (Base64-eingebettet) an.

```python
def recording_to_doc_data_no_ai(events, title, template_id) -> dict
```
Vollständige Pipeline ohne KI: `build_recording_markdown` → `events_to_doc_data`.

```python
def events_to_doc_data(markdown, title, template_id) -> dict
```
Assembliert das finale Dict, das alle Generatoren als Eingabe erwarten (Felder: `titleSubject`, `titleTopic`, `markdownContent`, `format`, `template_id`).

---

## Modul: generator/

Alle Generatoren teilen dieselbe Schnittstelle:

```python
def generate_*(data: dict, template_path: str) -> BytesIO
```

`data` enthält mindestens:

| Schlüssel | Typ | Inhalt |
|-----------|-----|--------|
| `titleSubject` | `str` | Dokumenttitel |
| `titleTopic` | `str` | Untertitel / Thema |
| `markdownContent` | `str` | Vollständiges Markdown |
| `format` | `str` | `"word"` / `"excel"` / `"ppt"` |
| `template_id` | `str` | ID aus `TEMPLATE_FILES` |

### word_generator.py

1. Lädt `.docx`-Vorlage via `python-docx`.
2. Parst Markdown in Abschnitte und Schritte (Regex-basiert).
3. Mappt `#` → Dokumenttitel, `##` → Überschrift 1, `###` → Überschrift 2.
4. Fügt Base64-Bilder als eingebettete `Picture`-Runs ein (PNG-Format, in-memory).
5. Gibt `BytesIO`-Objekt zurück.

### excel_generator.py

Spezialisierte Handler je nach `template_id`:

- **`netzwerk`**: Tabellenspalten für Geräte, IP-Adressen, VLAN, Beschreibung.
- **`brief`**: Briefformat mit Absenderblock.
- **Generic Fallback**: Markdown-Listen → Tabellenzeilen, Überschriften → fettgedruckte Kopfzeilen.

### pptx_generator.py

1. Lädt `.pptx`-Vorlage.
2. Erstellt Titelfolie, Kapitelfolien (Überschriften), Inhaltsfolien (Schritte).
3. Bettet Base64-Bilder als `Picture`-Shapes ein.
4. Gibt `BytesIO` zurück.

---

## Modul: ai_providers/

### Abstrakte Basisklasse (base.py)

```python
class AIProvider(ABC):
    @abstractmethod
    def generate_document(
        self,
        description: str,   # Text-Aufzeichnung (kein Base64)
        title: str,
        fmt: str,           # "word" | "excel" | "ppt"
        template_id: str,
        chapters: list[str],
        aushang: str,
        refs: str,
    ) -> str                # Markdown-String
```

`_build_prompt()` erzeugt den System-Prompt:

- Rolle: IT-Technischer Redakteur
- Ausgabeformat: Markdown mit `#`/`##`/`###`-Hierarchie, nummerierte Listen
- Kontext: Vorlage, Format, Kapitelstruktur
- Hinweis: Keine Markdown-Code-Blöcke im Output

### Implementierungen

| Klasse | Modell | Besonderheiten |
|--------|--------|---------------|
| `OpenAIProvider` | `gpt-4o` | 3 Wiederholungen bei 5xx, Timeout 300 s, `max_tokens=4000` |
| `AnthropicProvider` | `claude-sonnet-4-6` | `max_tokens=8000`, kollabiert mehrfache Leerzeilen |
| `AzureOpenAIProvider` | konfigurierbar | Gleiche Logik wie OpenAI, Azure-Endpoint/-Deployment |

Alle Provider bereinigen die Antwort von Markdown-Code-Blöcken (` ```...``` `).

---

## Datenstrukturen

### ActionEvent

```python
@dataclass
class ActionEvent:
    timestamp: float           # Unix-Timestamp
    action_type: str           # Typ (siehe unten)
    description: str           # Menschenlesbarer Text
    screenshot_b64: Optional[str] = None  # Base64-JPEG oder None
    x: int = 0                 # Mauskoordinate X
    y: int = 0                 # Mauskoordinate Y
```

**`action_type`-Werte:**

| Wert | Auslöser |
|------|---------|
| `start` | Aufnahmebeginn |
| `stop` | Aufnahmeende |
| `click` | Mausklick |
| `scroll` | Mausrad |
| `key` | Sondertaste |
| `screenshot` | Manueller Screenshot (F9) |
| `section` | Kapitelmarkierung (F10) |
| `step` | Unterabschnitt |
| `note` | Freitextnotiz |

### Annotation-Dict (annotation_editor.py)

Jede Annotation im Editor wird als `dict` gespeichert:

```python
# Gemeinsame Felder
{
    "type":        str,           # "rect" | "arrow" | "text" | "number" | "blur"
    "color":       str,           # Hex-Farbcode, z. B. "#e74c3c"
    "canvas_item": int | None,    # Tkinter-Canvas-Item-ID
}

# rect
{"type": "rect",   "coords": (x1, y1, x2, y2)}      # Bildkoordinaten

# arrow
{"type": "arrow",  "coords": (x1, y1, x2, y2)}

# text
{"type": "text",   "coords": (ix, iy), "text": str}

# number
{"type": "number", "coords": (ix, iy),
 "number": int, "radius": int,
 "canvas_text": int | None}        # zweites Canvas-Item (Zifferntext)

# blur
{"type": "blur",   "coords": (x1, y1, x2, y2)}
```

Alle `coords`-Werte sind in **Bildkoordinaten** gespeichert (nicht Canvas-Pixel). Die Umrechnung erfolgt über `_to_img()` / `_to_cv()` mit dem Skalierungsfaktor `_s`.

---

## Datenfluss

### Ohne KI

```
ActionEvent[]
    └─ build_recording_markdown()
           └─ events_to_doc_data()
                  └─ generate_{word|excel|pptx}(data, template_path)
                         └─ BytesIO → Datei
```

### Mit KI

```
ActionEvent[]
    ├─ build_ai_description()     → Text-Prompt (kein Base64)
    │       └─ ai_provider.generate_document()
    │               └─ Markdown (KI-generiert)
    │                       └─ inject_screenshots_into_markdown()
    └─ events_to_doc_data(ai_markdown)
               └─ generate_{word|excel|pptx}(data, template_path)
                      └─ BytesIO → Datei
```

---

## Annotations-Editor: Interna

### Koordinatensystem

Der Editor arbeitet in zwei Koordinatensystemen:

| System | Beschreibung | Verwendung |
|--------|-------------|-----------|
| **Bildkoordinaten** | Pixel im Original-Screenshot | Speicherung in `ann["coords"]` |
| **Canvas-Koordinaten** | Pixel auf dem Tkinter-Canvas | Zeichnen und Interaktion |

Umrechnung via Skalierungsfaktor `_s` (Canvas-Breite / Bild-Breite):

```python
def _to_cv(self, ix, iy):   return int(ix * self._s), int(iy * self._s)
def _to_img(self, cx, cy):  return int(cx / self._s), int(cy / self._s)
```

### Canvas-Item-Binding

Statt geometrischer Hit-Tests verwendet der Editor **Tkinter-interne Item-Bindings** für Klick-Erkennung:

```python
self.cv.tag_bind(item_id, "<ButtonPress-1>", on_press)
```

Dadurch delegiert Tkinter die Hit-Detection intern (mit korrektem Anti-Aliasing für diagonale Pfeile). Das vermeidet fehleranfällige manuelle Koordinaten-Prüfungen.

### Doppelklick-Behandlung (Text/Nummer)

Da `<Double-Button-1>` auf Item-Ebene für Nummernkreise unzuverlässig ist (der Move-Handle liegt im Kreismittelpunkt und „frisst" den ersten Klick), ist das Doppelklick-Binding auf **Canvas-Ebene** registriert:

```python
self.cv.bind("<Double-Button-1>", self._on_dblclick)
```

`_on_dblclick` nutzt `_selected_idx` (bereits durch den Einfach-Klick gesetzt) um die zu bearbeitende Annotation zu identifizieren.

### Inline-Bearbeitung

`_start_edit(ann)` platziert ein Tkinter-`Entry`-Widget direkt auf dem Canvas über der Annotation:

```python
entry = tk.Entry(self.cv, ...)
self.cv.create_window(cx, cy, window=entry)
entry.bind("<Return>",  commit)
entry.bind("<Escape>",  cancel)
entry.bind("<FocusOut>", commit)
```

### Export-Pipeline

`_accept()` rendert alle Annotationen auf ein Pillow-`ImageDraw`-Objekt:

| Typ | Pillow-Methode |
|-----|---------------|
| `rect` | `draw.rectangle()` |
| `arrow` | `draw.line()` + `draw.polygon()` (Pfeilspitze) |
| `text` | `draw.text()` mit `ImageFont.truetype()` |
| `number` | `draw.ellipse()` + `draw.text(..., anchor="mm")` |
| `blur` | `img.crop()` + `ImageFilter.GaussianBlur()` + `img.paste()` |

Das Ergebnis wird als JPEG in einen `BytesIO`-Buffer kodiert und als Base64-String zurückgegeben.

### Werkzeug-Zustände

```
self._tool:       aktives Werkzeug ("select" | "rect" | "arrow" | "text" | "number" | "blur")
self._selected_idx: Index der selektierten Annotation (-1 = keine)
self._drag_mode:  aktueller Drag-Modus ("" | "move" | "resize_*")
self._drag_start: Canvas-Koordinaten beim Drag-Beginn
self._drag_orig:  Original-Koordinaten der Annotation beim Drag-Beginn
self._edit_entry: aktives Entry-Widget (None wenn kein Inline-Edit)
```

---

## Template-System

Dokumentvorlagen sind Office-Dateien (`.docx`, `.xlsx`, `.pptx`), die im Ordner `vorlagen/` liegen. Die Generatoren öffnen die Vorlage, befüllen sie mit Inhalt und geben das Ergebnis als `BytesIO` zurück – die Originaldatei wird nicht verändert.

### Template-ID-Mapping

```python
# main.py
TEMPLATE_FILES = {
    template_id: (filename, label, format_type)
}
```

Beim Export wird `template_id` an den Generator und die KI-Provider übergeben, sodass beide das Ausgabeformat kennen.

### KI-Prompt-Anpassung per Template

`AIProvider._build_prompt()` enthält ein Mapping von Template-ID auf menschenlesbare Bezeichnung:

```python
TEMPLATE_LABELS = {
    "intern":        "Internes Dokument",
    "extern":        "Externes Dokument",
    "kunde":         "Kundenanleitung",
    "netzwerk":      "Netzwerkdokumentation",
    "praesentation": "Präsentation",
    ...
}
```

Diese Bezeichnung wird in den System-Prompt eingebettet, damit das KI-Modell Ton und Struktur des Dokuments an die Zielgruppe anpasst.

---

## Erweiterung und Beiträge

### Neuen KI-Provider hinzufügen

1. Neue Datei `ai_providers/mein_provider.py` anlegen.
2. Klasse von `AIProvider` (aus `base.py`) ableiten.
3. `generate_document()`-Methode implementieren.
4. In `main.py` den Provider bei `provider == "mein_provider"` instanziieren.
5. In `config.ini.example` die neuen Konfigurationsfelder dokumentieren.

### Neues Annotationswerkzeug hinzufügen

1. In `annotation_editor.py` Button in `_build_toolbar()` eintragen.
2. In `_on_press()` den `elif self._tool == "mein_tool":` -Zweig ergänzen.
3. In `_redraw_annotation_item()` den Zeichencode für das Canvas hinzufügen.
4. In `_accept()` den PIL-Export-Code hinzufügen.
5. In `_hit_annotation()` die Trefferprüfung ergänzen.
6. In `_apply_transform()` und `_get_coords()` die Resize/Move-Logik ergänzen.

### Neue Dokumentvorlage hinzufügen

1. Office-Vorlagendatei in `vorlagen/` ablegen.
2. Eintrag in `TEMPLATE_FILES` in `main.py` ergänzen.
3. Falls nötig: spezialisierten Generator-Handler in `generator/` ergänzen.
4. KI-Prompt-Label in `ai_providers/base.py` eintragen.
