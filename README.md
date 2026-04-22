# IT-Docu-Maker

Verbindet **IT-Docu-Assistant** (Bildschirmaufzeichnung) mit **WBI-Docu-Assist** (Dokumentgenerierung).

Aufzeichnungen aus Text + Screenshots werden direkt in fertige Word-, Excel- oder PPTX-Dokumente nach WBI-Vorlage umgewandelt – mit oder ohne KI.

## Workflow

```
1. Aufnahme starten (F8)
   └─ Mausklicks, Tastatureingaben und Screenshots werden aufgezeichnet

2. Schritt-Marker setzen (F10)
   └─ Bezeichnung eingeben → wird zur Kapitelüberschrift im Dokument

3. Manuelle Screenshots (F9)
   └─ Wichtige Momente explizit festhalten

4. Aufnahme stoppen (F8)

5. Vorlage wählen und exportieren
   ├─ Ohne KI: Aufzeichnung direkt als strukturiertes Dokument
   └─ Mit KI:  KI erstellt professionellen Text, Screenshots als Anhang
```

## Installation

### Voraussetzungen

- Python 3.10 oder neuer
- Vorlagendateien aus WBI-Docu-Assist (siehe `vorlagen/README.md`)

### Schnellstart (Windows)

```bat
# Doppelklick auf start.bat
# oder:
pip install -r requirements.txt
python main.py
```

### Manuelle Installation

```bash
pip install -r requirements.txt
python main.py
```

## Vorlagen einrichten

1. Kopiere die `.docx`/`.xlsx`/`.pptx`-Dateien aus dem WBI-Docu-Assist  
   (Pfad: `wbi-doku/vorlagen/`) in den Ordner `vorlagen/` dieses Projekts.
2. Starte `main.py`

## KI-Funktionen aktivieren (optional)

1. Kopiere `config.ini.example` → `config.ini`
2. Trage deinen API-Key ein und setze `enabled = true`
3. Nach dem Stoppen der Aufnahme erscheint zusätzlich der Button **Exportieren (mit KI)**

Unterstützte KI-Provider: OpenAI (GPT-4o), Anthropic (Claude), Azure OpenAI

## Unterschied: Ohne KI vs. Mit KI

| | Ohne KI | Mit KI |
|---|---|---|
| **Geschwindigkeit** | Sofort | 5–30 Sek. (API-Aufruf) |
| **Inhalt** | Rohe Aufzeichnung strukturiert | Professionell ausformulierter Text |
| **Screenshots** | Inline im Dokument | Als Anhang |
| **Kosten** | Kostenlos | API-Kosten je Provider |
| **Empfohlen für** | Schnelle interne Protokolle | Saubere Kunden- und externe Dokumente |

## Hotkeys

| Taste | Funktion |
|-------|----------|
| `F8` | Aufnahme starten / stoppen |
| `F9` | Manuellen Screenshot erstellen |
| `F10` | Neuen Schritt / Abschnitt (Kapitelmarker) setzen |

## Projektstruktur

```
it-docu-maker/
├── main.py                    # Hauptanwendung (Tkinter UI)
├── recorder/
│   ├── screen_capture.py      # Bildschirmaufnahme (aus IT-Docu-Assistant)
│   └── event_tracker.py       # Maus/Tastatur-Tracker + ActionEvent
├── bridge/
│   └── recording_to_doc.py    # Konvertierung Aufzeichnung → Dokumentdaten
├── generator/
│   ├── word_generator.py      # Word-Export (aus WBI-Docu-Assist)
│   ├── excel_generator.py     # Excel-Export (aus WBI-Docu-Assist)
│   └── pptx_generator.py      # PPTX-Export (aus WBI-Docu-Assist)
├── ai_providers/
│   ├── base.py                # Abstrakte KI-Basisklasse + IT-spezifischer Prompt
│   ├── anthropic_provider.py  # Claude-Provider
│   ├── openai_provider.py     # GPT-Provider
│   └── azure_openai_provider.py
├── vorlagen/                  # Vorlagendateien hier ablegen (.docx/.xlsx/.pptx)
├── config.py                  # Konfigurationsverwaltung
├── config.ini.example         # Vorlage für KI-Konfiguration
├── requirements.txt
└── start.bat                  # Windows-Startskript
```

## Herkunft der Komponenten

| Komponente | Quelle |
|---|---|
| `recorder/` | IT-Docu-Assistant |
| `generator/` | WBI-Docu-Assist |
| `ai_providers/` | WBI-Docu-Assist (Prompt angepasst für IT-Anleitungen) |
| `bridge/` | Neu – verbindet beide Programme |
| `main.py` | Neu – kombiniertes UI |
