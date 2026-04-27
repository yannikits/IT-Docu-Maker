# IT-Docu-Maker – Benutzeranleitung

## Inhaltsverzeichnis

1. [Überblick](#überblick)
2. [Installation](#installation)
3. [Programmstart](#programmstart)
4. [Erster Start](#erster-start)
5. [Aufnahme starten und stoppen](#aufnahme-starten-und-stoppen)
6. [Struktur anlegen: Kapitel und Abschnitte](#struktur-anlegen-kapitel-und-abschnitte)
7. [Screenshots aufnehmen und bearbeiten](#screenshots-aufnehmen-und-bearbeiten)
8. [Annotationswerkzeuge](#annotationswerkzeuge)
9. [Auswahl-Werkzeug (Bearbeiten, Verschieben, Löschen)](#auswahl-werkzeug)
10. [Dokument exportieren](#dokument-exportieren)
11. [Vorlage auswählen](#vorlage-auswählen)
12. [KI-Konfiguration (optional)](#ki-konfiguration-optional)
13. [Tastenkürzel im Überblick](#tastenkürzel-im-überblick)
14. [Häufige Fragen (FAQ)](#häufige-fragen-faq)

---

## Überblick

IT-Docu-Maker zeichnet Ihren Bildschirm und Mausbewegungen auf, während Sie einen IT-Prozess durchführen, und erstellt daraus automatisch ein professionelles Dokument – wahlweise direkt oder mit KI-Unterstützung (GPT-4o, Claude, Azure).

**Typischer Arbeitsablauf:**

```
Aufnahme starten  →  Vorgang durchführen  →  Aufnahme stoppen  →  Exportieren
```

---

## Installation

### Voraussetzungen

- Windows 10/11 oder Linux
- Python 3.10 oder neuer
- Internetzugang für die erstmalige Abhängigkeitsinstallation

### Schritte

1. Repository klonen oder ZIP entpacken.
2. (Optional) Virtuelle Umgebung erstellen:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate        # Windows
   source .venv/bin/activate     # Linux/macOS
   ```
3. Abhängigkeiten installieren:
   ```bash
   pip install -r requirements.txt
   ```
   > Alternativ installiert das Programm fehlende Pakete beim ersten Start selbst.
4. Vorlagendateien in den Ordner `vorlagen/` kopieren (siehe Abschnitt [Vorlage auswählen](#vorlage-auswählen)).

---

## Programmstart

**Windows (empfohlen):**
```
start.bat
```

**Python direkt:**
```bash
python main.py
```

Nach dem Start öffnet sich das Hauptfenster.

---

## Erster Start

Beim ersten Start prüft IT-Docu-Maker automatisch, ob alle Abhängigkeiten installiert sind, und installiert fehlende Pakete nach. Dieser Vorgang dauert einmalig 1–2 Minuten und erfordert eine Internetverbindung.

---

## Aufnahme starten und stoppen

| Aktion | Schaltfläche | Tastenkürzel |
|--------|-------------|--------------|
| Aufnahme starten | **Aufnahme starten** | `F8` |
| Aufnahme stoppen | **Aufnahme stoppen** | `F8` |

Während der Aufnahme werden folgende Ereignisse automatisch protokolliert:

- **Mausklicks** (links, rechts, mittel) mit Koordinaten
- **Mausrollen** (oben/unten) mit Koordinaten
- **Tastendrücke** (Sondertasten wie Enter, Tab, Strg usw.)
- **Manuelle Screenshots** (siehe nächster Abschnitt)
- **Abschnittsmarkierungen** (Kapitel/Unterabschnitte)

Die Statusanzeige oben im Fenster zeigt an, wie viele Ereignisse bisher aufgezeichnet wurden.

---

## Struktur anlegen: Kapitel und Abschnitte

Um das spätere Dokument zu gliedern, können Sie während der Aufnahme jederzeit Abschnitte einfügen.

### Neues Kapitel anlegen

1. Text ins Feld **Kapitel** eingeben.
2. Schaltfläche **+ Kapitel** klicken oder `F10` drücken.

### Neuen Unterabschnitt anlegen

1. Text ins Feld **Unterabschnitt** eingeben.
2. Schaltfläche **+ Abschnitt** klicken.

Die eingegebenen Beschriftungen erscheinen als Überschriften im generierten Dokument.

---

## Screenshots aufnehmen und bearbeiten

### Screenshot mit Auswahlrahmen aufnehmen

1. Schaltfläche **Screenshot** klicken oder `F9` drücken.
2. Den gewünschten Bildschirmbereich mit der Maus aufziehen.
3. Maustaste loslassen – der Bereich wird ausgeschnitten und der **Annotations-Editor** öffnet sich.

### Im Annotations-Editor

Der Editor zeigt das aufgenommene Bild. Sie können es nun mit den Werkzeugen in der linken Werkzeugleiste kommentieren und anschließend mit **Übernehmen** speichern oder mit **Abbrechen** verwerfen.

---

## Annotationswerkzeuge

Die Werkzeugleiste befindet sich links im Annotations-Editor.

| Symbol | Werkzeug | Funktion |
|--------|---------|---------|
| `↖ Auswahl` | Auswahl | Annotationen auswählen, verschieben, skalieren oder löschen |
| `▭ Rahmen` | Rahmen | Rechteck auf dem Bild zeichnen |
| `→ Pfeil` | Pfeil | Linie/Pfeil von Punkt A zu Punkt B ziehen |
| `T Text` | Text | Klicken und Text eingeben |
| `① Nr.` | Nummer | Nummerierte Kreise platzieren (1, 2, 3 …) |
| `⊘ Blur` | Blur | Bereich unscharf machen (z. B. für sensible Daten) |

### Farbe wählen

Klicken Sie auf das **Farbfeld** in der Werkzeugleiste, um die Zeichenfarbe zu ändern. Die gewählte Farbe gilt für Rahmen, Pfeile, Text und Nummernkreise.

### Rahmen zeichnen

1. Werkzeug **▭ Rahmen** auswählen.
2. Maus auf dem Bild gedrückt halten und Rechteck aufziehen.
3. Maustaste loslassen – der Rahmen erscheint.

### Pfeil zeichnen

1. Werkzeug **→ Pfeil** auswählen.
2. Startpunkt klicken und zur Zielposition ziehen.
3. Maustaste loslassen – der Pfeil erscheint.

### Text einfügen

1. Werkzeug **T Text** auswählen.
2. An die gewünschte Stelle im Bild klicken.
3. Text im Eingabefenster eintippen und mit **OK** bestätigen.

### Nummerierungen einfügen

1. Werkzeug **① Nr.** auswählen.
2. Auf die gewünschte Stelle klicken – ein nummerierter Kreis wird platziert.
3. Weiteres Klicken platziert Kreise mit aufsteigender Nummer (1, 2, 3 …).

> **Tipp:** Mit dem Auswahl-Werkzeug können Sie die Nummer nachträglich durch Doppelklick ändern.

### Blur (Unschärfe) anwenden

1. Werkzeug **⊘ Blur** auswählen.
2. Den zu unscharf machenden Bereich aufziehen.
3. Der Bereich wird sofort unkenntlich gemacht.

---

## Auswahl-Werkzeug

Mit dem Werkzeug **↖ Auswahl** können Sie bereits gezeichnete Annotationen bearbeiten.

### Annotation auswählen

Klicken Sie auf eine vorhandene Annotation (Rahmen, Pfeil, Text, Nummerierung). Die Annotation wird mit Griffpunkten (kleinen Quadraten) hervorgehoben.

### Annotation verschieben

Klicken Sie auf die Annotation und ziehen Sie sie an die gewünschte Position.

### Annotation skalieren (Rahmen/Pfeil)

Ziehen Sie einen der Griffpunkte an den Ecken oder Enden der Annotation.

### Text oder Nummer bearbeiten (Doppelklick)

Doppelklicken Sie auf eine Text- oder Nummerierungs-Annotation, um den Inhalt inline zu bearbeiten:
- **Enter** oder Klick außerhalb: Änderung übernehmen.
- **Escape**: Bearbeitung abbrechen.

### Annotation löschen

1. Annotation mit dem Auswahl-Werkzeug auswählen.
2. Taste `Entf` (Delete) drücken.

---

## Dokument exportieren

Nach dem Stoppen der Aufnahme können Sie das Dokument in zwei Modi exportieren.

### Export ohne KI (sofort)

Klicken Sie auf **Export ohne KI**.

- Das Dokument wird unmittelbar aus den aufgezeichneten Ereignissen generiert.
- Der aufgezeichnete Ablauf wird als strukturierter Text in die gewählte Vorlage übertragen.
- Geeignet, wenn Sie die rohen Aufzeichnungsdaten bevorzugen oder keine KI konfiguriert ist.

### Export mit KI (empfohlen)

Klicken Sie auf **Export mit KI**.

- Die Aufzeichnung wird an das konfigurierte KI-Modell (OpenAI, Claude oder Azure) gesendet.
- Die KI erstellt einen professionellen, lesbaren Dokumenttext.
- Screenshots werden als Anhang an das Dokument angefügt.
- Dauer: ca. 5–30 Sekunden (abhängig von Modell und Aufzeichnungslänge).
- Fortschritt wird in der Statusleiste angezeigt.

> **Voraussetzung:** KI muss in `config.ini` aktiviert und ein gültiger API-Schlüssel hinterlegt sein (siehe [KI-Konfiguration](#ki-konfiguration-optional)).

### Speichern

Nach dem Export öffnet sich ein Datei-Speicherdialog. Wählen Sie Speicherort und Dateinamen. Das Dateiformat richtet sich nach der gewählten Vorlage (`.docx`, `.xlsx` oder `.pptx`).

---

## Vorlage auswählen

Klicken Sie auf **Vorlage wählen**, um die Dokumentvorlage zu wechseln.

| Vorlage | Format | Einsatz |
|---------|--------|---------|
| Internes Dokument (Word) | `.docx` | Interne IT-Dokumentation |
| Externes Dokument (Word) | `.docx` | Kundendokumentation (extern) |
| Kundenanleitung (Word) | `.docx` | Schritt-für-Schritt-Anleitungen für Kunden |
| Netzwerkdokumentation (Excel) | `.xlsx` | Netzwerkpläne und Gerätekonfigurationen |
| Internes Dokument (Excel) | `.xlsx` | Interne Dokumentation im Tabellenformat |
| Externes Dokument (Excel) | `.xlsx` | Externe Dokumentation im Tabellenformat |
| Präsentation (PPTX) | `.pptx` | Präsentationsfolien |

> Die Vorlagendateien müssen im Ordner `vorlagen/` liegen. Weitere Informationen dazu finden Sie in `vorlagen/README.md`.

---

## KI-Konfiguration (optional)

Um den Export mit KI zu nutzen, muss die Datei `config.ini` im Programmverzeichnis erstellt werden. Kopieren Sie dafür die mitgelieferte Vorlage:

```bash
cp config.ini.example config.ini
```

Öffnen Sie `config.ini` und tragen Sie Ihren API-Schlüssel ein:

### OpenAI (GPT-4o)

```ini
[ai]
enabled = true
provider = openai
openai_api_key = sk-...
openai_model = gpt-4o
```

### Anthropic Claude

```ini
[ai]
enabled = true
provider = anthropic
anthropic_api_key = sk-ant-...
anthropic_model = claude-sonnet-4-6
```

### Azure OpenAI

```ini
[ai]
enabled = true
provider = azure
azure_api_key = ...
azure_endpoint = https://IHRE-RESSOURCE.openai.azure.com/
azure_deployment = gpt-4o
azure_api_version = 2024-02-01
```

> **Alternative:** Statt `config.ini` können Umgebungsvariablen verwendet werden, z. B. `ITDM_AI_ENABLED=true` und `ITDM_AI_OPENAI_API_KEY=sk-...`.

---

## Tastenkürzel im Überblick

| Tastenkürzel | Funktion |
|-------------|---------|
| `F8` | Aufnahme starten / stoppen |
| `F9` | Manueller Screenshot (Auswahlrahmen) |
| `F10` | Aktuelles Kapitel eintragen |
| `Entf` | Ausgewählte Annotation löschen (im Annotations-Editor) |

---

## Häufige Fragen (FAQ)

**Das Programm startet nicht.**
Stellen Sie sicher, dass Python 3.10+ installiert ist und die Abhängigkeiten vollständig installiert wurden (`pip install -r requirements.txt`).

**„Export mit KI" ist ausgegraut.**
Die KI ist nicht aktiviert oder kein API-Schlüssel wurde hinterlegt. Prüfen Sie `config.ini` (Wert `enabled = true` und gültiger API-Key).

**Vorlagen fehlen / Exportfehler.**
Kopieren Sie die benötigten `.docx`/`.xlsx`/`.pptx`-Vorlagen in den Ordner `vorlagen/`. Genaue Dateinamen finden Sie in `vorlagen/README.md`.

**Annotationen lassen sich nicht verschieben.**
Stellen Sie sicher, dass das Werkzeug **↖ Auswahl** aktiv ist (nicht Rahmen oder Text).

**Der Blur-Bereich lässt sich nicht rückgängig machen.**
Blur-Bereiche sind permanent. Nutzen Sie **Abbrechen** im Editor, um den Screenshot ohne Änderungen zu verwerfen.

**Nummerierung beginnt nicht bei 1.**
Die Nummerierung setzt automatisch bei der nächsten freien Zahl an. Um von vorne zu beginnen, löschen Sie alle vorhandenen Nummernkreise mit dem Auswahl-Werkzeug + `Entf`.
