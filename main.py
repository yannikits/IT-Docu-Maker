#!/usr/bin/env python3
"""
IT-Docu-Maker
Kombiniert Bildschirmaufzeichnung (IT-Docu-Assistant) mit Dokumentgenerierung (WBI-Docu-Assist).

Workflow:
  1. Aufnahme starten (F8) - Bildschirmaktionen aufzeichnen
  2. Manuelle Screenshots (F9) und Schritt-/Kapitelmarker (F10) setzen
  3. Aufnahme stoppen (F8)
  4. Vorlage und Format wählen
  5. Dokument exportieren – ohne KI (direkte Konvertierung) oder mit KI (aufbereiteter Text)

Hotkeys:
  F8  - Aufnahme starten / stoppen
  F9  - Manuellen Screenshot erstellen
  F10 - Neuen Schritt / Abschnitt (Kapitelmarker) hinzufügen
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
from pathlib import Path
import queue

# ---------------------------------------------------------------------------
# Dependency check
# ---------------------------------------------------------------------------
REQUIRED = {"mss": "mss", "PIL": "Pillow", "pynput": "pynput"}
missing = []
for mod, pkg in REQUIRED.items():
    try:
        __import__(mod)
    except ImportError:
        missing.append(pkg)
if missing:
    print(f"Installiere fehlende Pakete: {', '.join(missing)}")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + missing)

from recorder.event_tracker import ActionEvent, EventTracker
from recorder.screen_capture import ScreenCapture

# ---------------------------------------------------------------------------
# Vorlage-Konfiguration
# ---------------------------------------------------------------------------
VORLAGEN_DIR = os.path.join(os.path.dirname(__file__), "vorlagen")

# (Dateiname, Anzeigename, Format)
TEMPLATE_FILES = {
    "intern":        ("Internes_Dokument_Vorlage.docx",       "Internes Dokument (Word)",       "word"),
    "extern":        ("Externes_Dokument_Vorlage.docx",       "Externes Dokument (Word)",       "word"),
    "kunde":         ("Kundenanleitung_Vorlage.docx",         "Kundenanleitung (Word)",         "word"),
    "netzwerk":      ("Netzwerkdoku_Vorlage.xlsx",            "Netzwerkdokumentation (Excel)",  "excel"),
    "intern_xl":     ("Internes_Dokument_Vorlage.xlsx",       "Internes Dokument (Excel)",      "excel"),
    "extern_xl":     ("Externes_Dokument_Vorlage.xlsx",       "Externes Dokument (Excel)",      "excel"),
    "praesentation": ("Präsentationsvorlage_Vorlage.pptx",   "Präsentation (PPTX)",            "ppt"),
}


def get_template_path(template_id: str) -> str:
    info = TEMPLATE_FILES.get(template_id)
    if not info:
        raise ValueError(f"Unbekannte Vorlage: {template_id}")
    filename, _, _ = info
    path = os.path.join(VORLAGEN_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Vorlagendatei nicht gefunden: {filename}\n\n"
            f"Bitte kopiere die Vorlagen aus dem WBI-Docu-Assist\n"
            f"(Ordner wbi-doku/vorlagen/) in den Ordner:\n{VORLAGEN_DIR}"
        )
    return path


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------

class ITDocuMakerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IT-Docu-Maker")
        self.root.resizable(False, False)

        self.events = []
        self.recording = False
        self.capture = ScreenCapture()
        self.tracker = EventTracker(self._on_tracked_event)
        self.event_queue = queue.Queue()

        self.capture_on_click  = tk.BooleanVar(value=True)
        self.capture_on_scroll = tk.BooleanVar(value=False)
        self.capture_on_key    = tk.BooleanVar(value=False)
        self.doc_title         = tk.StringVar(value=f"IT-Dokumentation {datetime.now().strftime('%Y-%m-%d')}")
        self.screenshot_quality = tk.IntVar(value=75)
        self.max_width          = tk.IntVar(value=1280)

        self._build_ui()
        self._setup_hotkeys()
        self.root.after(100, self._process_queue)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.root.configure(bg="#f3f2f1")
        pad = {"padx": 12, "pady": 5}

        # Titelleiste
        title_frame = tk.Frame(self.root, bg="#0078d4", pady=10)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="IT-Docu-Maker",
                 bg="#0078d4", fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT, padx=14)
        tk.Label(title_frame, text="Aufzeichnung  →  Fertiges Dokument",
                 bg="#0078d4", fg="#c7e0f4",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)

        # Status
        self.status_var = tk.StringVar(value="Bereit – Aufnahme starten, Aktionen durchführen, dann exportieren.")
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            bg="#f3f2f1", font=("Segoe UI", 10), anchor="w"
        )
        self.status_label.pack(fill=tk.X, **pad)

        # Aufnahme-Buttons
        btn_frame = tk.Frame(self.root, bg="#f3f2f1")
        btn_frame.pack(fill=tk.X, padx=12, pady=4)

        self.rec_btn = tk.Button(
            btn_frame, text="▶  Aufnahme starten (F8)",
            command=self.toggle_recording,
            bg="#107c10", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=10, pady=6, cursor="hand2"
        )
        self.rec_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.manual_btn = tk.Button(
            btn_frame, text="📷 Screenshot (F9)",
            command=self.manual_screenshot,
            state=tk.DISABLED,
            bg="#004578", fg="white",
            font=("Segoe UI", 10),
            relief=tk.FLAT, padx=10, pady=6, cursor="hand2"
        )
        self.manual_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.note_btn = tk.Button(
            btn_frame, text="✎ Schritt/Abschnitt (F10)",
            command=self.add_note,
            state=tk.DISABLED,
            bg="#8764b8", fg="white",
            font=("Segoe UI", 10),
            relief=tk.FLAT, padx=10, pady=6, cursor="hand2"
        )
        self.note_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Ereigniszähler
        self.counter_var = tk.StringVar(value="Ereignisse: 0  |  Screenshots: 0  |  Abschnitte: 0")
        tk.Label(
            self.root, textvariable=self.counter_var,
            bg="#f3f2f1", fg="#605e5c",
            font=("Segoe UI", 9)
        ).pack(**pad)

        # Aufnahme-Einstellungen
        rec_frame = ttk.LabelFrame(self.root, text="Aufnahme-Einstellungen", padding=8)
        rec_frame.pack(fill=tk.X, padx=12, pady=4)

        tk.Label(rec_frame, text="Titel:", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        tk.Entry(rec_frame, textvariable=self.doc_title, width=38,
                 font=("Segoe UI", 9)).grid(row=0, column=1, columnspan=3, sticky="ew")

        tk.Label(rec_frame, text="Screenshot bei:", font=("Segoe UI", 9)).grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        tk.Checkbutton(rec_frame, text="Klick",     variable=self.capture_on_click,  font=("Segoe UI", 9)).grid(row=1, column=1, sticky="w")
        tk.Checkbutton(rec_frame, text="Scrollen",  variable=self.capture_on_scroll, font=("Segoe UI", 9)).grid(row=1, column=2, sticky="w")
        tk.Checkbutton(rec_frame, text="Sondertaste", variable=self.capture_on_key,  font=("Segoe UI", 9)).grid(row=1, column=3, sticky="w")

        tk.Label(rec_frame, text="Qualität:", font=("Segoe UI", 9)).grid(
            row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Scale(rec_frame, from_=40, to=95, variable=self.screenshot_quality,
                  orient=tk.HORIZONTAL, length=140).grid(row=2, column=1, columnspan=2, sticky="ew")
        self.quality_label = tk.Label(rec_frame, font=("Segoe UI", 9))
        self.quality_label.grid(row=2, column=3, sticky="w")
        self.screenshot_quality.trace_add("write",
            lambda *_: self.quality_label.config(text=f"{self.screenshot_quality.get()} %"))
        self.quality_label.config(text=f"{self.screenshot_quality.get()} %")

        # Trennlinie
        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=12, pady=8)

        # Dokument-Einstellungen
        doc_frame = ttk.LabelFrame(self.root, text="Dokument-Einstellungen", padding=8)
        doc_frame.pack(fill=tk.X, padx=12, pady=4)

        tk.Label(doc_frame, text="Vorlage:", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        template_labels = [v[1] for v in TEMPLATE_FILES.values()]
        self.template_combo = ttk.Combobox(
            doc_frame, values=template_labels,
            state="readonly", width=36, font=("Segoe UI", 9)
        )
        self.template_combo.current(0)
        self.template_combo.grid(row=0, column=1, sticky="ew")

        # Export-Buttons
        exp_frame = tk.Frame(self.root, bg="#f3f2f1")
        exp_frame.pack(fill=tk.X, padx=12, pady=6)

        self.export_btn = tk.Button(
            exp_frame, text="💾  Exportieren (ohne KI)",
            command=self.export_document,
            state=tk.DISABLED,
            bg="#0078d4", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=12, pady=8, cursor="hand2"
        )
        self.export_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.ai_export_btn = tk.Button(
            exp_frame, text="🤖  Exportieren (mit KI)",
            command=self.export_with_ai,
            state=tk.DISABLED,
            bg="#5c2d91", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=12, pady=8, cursor="hand2"
        )
        self.ai_export_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            self.root,
            text="F8: Aufnahme starten/stoppen  |  F9: Screenshot  |  F10: Schritt / Kapitelmarker setzen",
            bg="#f3f2f1", fg="#605e5c", font=("Segoe UI", 8)
        ).pack(pady=(0, 8))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Hotkeys
    # ------------------------------------------------------------------

    def _setup_hotkeys(self):
        from pynput import keyboard as pkeyboard

        def on_press(key):
            try:
                k = key.name if hasattr(key, "name") else str(key)
            except AttributeError:
                k = str(key)
            k = k.lower().replace("key.", "")
            if k == "f8":  self.root.after(0, self.toggle_recording)
            elif k == "f9":  self.root.after(0, self.manual_screenshot)
            elif k == "f10": self.root.after(0, self.add_note)

        self._hk_listener = pkeyboard.Listener(on_press=on_press, daemon=True)
        self._hk_listener.start()

    # ------------------------------------------------------------------
    # Aufnahme-Steuerung
    # ------------------------------------------------------------------

    def toggle_recording(self):
        if self.recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self):
        self.recording = True
        self.events.clear()
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="start",
            description="Aufnahme gestartet"
        ))
        self.tracker.start()
        self.rec_btn.config(text="■  Aufnahme stoppen (F8)", bg="#a4262c")
        self.manual_btn.config(state=tk.NORMAL)
        self.note_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)
        self._set_status("🔴 Aufnahme läuft – Aktionen werden aufgezeichnet ...", "#a4262c")
        self._update_counter()

    def _stop_recording(self):
        self.recording = False
        self.tracker.stop()
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="stop",
            description="Aufnahme beendet"
        ))
        self.rec_btn.config(text="▶  Aufnahme starten (F8)", bg="#107c10")
        self.manual_btn.config(state=tk.DISABLED)
        self.note_btn.config(state=tk.DISABLED)
        if len(self.events) > 2:
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()
        count = len([e for e in self.events if e.action_type not in ("start", "stop")])
        self._set_status(
            f"✅ Gestoppt – {count} Ereignisse aufgezeichnet. Vorlage wählen und exportieren.",
            "#107c10"
        )
        self._update_counter()

    def _check_ai_available(self):
        try:
            import config as cfg
            if cfg.ai_enabled():
                self.ai_export_btn.config(state=tk.NORMAL)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Ereignisverarbeitung
    # ------------------------------------------------------------------

    def _on_tracked_event(self, action_type, description, x, y):
        self.event_queue.put((action_type, description, x, y))

    def _process_queue(self):
        while not self.event_queue.empty():
            try:
                action_type, description, x, y = self.event_queue.get_nowait()
                self._handle_event(action_type, description, x, y)
            except queue.Empty:
                break
        self.root.after(100, self._process_queue)

    def _handle_event(self, action_type, description, x, y):
        if not self.recording:
            return
        take_ss = (
            (action_type == "click"  and self.capture_on_click.get())  or
            (action_type == "scroll" and self.capture_on_scroll.get()) or
            (action_type == "key"    and self.capture_on_key.get())
        )
        ss_b64 = None
        if take_ss:
            ss_b64 = self.capture.capture_thumbnail(
                quality=self.screenshot_quality.get(),
                max_width=self.max_width.get()
            )
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type=action_type,
            description=description, screenshot_b64=ss_b64, x=x, y=y
        ))
        self._update_counter()

    # ------------------------------------------------------------------
    # Manuelle Aktionen
    # ------------------------------------------------------------------

    def manual_screenshot(self):
        if not self.recording:
            return
        ss_b64 = self.capture.capture_thumbnail(
            quality=self.screenshot_quality.get(),
            max_width=self.max_width.get()
        )
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="screenshot",
            description="Manueller Screenshot", screenshot_b64=ss_b64
        ))
        self._set_status("📷 Screenshot gespeichert.", "#004578")
        self._update_counter()

    def add_note(self):
        if not self.recording:
            return
        self.root.lift()
        note = simpledialog.askstring(
            "Schritt / Abschnitt",
            "Bezeichnung für diesen Schritt:\n(wird als Kapitelüberschrift im Dokument verwendet)",
            parent=self.root
        )
        if note and note.strip():
            ss_b64 = self.capture.capture_thumbnail(
                quality=self.screenshot_quality.get(),
                max_width=self.max_width.get()
            )
            self.events.append(ActionEvent(
                timestamp=time.time(), action_type="note",
                description=note.strip(), screenshot_b64=ss_b64,
                note=note.strip()
            ))
            self._set_status(f"✎ Schritt gespeichert: {note[:50]}", "#8764b8")
            self._update_counter()

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def _get_selected_template_id(self) -> str:
        idx = self.template_combo.current()
        return list(TEMPLATE_FILES.keys())[idx]

    def export_document(self):
        """Export ohne KI – direkte Konvertierung der Aufzeichnung."""
        if not self.events:
            messagebox.showwarning("Keine Daten", "Noch keine Aktionen aufgezeichnet.")
            return
        template_id = self._get_selected_template_id()
        _, _, fmt = TEMPLATE_FILES[template_id]
        try:
            template_path = get_template_path(template_id)
        except FileNotFoundError as e:
            messagebox.showerror("Vorlage fehlt", str(e))
            return

        from bridge.recording_to_doc import recording_to_doc_data_no_ai
        data = recording_to_doc_data_no_ai(
            self.events, self.doc_title.get(), template_id, fmt
        )
        self._run_export(data, template_path, fmt)

    def export_with_ai(self):
        """Export mit KI – AI generiert professionellen Dokumenttext, Screenshots als Anhang."""
        if not self.events:
            messagebox.showwarning("Keine Daten", "Noch keine Aktionen aufgezeichnet.")
            return
        template_id = self._get_selected_template_id()
        _, _, fmt = TEMPLATE_FILES[template_id]
        try:
            template_path = get_template_path(template_id)
        except FileNotFoundError as e:
            messagebox.showerror("Vorlage fehlt", str(e))
            return

        self._set_status("🤖 KI generiert Dokumenteninhalt ...", "#5c2d91")
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)

        def _ai_thread():
            try:
                import config as cfg
                from ai_providers.base import get_provider
                from bridge.recording_to_doc import (
                    build_ai_description, events_to_doc_data, inject_screenshots_into_markdown
                )
                provider = get_provider(cfg.get_ai_config())
                description = build_ai_description(self.events, self.doc_title.get())
                md = provider.generate_document(
                    description=description,
                    title=self.doc_title.get(),
                    fmt=fmt,
                    template_id=template_id,
                    chapters=[],
                    aushang=False,
                    refs=[],
                )
                md_with_screenshots = inject_screenshots_into_markdown(md, self.events)
                data = events_to_doc_data(
                    self.events, self.doc_title.get(), template_id, fmt,
                    markdown_content=md_with_screenshots
                )
                self.root.after(0, lambda: self._run_export(data, template_path, fmt))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: messagebox.showerror("KI-Fehler", err))
                self.root.after(0, lambda: self._set_status("KI-Export fehlgeschlagen.", "#a4262c"))
                self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
                self.root.after(0, self._check_ai_available)

        threading.Thread(target=_ai_thread, daemon=True).start()

    def _run_export(self, data: dict, template_path: str, fmt: str):
        ext_map = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}
        ext = ext_map.get(fmt, ".docx")
        default_name = self.doc_title.get().replace(" ", "_").replace("/", "-") + ext

        path = filedialog.asksaveasfilename(
            title="Dokument speichern",
            defaultextension=ext,
            initialfile=default_name,
            filetypes=[
                (("Word-Dokument", "*.docx") if fmt == "word" else
                 ("Excel-Datei", "*.xlsx") if fmt == "excel" else
                 ("PowerPoint-Präsentation", "*.pptx")),
                ("Alle Dateien", "*.*"),
            ]
        )
        if not path:
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()
            return

        try:
            if fmt == "word":
                from generator.word_generator import generate_word
                buf, _ = generate_word(data, template_path)
            elif fmt == "excel":
                from generator.excel_generator import generate_excel
                buf, _ = generate_excel(data, template_path)
            else:
                from generator.pptx_generator import generate_pptx
                buf, _ = generate_pptx(data, template_path)

            buf.seek(0)
            with open(path, "wb") as f:
                f.write(buf.read())

            self._set_status(f"✅ Exportiert: {Path(path).name}", "#107c10")
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()

            if messagebox.askyesno(
                "Export erfolgreich",
                f"Dokument gespeichert:\n{path}\n\nJetzt öffnen?"
            ):
                import webbrowser
                webbrowser.open(f"file:///{path}")

        except Exception as e:
            messagebox.showerror("Export fehlgeschlagen", str(e))
            self._set_status("Export fehlgeschlagen.", "#a4262c")
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()

    # ------------------------------------------------------------------
    # Hilfsfunktionen
    # ------------------------------------------------------------------

    def _set_status(self, text: str, color: str = "#323130"):
        self.status_var.set(text)
        self.status_label.config(fg=color)

    def _update_counter(self):
        active = [e for e in self.events if e.action_type not in ("start", "stop")]
        total      = len(active)
        screenshots = sum(1 for e in active if e.screenshot_b64)
        sections    = sum(1 for e in active if e.action_type == "note")
        self.counter_var.set(
            f"Ereignisse: {total}  |  Screenshots: {screenshots}  |  Abschnitte: {sections}"
        )

    def _on_close(self):
        if self.recording:
            if not messagebox.askyesno("Beenden",
                                       "Aufnahme läuft noch. Trotzdem beenden?\n"
                                       "(Nicht exportierte Daten gehen verloren.)"):
                return
        if hasattr(self, "_hk_listener"):
            self._hk_listener.stop()
        if self.recording:
            self.tracker.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = ITDocuMakerApp()
    app.run()
