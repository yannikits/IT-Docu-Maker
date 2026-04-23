#!/usr/bin/env python3
"""
IT-Docu-Maker
Workflow:
  1. Aufnahme starten (F8)
  2. Abschnitt/Kapitel benennen  →  »Neuer Abschnitt«
  3. Vorgehenstext ins Notizfeld schreiben
  4. F9: Bereich-Screenshot → speichert Text + Screenshot als Schritt
     F10: Text-Schritt ohne Screenshot speichern
  5. Aufnahme stoppen (F8)  →  Vorlage wählen  →  exportieren

Hotkeys:
  F8   Aufnahme starten / stoppen
  F9   Bereich-Screenshot auswählen (speichert aktuellen Text + Bild)
  F10  Textnotiz ohne Screenshot speichern
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from pathlib import Path
from io import BytesIO
import base64
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
    print(f"Installiere: {', '.join(missing)}")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + missing)

from PIL import Image, ImageTk
from recorder.event_tracker import ActionEvent, EventTracker
from recorder.screen_capture import ScreenCapture
from recorder.snipping import SnippingTool

# ---------------------------------------------------------------------------
# Vorlagen-Konfiguration
# ---------------------------------------------------------------------------
VORLAGEN_DIR = os.path.join(os.path.dirname(__file__), "vorlagen")

TEMPLATE_FILES = {
    "intern":        ("Internes_Dokument_Vorlage.docx",      "Internes Dokument (Word)",      "word"),
    "extern":        ("Externes_Dokument_Vorlage.docx",      "Externes Dokument (Word)",      "word"),
    "kunde":         ("Kundenanleitung_Vorlage.docx",        "Kundenanleitung (Word)",        "word"),
    "netzwerk":      ("Netzwerkdoku_Vorlage.xlsx",           "Netzwerkdokumentation (Excel)", "excel"),
    "intern_xl":     ("Internes_Dokument_Vorlage.xlsx",      "Internes Dokument (Excel)",     "excel"),
    "extern_xl":     ("Externes_Dokument_Vorlage.xlsx",      "Externes Dokument (Excel)",     "excel"),
    "praesentation": ("Präsentationsvorlage_Vorlage.pptx", "Präsentation (PPTX)",          "ppt"),
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
            f"Bitte kopiere die Vorlagen aus WBI-Docu-Assist\n"
            f"(wbi-doku/vorlagen/) in den Ordner:\n{VORLAGEN_DIR}"
        )
    return path


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------

class ITDocuMakerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IT-Docu-Maker")
        self.root.resizable(True, True)
        self.root.minsize(480, 600)

        self.events: list         = []
        self.recording: bool      = False
        self._snipping_active     = False   # Re-entranz-Schutz für Snipping-Tool

        self.capture   = ScreenCapture()
        self.tracker   = EventTracker(self._on_tracked_event)
        self.snipping  = SnippingTool()
        self.event_queue: queue.Queue = queue.Queue()

        # Konfigurationsvariablen
        self.doc_title          = tk.StringVar(
            value=f"IT-Dokumentation {datetime.now().strftime('%Y-%m-%d')}"
        )
        self.capture_on_click   = tk.BooleanVar(value=True)
        self.capture_on_scroll  = tk.BooleanVar(value=False)
        self.screenshot_quality = tk.IntVar(value=85)

        self._build_ui()
        self._setup_hotkeys()
        self.root.after(100, self._process_queue)

    # -----------------------------------------------------------------------
    # UI
    # -----------------------------------------------------------------------

    def _build_ui(self):
        self.root.configure(bg="#f3f2f1")
        PAD = {"padx": 12, "pady": 4}

        # ── Titelleiste ─────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg="#0078d4", pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="IT-Docu-Maker",
                 bg="#0078d4", fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT, padx=14)
        tk.Label(hdr, text="Aufzeichnung  →  Fertiges Dokument",
                 bg="#0078d4", fg="#c7e0f4",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)

        # ── Statuszeile ───────────────────────────────────────────────────────
        self.status_var = tk.StringVar(
            value="Bereit – Aufnahme starten, Schritte dokumentieren, dann exportieren."
        )
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            bg="#f3f2f1", font=("Segoe UI", 10), anchor="w"
        )
        self.status_label.pack(fill=tk.X, padx=12, pady=(6, 2))

        # ── Aufnahme-Steuerung ───────────────────────────────────────────────
        self.rec_btn = tk.Button(
            self.root,
            text="▶  Aufnahme starten (F8)",
            command=self.toggle_recording,
            bg="#107c10", fg="white",
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT, padx=10, pady=8, cursor="hand2",
        )
        self.rec_btn.pack(fill=tk.X, padx=12, pady=(4, 2))

        self.counter_var = tk.StringVar(value="Schritte: 0  |  Abschnitte: 0  |  Screenshots: 0")
        tk.Label(
            self.root, textvariable=self.counter_var,
            bg="#f3f2f1", fg="#605e5c", font=("Segoe UI", 9)
        ).pack(anchor="w", padx=12)

        # ── Schritt erfassen ──────────────────────────────────────────────────
        step_lf = ttk.LabelFrame(self.root, text="Schritt erfassen", padding=(10, 6))
        step_lf.pack(fill=tk.X, padx=12, pady=6)

        # – Abschnitt/Kapitel –
        sec_row = tk.Frame(step_lf)
        sec_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(sec_row, text="Abschnitt / Kapitel:",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.section_var = tk.StringVar()
        self.section_entry = tk.Entry(
            sec_row, textvariable=self.section_var,
            font=("Segoe UI", 9), width=22,
            state=tk.DISABLED,
        )
        self.section_entry.pack(side=tk.LEFT, padx=(6, 4), fill=tk.X, expand=True)
        # Enter im Abschnitt-Feld löst "Neuer Abschnitt" aus
        self.section_entry.bind("<Return>", lambda _e: self.add_section())

        self.new_section_btn = tk.Button(
            sec_row, text="+ Neuer Abschnitt",
            command=self.add_section,
            state=tk.DISABLED,
            bg="#ca5010", fg="white",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT, padx=8, pady=3, cursor="hand2",
        )
        self.new_section_btn.pack(side=tk.LEFT)

        # – Notiz-Textfeld –
        tk.Label(step_lf, text="Notiz / Beschreibung des Schritts:",
                 font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, pady=(0, 3))

        text_outer = tk.Frame(step_lf, bg="#c8c6c4", bd=1, relief=tk.SOLID)
        text_outer.pack(fill=tk.X)
        self.note_text = tk.Text(
            text_outer,
            font=("Segoe UI", 10),
            height=4,
            relief=tk.FLAT,
            bg="#fafafa",
            fg="#323130",
            insertbackground="#0078d4",
            wrap=tk.WORD,
            padx=7, pady=5,
            state=tk.DISABLED,
        )
        note_scroll = tk.Scrollbar(text_outer, command=self.note_text.yview)
        self.note_text.configure(yscrollcommand=note_scroll.set)
        self.note_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        note_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Ctrl+Enter speichert Schritt (ohne Screenshot)
        self.note_text.bind("<Control-Return>", lambda _e: self.save_step_text_only())

        # – Aktions-Buttons –
        btn_row = tk.Frame(step_lf)
        btn_row.pack(fill=tk.X, pady=(6, 0))

        self.screenshot_btn = tk.Button(
            btn_row,
            text="\U0001f4f7  Bereich-Screenshot (F9)",
            command=self.take_area_screenshot,
            state=tk.DISABLED,
            bg="#004578", fg="white",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
        )
        self.screenshot_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.save_btn = tk.Button(
            btn_row,
            text="✎  Text speichern (F10)",
            command=self.save_step_text_only,
            state=tk.DISABLED,
            bg="#8764b8", fg="white",
            font=("Segoe UI", 9),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
        )
        self.save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # – Screenshot-Vorschau (feste Größe, immer sichtbar) –
        self.preview_canvas = tk.Canvas(
            step_lf,
            width=456, height=150,
            bg="#edebe9",
            highlightthickness=1,
            highlightbackground="#c8c6c4",
            cursor="hand2",
        )
        self.preview_canvas.pack(pady=(8, 0), fill=tk.X)
        self._preview_placeholder()
        self.preview_canvas.bind("<Button-1>", self._on_preview_click)
        self._preview_b64: str = ""   # aktuell angezeigtes Bild

        # ── Einstellungen ───────────────────────────────────────────────────────
        cfg_lf = ttk.LabelFrame(self.root, text="Einstellungen", padding=(10, 4))
        cfg_lf.pack(fill=tk.X, padx=12, pady=4)

        row1 = tk.Frame(cfg_lf)
        row1.pack(fill=tk.X)
        tk.Label(row1, text="Titel:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.doc_title, font=("Segoe UI", 9),
                 width=40).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)

        row2 = tk.Frame(cfg_lf)
        row2.pack(fill=tk.X, pady=(4, 0))
        tk.Label(row2, text="Auto-SS bei:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        tk.Checkbutton(row2, text="Klick",   variable=self.capture_on_click,
                       font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(6, 0))
        tk.Checkbutton(row2, text="Scrollen", variable=self.capture_on_scroll,
                       font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(row2, text="  Qualität:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        ttk.Scale(row2, from_=50, to=95, variable=self.screenshot_quality,
                  orient=tk.HORIZONTAL, length=90).pack(side=tk.LEFT, padx=(4, 0))
        self.qlbl = tk.Label(row2, text=f"{self.screenshot_quality.get()}%",
                             font=("Segoe UI", 9), width=4)
        self.qlbl.pack(side=tk.LEFT)
        self.screenshot_quality.trace_add(
            "write",
            lambda *_: self.qlbl.config(text=f"{self.screenshot_quality.get()}%")
        )

        # ── Export ───────────────────────────────────────────────────────────
        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=12, pady=6)

        exp_lf = ttk.LabelFrame(self.root, text="Dokument exportieren", padding=(10, 6))
        exp_lf.pack(fill=tk.X, padx=12, pady=4)

        vrow = tk.Frame(exp_lf)
        vrow.pack(fill=tk.X, pady=(0, 6))
        tk.Label(vrow, text="Vorlage:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.template_combo = ttk.Combobox(
            vrow,
            values=[v[1] for v in TEMPLATE_FILES.values()],
            state="readonly", width=38, font=("Segoe UI", 9),
        )
        self.template_combo.current(0)
        self.template_combo.pack(side=tk.LEFT, padx=(6, 0))

        exp_btns = tk.Frame(exp_lf)
        exp_btns.pack(fill=tk.X)
        self.export_btn = tk.Button(
            exp_btns, text="\U0001f4be  Exportieren (ohne KI)",
            command=self.export_document,
            state=tk.DISABLED,
            bg="#0078d4", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=12, pady=8, cursor="hand2",
        )
        self.export_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.ai_export_btn = tk.Button(
            exp_btns, text="\U0001f916  Exportieren (mit KI)",
            command=self.export_with_ai,
            state=tk.DISABLED,
            bg="#5c2d91", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=12, pady=8, cursor="hand2",
        )
        self.ai_export_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            self.root,
            text="F8: Aufnahme  │  F9: Bereich-Screenshot  │  F10: Text speichern  "
                 "│  Strg+Enter: Text speichern",
            bg="#f3f2f1", fg="#605e5c", font=("Segoe UI", 8),
        ).pack(pady=(2, 8))
       
      # Fortschrittsanzeige (nur während KI-Export sichtbar)
        self.progress_frame = tk.Frame(exp_lf)
        self.progress_lbl = tk.Label(
            self.progress_frame, text="", anchor="w",
            font=("Segoe UI", 9), fg="#5c2d91",
        )
        self.progress_lbl.pack(fill=tk.X, pady=(6, 1))
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, mode="determinate", maximum=100,
        )
        self.progress_bar.pack(fill=tk.X)
        self._progress_timer = None
        self._progress_value = 0.0

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -----------------------------------------------------------------------
    # Screenshot-Vorschau
    # -----------------------------------------------------------------------

    def _preview_placeholder(self):
        self.preview_canvas.delete("all")
        self.preview_canvas.create_text(
            self.preview_canvas.winfo_reqwidth() // 2 or 228,
            75,
            text="Kein Screenshot aufgenommen",
            fill="#a0a0a0",
            font=("Segoe UI", 9),
        )

    def _show_preview(self, b64_data: str):
        try:
            img = Image.open(BytesIO(base64.b64decode(b64_data)))
            cw = self.preview_canvas.winfo_width() or 456
            ch = self.preview_canvas.winfo_height() or 150
            img.thumbnail((cw - 4, ch - 4), Image.LANCZOS)
            ph = ImageTk.PhotoImage(img)
            self.preview_canvas._ph = ph
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(
                cw // 2, ch // 2, anchor="center", image=ph
            )
            self._preview_b64 = b64_data
        except Exception:
            pass

    def _on_preview_click(self, _event):
        if not self._preview_b64:
            return
        try:
            top = tk.Toplevel(self.root)
            top.title("Screenshot-Vorschau")
            top.attributes("-topmost", True)
            img = Image.open(BytesIO(base64.b64decode(self._preview_b64)))
            img.thumbnail((1000, 750), Image.LANCZOS)
            ph = ImageTk.PhotoImage(img)
            lbl = tk.Label(top, image=ph, bg="#1a1a2e")
            lbl._ph = ph
            lbl.pack(padx=4, pady=4)
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Hotkeys (pynput)
    # -----------------------------------------------------------------------

    def _setup_hotkeys(self):
        from pynput import keyboard as pkeyboard

        def on_press(key):
            try:
                k = key.name if hasattr(key, "name") else str(key)
            except AttributeError:
                k = str(key)
            k = k.lower().replace("key.", "")
            if k == "f8":   self.root.after(0, self.toggle_recording)
            elif k == "f9": self.root.after(0, self.take_area_screenshot)
            elif k == "f10": self.root.after(0, self.save_step_text_only)

        self._hk_listener = pkeyboard.Listener(on_press=on_press, daemon=True)
        self._hk_listener.start()

    # -----------------------------------------------------------------------
    # Aufnahme-Steuerung
    # -----------------------------------------------------------------------

    def toggle_recording(self):
        if self.recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self):
        self.recording = True
        self.events.clear()
        self._preview_placeholder()
        self._preview_b64 = ""
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="start",
            description="Aufnahme gestartet"
        ))
        self.tracker.start()

        self.rec_btn.config(text="■  Aufnahme stoppen (F8)", bg="#a4262c")
        self._set_input_state(tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)
        self._set_status("\U0001f534 Aufnahme läuft – Schritte dokumentieren ...", "#a4262c")
        self._update_counter()

    def _stop_recording(self):
        self.recording = False
        self.tracker.stop()
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="stop",
            description="Aufnahme beendet"
        ))

        self.rec_btn.config(text="▶  Aufnahme starten (F8)", bg="#107c10")
        self._set_input_state(tk.DISABLED)

        has_content = any(
            e.action_type in ("step", "click", "scroll")
            for e in self.events
        )
        if has_content:
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()

        steps    = sum(1 for e in self.events if e.action_type == "step")
        auto_ev  = sum(1 for e in self.events if e.action_type in ("click", "scroll"))
        self._set_status(
            f"✅ Gestoppt – {steps} Schritte, {auto_ev} Auto-Ereignisse. Vorlage wählen und exportieren.",
            "#107c10",
        )
        self._update_counter()

    def _set_input_state(self, state):
        """Aktiviert oder deaktiviert alle Eingabeelemente."""
        bg = "#fafafa" if state == tk.NORMAL else "#f3f2f1"
        self.note_text.config(state=state, bg=bg)
        self.section_entry.config(state=state)
        self.screenshot_btn.config(state=state)
        self.save_btn.config(state=state)
        self.new_section_btn.config(state=state)

    def _check_ai_available(self):
        try:
            import config as cfg
            if cfg.ai_enabled():
                self.ai_export_btn.config(state=tk.NORMAL)
        except Exception:
            pass
    def _start_progress(self):
        self._progress_value = 0.0
        self.progress_bar["value"] = 0
        self.progress_frame.pack(fill=tk.X, pady=(6, 0))
        self._tick_progress()

    def _tick_progress(self):
        if self._progress_value < 90:
            self._progress_value = min(self._progress_value + 0.8, 90)
            self.progress_bar["value"] = self._progress_value
            if self._progress_value < 10:   stage = "Analysiere Aufzeichnung"
            elif self._progress_value < 30: stage = "Sende Anfrage an KI-Server"
            elif self._progress_value < 85: stage = "KI generiert Dokument"
            else:                           stage = "Finalisiere Dokument"
            self.progress_lbl.config(text=f"{stage} … {int(self._progress_value)}%")
            self._progress_timer = self.root.after(400, self._tick_progress)

    def _stop_progress(self, success: bool = True):
        if self._progress_timer:
            self.root.after_cancel(self._progress_timer)
            self._progress_timer = None
        if success:
            self.progress_bar["value"] = 100
            self.progress_lbl.config(text="Fertig – 100%")
            self.root.after(2000, self.progress_frame.pack_forget)
        else:
            self.progress_frame.pack_forget()

    # -----------------------------------------------------------------------
    # Auto-Tracking (sekundär)
    # -----------------------------------------------------------------------

    def _on_tracked_event(self, action_type, description, x, y):
        self.event_queue.put((action_type, description, x, y))

    def _process_queue(self):
        while not self.event_queue.empty():
            try:
                at, desc, x, y = self.event_queue.get_nowait()
                self._handle_tracked_event(at, desc, x, y)
            except queue.Empty:
                break
        self.root.after(100, self._process_queue)

    def _handle_tracked_event(self, action_type, description, x, y):
        if not self.recording or action_type == "key":
            return
        take_ss = (
            (action_type == "click"  and self.capture_on_click.get()) or
            (action_type == "scroll" and self.capture_on_scroll.get())
        )
        ss_b64 = (
            self.capture.capture_thumbnail(quality=self.screenshot_quality.get())
            if take_ss else None
        )
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type=action_type,
            description=description, screenshot_b64=ss_b64, x=x, y=y,
        ))
        self._update_counter()

    # -----------------------------------------------------------------------
    # Manuelle Aktionen
    # -----------------------------------------------------------------------

    def add_section(self):
        if not self.recording:
            return
        title = self.section_var.get().strip()
        if not title:
            self._set_status("Bitte erst einen Abschnittsnamen eingeben.", "#a4262c")
            return
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="section",
            description=title,
        ))
        self.section_var.set("")
        self._set_status(f"\U0001f4cc Abschnitt gespeichert: {title[:60]}", "#ca5010")
        self._update_counter()

    def _get_note(self) -> str:
        try:
            return self.note_text.get("1.0", tk.END).strip()
        except Exception:
            return ""

    def _clear_note(self):
        self.note_text.config(state=tk.NORMAL)
        self.note_text.delete("1.0", tk.END)

    # – F9: Bereichs-Screenshot –

    def take_area_screenshot(self):
        if not self.recording or self._snipping_active:
            return
        text = self._get_note()
        # Hauptfenster verstecken damit es nicht im Screenshot erscheint
        self.root.withdraw()
        self.root.after(220, lambda: self._do_snip(text))

    def _do_snip(self, text: str):
        self._snipping_active = True
        try:
            b64 = self.snipping.capture_area(self.root)
        finally:
            self._snipping_active = False
            self.root.deiconify()
            self.root.lift()

        if b64 is None:
            self._set_status("Screenshot abgebrochen.", "#605e5c")
            return

        desc = text if text else "Screenshot"
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="step",
            description=desc, screenshot_b64=b64, note=desc,
        ))
        self._show_preview(b64)
        self._set_status(f"\U0001f4f7 Schritt gespeichert: {desc[:60]}", "#004578")
        self._update_counter()

    # – F10: Nur Text –

    def save_step_text_only(self):
        if not self.recording:
            return
        text = self._get_note()
        if not text:
            self._set_status("Kein Text zum Speichern eingegeben.", "#605e5c")
            return
        self.events.append(ActionEvent(
            timestamp=time.time(), action_type="step",
            description=text, note=text,
        ))
        self._set_status(f"✎ Schritt gespeichert: {text[:60]}", "#8764b8")
        self._update_counter()

    # -----------------------------------------------------------------------
    # Zähler & Status
    # -----------------------------------------------------------------------

    def _update_counter(self):
        steps    = sum(1 for e in self.events if e.action_type == "step")
        sections = sum(1 for e in self.events if e.action_type == "section")
        ss       = sum(1 for e in self.events if e.screenshot_b64)
        self.counter_var.set(
            f"Schritte: {steps}  |  Abschnitte: {sections}  |  Screenshots: {ss}"
        )

    def _set_status(self, text: str, color: str = "#323130"):
        self.status_var.set(text)
        self.status_label.config(fg=color)

    # -----------------------------------------------------------------------
    # Export
    # -----------------------------------------------------------------------

    def _get_template_id(self) -> str:
        return list(TEMPLATE_FILES.keys())[self.template_combo.current()]

    def export_document(self):
        template_id = self._get_template_id()
        _, _, fmt = TEMPLATE_FILES[template_id]
        try:
            tpath = get_template_path(template_id)
        except FileNotFoundError as e:
            messagebox.showerror("Vorlage fehlt", str(e))
            return
        from bridge.recording_to_doc import recording_to_doc_data_no_ai
        data = recording_to_doc_data_no_ai(
            self.events, self.doc_title.get(), template_id, fmt
        )
        self._run_export(data, tpath, fmt)

    def export_with_ai(self):
        template_id = self._get_template_id()
        _, _, fmt = TEMPLATE_FILES[template_id]
        try:
            tpath = get_template_path(template_id)
        except FileNotFoundError as e:
            messagebox.showerror("Vorlage fehlt", str(e))
            return

        self._set_status("\U0001f916 KI generiert Dokumenteninhalt ...", "#5c2d91")
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)
        self.root.after(0, self._start_progress)


        def _thread():
            try:
                import config as cfg
                from ai_providers.base import get_provider
                from bridge.recording_to_doc import (
                    build_ai_description, events_to_doc_data,
                    inject_screenshots_into_markdown,
                )
                provider = get_provider(cfg.get_ai_config())
                description = build_ai_description(
                    self.events, self.doc_title.get()
                )
                md = provider.generate_document(
                    description=description,
                    title=self.doc_title.get(),
                    fmt=fmt,
                    template_id=template_id,
                    chapters=[], aushang=False, refs=[],
                )
                md_ss = inject_screenshots_into_markdown(md, self.events)
                data  = events_to_doc_data(
                    self.events, self.doc_title.get(),
                    template_id, fmt, md_ss
                )
                
                self.root.after(0, lambda: self._stop_progress(True))
                self.root.after(0, lambda: self._run_export(data, tpath, fmt))
            except Exception as exc:
                msg = str(exc)
                self.root.after(0, lambda: messagebox.showerror("KI-Fehler", msg))
                self.root.after(0, lambda: self._set_status("KI-Export fehlgeschlagen.", "#a4262c"))
                self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
                self.root.after(0, self._check_ai_available)
                self.root.after(0, lambda: self._stop_progress(False))


        threading.Thread(target=_thread, daemon=True).start()

    def _run_export(self, data: dict, template_path: str, fmt: str):
        ext = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}.get(fmt, ".docx")
        default = self.doc_title.get().replace(" ", "_").replace("/", "-") + ext

        path = filedialog.asksaveasfilename(
            title="Dokument speichern",
            defaultextension=ext,
            initialfile=default,
            filetypes=[
                (("Word-Dokument", "*.docx") if fmt == "word" else
                 ("Excel-Datei",   "*.xlsx") if fmt == "excel" else
                 ("PowerPoint",    "*.pptx")),
                ("Alle Dateien", "*.*"),
            ],
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

    # -----------------------------------------------------------------------
    # Fenster schließen
    # -----------------------------------------------------------------------

    def _on_close(self):
        if self.recording:
            if not messagebox.askyesno(
                "Beenden",
                "Aufnahme läuft noch. Trotzdem beenden?\n"
                "(Nicht exportierte Daten gehen verloren.)"
            ):
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
