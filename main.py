#!/usr/bin/env python3
"""
IT-Docu-Maker
Workflow:
  1. Aufnahme starten (F8)
  2. Kapitel (# H1) und Unterkapitel (## H2) in Text einfügen
  3. Dokumentationstext im Textfeld schreiben
  4. F9: Bereich-Screenshot → wird in Ordner gespeichert + Verweis in Text eingefügt
     F10: Markdown speichern
  5. Aufnahme stoppen (F8)  →  Vorlage wählen  →  exportieren

Hotkeys:
  F8   Aufnahme starten / stoppen
  F9   Bereich-Screenshot auswählen und als Datei speichern
  F10  Markdown-Datei speichern
"""

import os
import re
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
    "praesentation": ("Präsentationsvorlage_Vorlage.pptx",  "Präsentation (PPTX)",           "ppt"),
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
        self.root.minsize(480, 640)

        self.events: list           = []
        self.recording: bool        = False
        self._snipping_active       = False

        self._progress_timer        = None
        self._progress_value        = 0.0
        self._session_dir: str      = ""
        self._screenshot_count: int = 0

        self.capture  = ScreenCapture()
        self.tracker  = EventTracker(self._on_tracked_event)
        self.snipping = SnippingTool()
        self.event_queue: queue.Queue = queue.Queue()

        self.doc_title          = tk.StringVar(
            value=f"IT-Dokumentation {datetime.now().strftime('%Y-%m-%d')}"
        )
        self.capture_on_click   = tk.BooleanVar(value=False)
        self.capture_on_scroll  = tk.BooleanVar(value=False)
        self.screenshot_quality = tk.IntVar(value=85)
        self.subsection_var     = tk.StringVar()

        self._build_ui()
        self._setup_hotkeys()
        self.root.after(100, self._process_queue)

    # -----------------------------------------------------------------------
    # UI
    # -----------------------------------------------------------------------

    def _build_ui(self):
        self.root.configure(bg="#f3f2f1")

        hdr = tk.Frame(self.root, bg="#0078d4", pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="IT-Docu-Maker",
                 bg="#0078d4", fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT, padx=14)
        tk.Label(hdr, text="Aufzeichnung  →  Fertiges Dokument",
                 bg="#0078d4", fg="#c7e0f4",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)

        self.status_var = tk.StringVar(
            value="Bereit – Aufnahme starten, Inhalt erfassen, dann exportieren."
        )
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            bg="#f3f2f1", font=("Segoe UI", 10), anchor="w"
        )
        self.status_label.pack(fill=tk.X, padx=12, pady=(6, 2))

        self.rec_btn = tk.Button(
            self.root,
            text="▶  Aufnahme starten (F8)",
            command=self.toggle_recording,
            bg="#107c10", fg="white",
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT, padx=10, pady=8, cursor="hand2",
        )
        self.rec_btn.pack(fill=tk.X, padx=12, pady=(4, 2))

        self.counter_var = tk.StringVar(value="Kapitel: 0  |  Unterkapitel: 0  |  Screenshots: 0")
        tk.Label(
            self.root, textvariable=self.counter_var,
            bg="#f3f2f1", fg="#605e5c", font=("Segoe UI", 9)
        ).pack(anchor="w", padx=12)

        step_lf = ttk.LabelFrame(self.root, text="Dokumentstruktur & Inhalt", padding=(10, 6))
        step_lf.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # – Kapitel (H1) –
        sec_row = tk.Frame(step_lf)
        sec_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(sec_row, text="Kapitel (H1):",
                 font=("Segoe UI", 9), width=17, anchor="w").pack(side=tk.LEFT)
        self.section_var = tk.StringVar()
        self.section_entry = tk.Entry(
            sec_row, textvariable=self.section_var,
            font=("Segoe UI", 9),
            state=tk.DISABLED,
        )
        self.section_entry.pack(side=tk.LEFT, padx=(4, 4), fill=tk.X, expand=True)
        self.section_entry.bind("<Return>", lambda _e: self.add_section())
        self.new_section_btn = tk.Button(
            sec_row, text="+ Kapitel",
            command=self.add_section,
            state=tk.DISABLED,
            bg="#ca5010", fg="white",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT, padx=8, pady=3, cursor="hand2",
        )
        self.new_section_btn.pack(side=tk.LEFT)

        # – Unterkapitel (H2) –
        subsec_row = tk.Frame(step_lf)
        subsec_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(subsec_row, text="Unterkapitel (H2):",
                 font=("Segoe UI", 9), width=17, anchor="w").pack(side=tk.LEFT)
        self.subsection_entry = tk.Entry(
            subsec_row, textvariable=self.subsection_var,
            font=("Segoe UI", 9),
            state=tk.DISABLED,
        )
        self.subsection_entry.pack(side=tk.LEFT, padx=(4, 4), fill=tk.X, expand=True)
        self.subsection_entry.bind("<Return>", lambda _e: self.add_subsection())
        self.new_subsection_btn = tk.Button(
            subsec_row, text="+ Unterkapitel",
            command=self.add_subsection,
            state=tk.DISABLED,
            bg="#ca5010", fg="white",
            font=("Segoe UI", 9),
            relief=tk.FLAT, padx=8, pady=3, cursor="hand2",
        )
        self.new_subsection_btn.pack(side=tk.LEFT)

        # – Dokument-Textfeld –
        tk.Label(step_lf, text="Dokumentinhalt (Markdown):",
                 font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, pady=(0, 3))

        text_outer = tk.Frame(step_lf, bg="#c8c6c4", bd=1, relief=tk.SOLID)
        text_outer.pack(fill=tk.BOTH, expand=True)
        self.note_text = tk.Text(
            text_outer,
            font=("Segoe UI", 10),
            height=10,
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
        self.note_text.bind("<Control-Return>", lambda _e: self.save_step_text_only())

        # – Aktions-Buttons –
        btn_row = tk.Frame(step_lf)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        self.screenshot_btn = tk.Button(
            btn_row,
            text="\U0001f4f7  Screenshot (F9)",
            command=self.take_area_screenshot,
            state=tk.DISABLED,
            bg="#004578", fg="white",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
        )
        self.screenshot_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.save_btn = tk.Button(
            btn_row,
            text="💾  Speichern (F10)",
            command=self.save_step_text_only,
            state=tk.DISABLED,
            bg="#8764b8", fg="white",
            font=("Segoe UI", 9),
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
        )
        self.save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # – Screenshot-Vorschau –
        self.preview_canvas = tk.Canvas(
            step_lf,
            width=456, height=110,
            bg="#edebe9",
            highlightthickness=1,
            highlightbackground="#c8c6c4",
            cursor="hand2",
        )
        self.preview_canvas.pack(pady=(8, 0), fill=tk.X)
        self._preview_placeholder()
        self.preview_canvas.bind("<Button-1>", self._on_preview_click)
        self._preview_b64: str = ""

        # – Einstellungen –
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
        tk.Checkbutton(row2, text="Klick",    variable=self.capture_on_click,
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

        self.progress_bar = ttk.Progressbar(exp_lf, mode="determinate", maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(6, 0))
        self.progress_bar["value"] = 0

        tk.Label(
            self.root,
            text="F8: Aufnahme  │  F9: Screenshot  │  F10: Speichern  │  Strg+Enter: Speichern",
            bg="#f3f2f1", fg="#605e5c", font=("Segoe UI", 8),
        ).pack(pady=(2, 8))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -----------------------------------------------------------------------
    # Screenshot-Vorschau
    # -----------------------------------------------------------------------

    def _preview_placeholder(self):
        self.preview_canvas.delete("all")
        self.preview_canvas.create_text(
            self.preview_canvas.winfo_reqwidth() // 2 or 228,
            55,
            text="Kein Screenshot aufgenommen",
            fill="#a0a0a0",
            font=("Segoe UI", 9),
        )

    def _show_preview(self, b64_data: str):
        try:
            img = Image.open(BytesIO(base64.b64decode(b64_data)))
            cw = self.preview_canvas.winfo_width() or 456
            ch = self.preview_canvas.winfo_height() or 110
            img.thumbnail((cw - 4, ch - 4), Image.LANCZOS)
            ph = ImageTk.PhotoImage(img)
            self.preview_canvas._ph = ph
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(cw // 2, ch // 2, anchor="center", image=ph)
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
    # Hotkeys
    # -----------------------------------------------------------------------

    def _setup_hotkeys(self):
        from pynput import keyboard as pkeyboard

        def on_press(key):
            try:
                k = key.name if hasattr(key, "name") else str(key)
            except AttributeError:
                k = str(key)
            k = k.lower().replace("key.", "")
            if k == "f8":    self.root.after(0, self.toggle_recording)
            elif k == "f9":  self.root.after(0, self.take_area_screenshot)
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

    def _safe_title(self) -> str:
        return re.sub(r'[^\w\s\-äöüÄÖÜß]', '_', self.doc_title.get()).strip()[:50]

    def _start_recording(self):
        self.recording = True
        self.events.clear()
        self._preview_placeholder()
        self._preview_b64 = ""
        self._screenshot_count = 0

        safe = self._safe_title()
        date = datetime.now().strftime("%Y-%m-%d_%H%M")
        sessions_base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sessions")
        self._session_dir = os.path.join(sessions_base, f"{safe}_{date}")
        os.makedirs(self._session_dir, exist_ok=True)

        self.note_text.config(state=tk.NORMAL)
        self.note_text.delete("1.0", tk.END)

        self.tracker.start()
        self.rec_btn.config(text="■  Aufnahme stoppen (F8)", bg="#a4262c")
        self._set_input_state(tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)
        self._set_status(
            f"\U0001f534 Aufnahme läuft – Ordner: {Path(self._session_dir).name}",
            "#a4262c"
        )
        self._update_counter()

    def _stop_recording(self):
        self.recording = False
        self.tracker.stop()

        has_content = bool(self.note_text.get("1.0", tk.END).strip())

        self.rec_btn.config(text="▶  Aufnahme starten (F8)", bg="#107c10")
        self._set_input_state(tk.DISABLED)

        if has_content:
            self.export_btn.config(state=tk.NORMAL)
            self._check_ai_available()

        self._autosave_md()
        self._set_status(
            f"✅ Gestoppt – gespeichert in: {Path(self._session_dir).name}",
            "#107c10",
        )
        self._update_counter()

    def _set_input_state(self, state):
        bg = "#fafafa" if state == tk.NORMAL else "#f3f2f1"
        self.note_text.config(state=state, bg=bg)
        self.section_entry.config(state=state)
        self.subsection_entry.config(state=state)
        self.screenshot_btn.config(state=state)
        self.save_btn.config(state=state)
        self.new_section_btn.config(state=state)
        self.new_subsection_btn.config(state=state)

    def _check_ai_available(self):
        try:
            import config as cfg
            if cfg.ai_enabled():
                self.ai_export_btn.config(state=tk.NORMAL)
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Auto-Tracking
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
        if take_ss and self._session_dir:
            ss_b64 = self.capture.capture_thumbnail(quality=self.screenshot_quality.get())
            if ss_b64:
                self._screenshot_count += 1
                n = self._screenshot_count
                ss_dir = os.path.join(self._session_dir, "screenshots")
                os.makedirs(ss_dir, exist_ok=True)
                img_path = os.path.join(ss_dir, f"screenshot_{n:03d}.png")
                try:
                    pil_img = Image.open(BytesIO(base64.b64decode(ss_b64)))
                    pil_img.save(img_path, format="PNG")
                except Exception:
                    pass

    # -----------------------------------------------------------------------
    # Manuelle Aktionen
    # -----------------------------------------------------------------------

    def _insert_at_cursor(self, text: str):
        try:
            self.note_text.insert(tk.INSERT, text)
        except Exception:
            self.note_text.insert(tk.END, text)

    def add_section(self):
        if not self.recording:
            return
        title = self.section_var.get().strip()
        if not title:
            self._set_status("Bitte einen Kapitelnamen eingeben.", "#a4262c")
            return
        self._insert_at_cursor(f"\n# {title}\n\n")
        self.section_var.set("")
        self._autosave_md()
        self._set_status(f"\U0001f4cc Kapitel eingefügt: {title[:60]}", "#ca5010")
        self._update_counter()

    def add_subsection(self):
        if not self.recording:
            return
        title = self.subsection_var.get().strip()
        if not title:
            self._set_status("Bitte einen Unterkapitelnamen eingeben.", "#a4262c")
            return
        self._insert_at_cursor(f"\n## {title}\n\n")
        self.subsection_var.set("")
        self._autosave_md()
        self._set_status(f"\U0001f4cc Unterkapitel eingefügt: {title[:60]}", "#ca5010")
        self._update_counter()

    def _get_note(self) -> str:
        try:
            return self.note_text.get("1.0", tk.END).strip()
        except Exception:
            return ""

    def take_area_screenshot(self):
        if not self.recording or self._snipping_active:
            return
        orig_geo = self.root.geometry()
        self.root.geometry("+99999+0")
        self.root.after(220, lambda: self._do_snip(orig_geo))

    def _do_snip(self, orig_geo: str = ""):
        self._snipping_active = True
        try:
            b64 = self.snipping.capture_area(self.root)
        finally:
            self._snipping_active = False
            if orig_geo:
                self.root.geometry(orig_geo)
            self.root.lift()

        if b64 is None:
            self._set_status("Screenshot abgebrochen.", "#605e5c")
            return

        from recorder.annotation_editor import AnnotationEditor
        editor = AnnotationEditor(self.root, b64)
        if editor.result_b64 is None:
            self._set_status("Screenshot verworfen.", "#605e5c")
            return
        b64 = editor.result_b64

        self._screenshot_count += 1
        n = self._screenshot_count
        ss_dir = os.path.join(self._session_dir, "screenshots")
        os.makedirs(ss_dir, exist_ok=True)

        img_filename = f"screenshot_{n:03d}.png"
        img_path = os.path.join(ss_dir, img_filename)

        pil_img = Image.open(BytesIO(base64.b64decode(b64)))
        pil_img.save(img_path, format="PNG")

        try:
            cursor_line = int(self.note_text.index(tk.INSERT).split('.')[0])
            start = f"{max(1, cursor_line - 3)}.0"
            context = self.note_text.get(start, tk.INSERT).strip()
            desc = context[-200:] if context else f"Screenshot {n}"
        except Exception:
            desc = f"Screenshot {n}"

        txt_path = os.path.join(ss_dir, f"screenshot_{n:03d}.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(desc)

        self._insert_at_cursor(f"\n![Screenshot {n}](screenshots/{img_filename})\n\n")
        self._show_preview(b64)
        self.note_text.config(bg="#e8f5e9")
        self.root.after(800, lambda: self.note_text.config(bg="#fafafa"))
        self._autosave_md()
        self._set_status(f"\U0001f4f7 Screenshot {n} gespeichert: {img_filename}", "#004578")
        self._update_counter()

    def save_step_text_only(self):
        if not self.recording:
            return
        self._autosave_md()
        self._set_status("💾 Markdown gespeichert", "#8764b8")

    # -----------------------------------------------------------------------
    # Session / Markdown
    # -----------------------------------------------------------------------

    def _autosave_md(self):
        if not self._session_dir:
            return
        content = self.note_text.get("1.0", tk.END)
        md_path = os.path.join(self._session_dir, f"{self._safe_title()}.md")
        try:
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception:
            pass

    def _get_markdown_content(self) -> str:
        content = self.note_text.get("1.0", tk.END).strip()
        if not self._session_dir:
            return content

        def replace_file_ref(m):
            alt  = m.group(1)
            path = m.group(2)
            if path.startswith('data:'):
                return m.group(0)
            full_path = os.path.join(self._session_dir, path)
            try:
                with open(full_path, 'rb') as f:
                    img_bytes = f.read()
                mime = "image/jpeg" if img_bytes[:2] == b'\xff\xd8' else "image/png"
                b64_str = base64.b64encode(img_bytes).decode('ascii')
                return f"![{alt}](data:{mime};base64,{b64_str})"
            except Exception:
                return m.group(0)

        return re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', replace_file_ref, content)

    # -----------------------------------------------------------------------
    # Zähler & Status
    # -----------------------------------------------------------------------

    def _update_counter(self):
        try:
            content = self.note_text.get("1.0", tk.END)
        except Exception:
            content = ""
        h1_count = len(re.findall(r'^# .+', content, re.MULTILINE))
        h2_count = len(re.findall(r'^## .+', content, re.MULTILINE))
        self.counter_var.set(
            f"Kapitel: {h1_count}  |  Unterkapitel: {h2_count}  |  Screenshots: {self._screenshot_count}"
        )

    def _set_status(self, text: str, color: str = "#323130"):
        self.status_var.set(text)
        self.status_label.config(fg=color)

    # -----------------------------------------------------------------------
    # Fortschrittsbalken
    # -----------------------------------------------------------------------

    def _start_progress(self):
        self._progress_value = 0.0
        self.progress_bar["value"] = 0
        self._tick_progress()

    def _tick_progress(self):
        if self._progress_value < 90:
            self._progress_value += 0.8
            self.progress_bar["value"] = self._progress_value
        self._progress_timer = self.root.after(200, self._tick_progress)

    def _stop_progress(self):
        if self._progress_timer is not None:
            self.root.after_cancel(self._progress_timer)
            self._progress_timer = None
        self.progress_bar["value"] = 100
        self.root.after(800, lambda: self.progress_bar.config(value=0))

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
        from bridge.recording_to_doc import events_to_doc_data
        markdown = self._get_markdown_content()
        data = events_to_doc_data([], self.doc_title.get(), template_id, fmt, markdown)
        self._run_export(data, tpath, fmt)

    def export_with_ai(self):
        template_id = self._get_template_id()
        _, _, fmt = TEMPLATE_FILES[template_id]
        try:
            tpath = get_template_path(template_id)
        except FileNotFoundError as e:
            messagebox.showerror("Vorlage fehlt", str(e))
            return

        self._set_status("\U0001f916 KI verbessert Dokumenteninhalt ...", "#5c2d91")
        self.export_btn.config(state=tk.DISABLED)
        self.ai_export_btn.config(state=tk.DISABLED)
        self._start_progress()

        markdown = self._get_markdown_content()

        # Screenshots vor KI-Anfrage extrahieren und durch Platzhalter ersetzen,
        # damit die KI keine riesigen Base64-Blöcke erhält und verliert.
        extracted: dict[str, str] = {}
        _counter = [0]

        def _extract(m):
            _counter[0] += 1
            key = f"[SCREENSHOT_{_counter[0]}]"
            extracted[key] = m.group(0)
            return key

        md_for_ai = re.sub(r'!\[[^\]]*\]\(data:image/[^)]+\)', _extract, markdown)

        def _thread():
            try:
                import config as cfg
                from ai_providers.base import get_provider
                from bridge.recording_to_doc import events_to_doc_data

                provider = get_provider(cfg.get_ai_config())
                description = (
                    "Bitte verbessere und formatiere diese IT-Dokumentation professionell. "
                    "Behalte alle Kapitelstrukturen (# und ##), Inhalte und Screenshot-Platzhalter "
                    "([SCREENSHOT_N]) exakt an ihrer Position bei. Antworte NUR mit dem verbesserten "
                    "Markdown-Dokument, ohne Einleitung oder Erklärung.\n\n"
                    + md_for_ai
                )
                md_enhanced = provider.generate_document(
                    description=description,
                    title=self.doc_title.get(),
                    fmt=fmt,
                    template_id=template_id,
                    chapters=[], aushang=False, refs=[],
                )
                # Originale Screenshots wieder einfügen
                for key, img_tag in extracted.items():
                    md_enhanced = md_enhanced.replace(key, img_tag)
                data = events_to_doc_data([], self.doc_title.get(), template_id, fmt, md_enhanced)
                self.root.after(0, self._stop_progress)
                self.root.after(0, lambda: self._run_export(data, tpath, fmt))
            except Exception as exc:
                msg = str(exc)
                self.root.after(0, self._stop_progress)
                self.root.after(0, lambda: messagebox.showerror("KI-Fehler", msg))
                self.root.after(0, lambda: self._set_status("KI-Export fehlgeschlagen.", "#a4262c"))
                self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
                self.root.after(0, self._check_ai_available)

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
                "(Nicht gespeicherte Daten gehen verloren.)"
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