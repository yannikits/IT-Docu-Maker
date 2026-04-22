"""
Snipping-Tool-artiger Bereichs-Screenshot.
Zeigt ein Vollbild-Overlay an; der Benutzer wählt per Klick+Ziehen einen Bereich aus.
Unterwanderung von HiDPI-Skalierung (Retina/Windows-Display-Scaling) wird behandelt.
"""

import base64
import tkinter as tk
from io import BytesIO
from typing import Optional

import mss
from PIL import Image, ImageTk


class SnippingTool:
    """Snipping-Tool: Bereich auf dem Bildschirm auswählen und als JPEG-Base64 zurückgeben."""

    def capture_area(self, parent: tk.Tk) -> Optional[str]:
        """
        Öffnet ein Vollbild-Overlay auf dem Primärmonitor.
        Rückgabe: Base64-JPEG des ausgewählten Bereichs  oder  None bei Abbruch.
        """
        # ── 1. Primärmonitor aufnehmen ────────────────────────────────────────
        with mss.mss() as sct:
            mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
            raw = sct.grab(mon)
            mon_left = mon.get("left", 0)
            mon_top  = mon.get("top",  0)

        img_full = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")

        # ── 2. Logische Bildschirmgröße (HiDPI-Skalierung) ───────────────────
        screen_w = parent.winfo_screenwidth()
        screen_h = parent.winfo_screenheight()

        scale_x = img_full.width  / screen_w
        scale_y = img_full.height / screen_h

        # Darstellungsbild in logischer Auflösung
        if abs(scale_x - 1.0) > 0.05 or abs(scale_y - 1.0) > 0.05:
            img_display = img_full.resize((screen_w, screen_h), Image.LANCZOS)
        else:
            img_display = img_full
            scale_x = scale_y = 1.0

        # Leicht abgedunkeltes Hintergrundbild
        img_dark = img_display.point(lambda p: int(p * 0.58))

        result = [None]

        # ── 3. Overlay-Fenster ────────────────────────────────────────────────
        overlay = tk.Toplevel(parent)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.geometry(f"{screen_w}x{screen_h}+{mon_left}+{mon_top}")

        canvas = tk.Canvas(
            overlay,
            width=screen_w,
            height=screen_h,
            highlightthickness=0,
            cursor="crosshair",
            bg="black",
        )
        canvas.pack(fill="both", expand=True)

        # Abgedunkeltes Bild als Hintergrund
        tk_dark = ImageTk.PhotoImage(img_dark)
        canvas._tk_dark = tk_dark          # GC-Schutz
        canvas.create_image(0, 0, anchor="nw", image=tk_dark)

        # Hinweisbanner
        bw = 440
        canvas.create_rectangle(
            screen_w // 2 - bw // 2 - 6,  8,
            screen_w // 2 + bw // 2 + 6, 50,
            fill="#1a1a2e", outline="#0078d4", width=1,
        )
        canvas.create_text(
            screen_w // 2, 29,
            text="Bereich auswählen   –   Esc: Abbrechen",
            fill="white",
            font=("Segoe UI", 13, "bold"),
        )

        # ── 4. Maus-Interaktion ───────────────────────────────────────────────
        state: dict = {"x0": 0, "y0": 0, "rect_id": None, "bright_id": None}

        def _clear():
            if state["rect_id"]:
                canvas.delete(state["rect_id"])
                state["rect_id"] = None
            if state["bright_id"]:
                canvas.delete(state["bright_id"])
                state["bright_id"] = None

        def on_press(e):
            state["x0"], state["y0"] = e.x, e.y
            _clear()

        def on_drag(e):
            _clear()
            x0, y0, x1, y1 = state["x0"], state["y0"], e.x, e.y
            bx0, by0 = min(x0, x1), min(y0, y1)
            bx1, by1 = max(x0, x1), max(y0, y1)

            # Helles Bild im ausgewählten Bereich
            if bx1 > bx0 and by1 > by0:
                bright_crop = img_display.crop((bx0, by0, bx1, by1))
                ph = ImageTk.PhotoImage(bright_crop)
                canvas._crop_ph = ph          # GC-Schutz
                state["bright_id"] = canvas.create_image(
                    bx0, by0, anchor="nw", image=ph
                )

            # Rechteck-Rahmen
            state["rect_id"] = canvas.create_rectangle(
                x0, y0, e.x, e.y,
                outline="#0078d4", width=2,
            )

        def on_release(e):
            x0 = min(state["x0"], e.x)
            y0 = min(state["y0"], e.y)
            x1 = max(state["x0"], e.x)
            y1 = max(state["y0"], e.y)

            if x1 - x0 < 5 or y1 - y0 < 5:       # zu klein → abbrechen
                overlay.destroy()
                return

            # In echte Pixelkoordinaten zurückrechnen (HiDPI)
            rx0, ry0 = int(x0 * scale_x), int(y0 * scale_y)
            rx1, ry1 = int(x1 * scale_x), int(y1 * scale_y)

            cropped = img_full.crop((rx0, ry0, rx1, ry1))
            buf = BytesIO()
            cropped.save(buf, format="JPEG", quality=90, optimize=True)
            result[0] = base64.b64encode(buf.getvalue()).decode("utf-8")
            overlay.destroy()

        canvas.bind("<ButtonPress-1>",   on_press)
        canvas.bind("<B1-Motion>",       on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", lambda _e: overlay.destroy())

        overlay.focus_force()
        parent.wait_window(overlay)
        return result[0]
