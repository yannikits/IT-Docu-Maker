"""
Snipping-Tool: Ein separates Overlay pro Monitor.
Vermeidet Windows-DWM-Probleme mit kombinierten Mega-Fenstern.
"""
import base64
import tkinter as tk
from io import BytesIO
from typing import Optional

import mss
from PIL import Image, ImageTk


class SnippingTool:
    def capture_area(self, parent: tk.Tk) -> Optional[str]:
        """
        Oeffnet auf jedem angeschlossenen Monitor ein eigenes Overlay.
        Rueckgabe: Base64-JPEG des ausgewaehlten Bereichs oder None bei Abbruch.
        """
        # 1. Alle Monitore einzeln aufnehmen
        monitor_data = []
        with mss.mss() as sct:
            monitors = sct.monitors[1:] if len(sct.monitors) > 1 else sct.monitors[:1]
            for mon in monitors:
                raw = sct.grab(mon)
                img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
                monitor_data.append({"mon": dict(mon), "img": img})

        result  = [None]
        overlays = []
        done     = [False]

        def close_all(b64=None):
            if done[0]:
                return
            done[0]  = True
            result[0] = b64
            for ov in overlays:
                try:
                    ov.destroy()
                except Exception:
                    pass

        for md in monitor_data:
            mon     = md["mon"]
            img_mon = md["img"]   # physische Pixel

            # HiDPI: physische Pixel vs. logische Pixel
            scale_x = img_mon.width  / mon["width"]
            scale_y = img_mon.height / mon["height"]
            if abs(scale_x - 1.0) <= 0.05 and abs(scale_y - 1.0) <= 0.05:
                img_display = img_mon
                scale_x = scale_y = 1.0
            else:
                img_display = img_mon.resize((mon["width"], mon["height"]), Image.LANCZOS)

            img_dark = img_display.point(lambda p: int(p * 0.58))

            overlay = tk.Toplevel(parent)
            overlay.overrideredirect(True)
            overlay.attributes("-topmost", True)
            overlay.geometry(f"{mon['width']}x{mon['height']}+{mon['left']}+{mon['top']}")
            overlays.append(overlay)

            canvas = tk.Canvas(overlay,
                               width=mon["width"], height=mon["height"],
                               highlightthickness=0, cursor="crosshair", bg="black")
            canvas.pack(fill="both", expand=True)

            tk_dark = ImageTk.PhotoImage(img_dark)
            canvas._tk_dark = tk_dark
            canvas.create_image(0, 0, anchor="nw", image=tk_dark)

            # Banner
            cw = mon["width"]
            bw = min(440, cw - 20)
            canvas.create_rectangle(
                cw // 2 - bw // 2 - 6,  8,
                cw // 2 + bw // 2 + 6, 50,
                fill="#1a1a2e", outline="#0078d4", width=1,
            )
            canvas.create_text(
                cw // 2, 29,
                text="Bereich auswaehlen   –   Esc: Abbrechen",
                fill="white", font=("Segoe UI", 13, "bold"),
            )

            # Mouse-Handler in Closure einschliessen
            def _attach(cv, _img_mon, _img_disp, _sx, _sy):
                state = {"x0": 0, "y0": 0, "rect": None, "bright": None}

                def _clear():
                    if state["rect"]:
                        cv.delete(state["rect"])
                        state["rect"] = None
                    if state["bright"]:
                        cv.delete(state["bright"])
                        state["bright"] = None

                def on_press(e):
                    state["x0"], state["y0"] = e.x, e.y
                    _clear()

                def on_drag(e):
                    _clear()
                    x0, y0, x1, y1 = state["x0"], state["y0"], e.x, e.y
                    bx0, by0 = min(x0, x1), min(y0, y1)
                    bx1, by1 = max(x0, x1), max(y0, y1)
                    if bx1 > bx0 and by1 > by0:
                        crop = _img_disp.crop((bx0, by0, bx1, by1))
                        ph   = ImageTk.PhotoImage(crop)
                        cv._crop_ph = ph
                        state["bright"] = cv.create_image(bx0, by0, anchor="nw", image=ph)
                    state["rect"] = cv.create_rectangle(
                        x0, y0, e.x, e.y, outline="#0078d4", width=2
                    )

                def on_release(e):
                    x0, y0 = min(state["x0"], e.x), min(state["y0"], e.y)
                    x1, y1 = max(state["x0"], e.x), max(state["y0"], e.y)
                    if x1 - x0 < 5 or y1 - y0 < 5:
                        close_all(None)
                        return
                    # Physische Pixel-Koordinaten im Monitor-Screenshot
                    rx0, ry0 = int(x0 * _sx), int(y0 * _sy)
                    rx1, ry1 = int(x1 * _sx), int(y1 * _sy)
                    cropped = _img_mon.crop((rx0, ry0, rx1, ry1))
                    buf = BytesIO()
                    cropped.save(buf, format="JPEG", quality=90, optimize=True)
                    close_all(base64.b64encode(buf.getvalue()).decode("utf-8"))

                cv.bind("<ButtonPress-1>",   on_press)
                cv.bind("<B1-Motion>",       on_drag)
                cv.bind("<ButtonRelease-1>", on_release)

            _attach(canvas, img_mon, img_display, scale_x, scale_y)
            overlay.bind("<Escape>", lambda _e: close_all(None))

        if not overlays:
            return None

        overlays[0].focus_force()
        parent.wait_window(overlays[0])
        return result[0]
