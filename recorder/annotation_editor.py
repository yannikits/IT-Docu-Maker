"""
Screenshot-Annotationseditor
Werkzeuge: Rahmen (▭), Pfeil (→), Text (T), Unschärfe (⊘)
"""

import math
import os
import base64
import tkinter as tk
from tkinter import ttk, colorchooser
from io import BytesIO
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageTk


class AnnotationEditor:
    def __init__(self, parent, b64_image: str):
        self.result_b64: str | None = None

        self._tool        = "rect"
        self._color       = "#ff0000"
        self._annotations = []      # all operations in order (incl. blur)
        self._drawing     = False
        self._start       = (0, 0)
        self._cur_item    = None
        self._text_entry  = None
        self._text_pos    = (0, 0)

        # Load image
        raw = base64.b64decode(b64_image)
        self.orig_img = Image.open(BytesIO(raw)).convert("RGB")
        self.img_w, self.img_h = self.orig_img.size

        # Display scale (fit into 85 x 80 % of screen)
        sw = parent.winfo_screenwidth()
        sh = parent.winfo_screenheight()
        self._s  = min((sw * 0.85) / self.img_w, (sh * 0.80) / self.img_h, 1.0)
        self._dw = int(self.img_w * self._s)
        self._dh = int(self.img_h * self._s)

        # Window
        self.win = tk.Toplevel(parent)
        self.win.title("Screenshot bearbeiten")
        self.win.resizable(False, False)
        self.win.grab_set()
        self.win.protocol("WM_DELETE_WINDOW", self._cancel)

        self._build_toolbar()

        frm = tk.Frame(self.win)
        frm.pack()
        self.cv = tk.Canvas(frm, width=self._dw, height=self._dh,
                             cursor="crosshair", highlightthickness=0)
        self.cv.pack()

        self._refresh_base()

        self.cv.bind("<ButtonPress-1>",  self._on_press)
        self.cv.bind("<B1-Motion>",       self._on_drag)
        self.cv.bind("<ButtonRelease-1>", self._on_release)

        self.win.update_idletasks()
        x = (sw - self.win.winfo_width())  // 2
        y = (sh - self.win.winfo_height()) // 2
        self.win.geometry(f"+{x}+{y}")

        parent.wait_window(self.win)

    # ── Toolbar ────────────────────────────────────────────────────────────────────

    def _build_toolbar(self):
        tb = tk.Frame(self.win, bg="#2d2d30", pady=5)
        tb.pack(fill=tk.X)

        self._tbtn: dict[str, tk.Button] = {}
        for tid, lbl in [("rect",  "▭  Rahmen"),
                         ("arrow", "→  Pfeil"),
                         ("text",  "T  Text"),
                         ("blur",  "⊘  Blur")]:
            btn = tk.Button(tb, text=lbl, bg="#3c3c3c", fg="white",
                            font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                            padx=10, pady=4, cursor="hand2",
                            command=lambda t=tid: self._select_tool(t))
            btn.pack(side=tk.LEFT, padx=(4, 0))
            self._tbtn[tid] = btn

        tk.Label(tb, text="  Farbe:", bg="#2d2d30", fg="#cccccc",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._cbtn = tk.Button(tb, bg=self._color, width=3,
                               relief=tk.RAISED, cursor="hand2",
                               command=self._pick_color)
        self._cbtn.pack(side=tk.LEFT, padx=(4, 0))

        tk.Label(tb, text="  Stärke:", bg="#2d2d30", fg="#cccccc",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._wvar = tk.IntVar(value=3)
        ttk.Combobox(tb, textvariable=self._wvar,
                     values=[1, 2, 3, 5, 8, 12],
                     width=3, state="readonly").pack(side=tk.LEFT, padx=(4, 0))

        tk.Button(tb, text="↩  Rükgängig", bg="#3c3c3c", fg="white",
                  font=("Segoe UI", 9), relief=tk.FLAT, padx=8, pady=4,
                  cursor="hand2", command=self._undo
                  ).pack(side=tk.LEFT, padx=(12, 0))

        tk.Button(tb, text="✗  Abbrechen", bg="#a4262c", fg="white",
                  font=("Segoe UI", 9), relief=tk.FLAT, padx=10, pady=4,
                  cursor="hand2", command=self._cancel
                  ).pack(side=tk.RIGHT, padx=(0, 4))
        tk.Button(tb, text="✓  Übernehmen", bg="#107c10", fg="white",
                  font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=10, pady=4,
                  cursor="hand2", command=self._accept
                  ).pack(side=tk.RIGHT, padx=(0, 4))

        self._select_tool("rect")

    def _select_tool(self, tool: str):
        self._tool = tool
        for tid, btn in self._tbtn.items():
            btn.config(bg="#0078d4" if tid == tool else "#3c3c3c")

    def _pick_color(self):
        _, hex_color = colorchooser.askcolor(color=self._color,
                                             title="Farbe wählen")
        if hex_color:
            self._color = hex_color
            self._cbtn.config(bg=hex_color)

    # ── Canvas base image ────────────────────────────────────────────────────────────────

    def _refresh_base(self):
        """Rebuild canvas background: orig image + all blur regions applied."""
        img = self.orig_img.copy()
        for ann in self._annotations:
            if ann["type"] == "blur":
                x0, y0, x1, y1 = ann["region"]
                region = img.crop((x0, y0, x1, y1))
                region = region.filter(ImageFilter.GaussianBlur(radius=18))
                img.paste(region, (x0, y0))
        disp = img.resize((self._dw, self._dh), Image.LANCZOS)
        self._base_ph = ImageTk.PhotoImage(disp)
        self.cv.delete("base")
        self.cv.create_image(0, 0, anchor="nw", image=self._base_ph, tags="base")
        self.cv.tag_lower("base")

    # ── Coordinate helper ────────────────────────────────────────────────────────────────

    def _to_img(self, cx: int, cy: int) -> tuple[int, int]:
        return int(cx / self._s), int(cy / self._s)

    # ── Mouse events ──────────────────────────────────────────────────────────────────

    def _on_press(self, e):
        if self._tool == "text":
            self._start_text(e.x, e.y)
            return
        self._drawing  = True
        self._start    = (e.x, e.y)
        self._cur_item = None

    def _on_drag(self, e):
        if not self._drawing:
            return
        x0, y0 = self._start
        if self._cur_item:
            self.cv.delete(self._cur_item)
        c, w = self._color, self._wvar.get()
        if self._tool == "rect":
            self._cur_item = self.cv.create_rectangle(
                x0, y0, e.x, e.y, outline=c, width=w)
        elif self._tool == "arrow":
            self._cur_item = self.cv.create_line(
                x0, y0, e.x, e.y, fill=c, width=w,
                arrow=tk.LAST, arrowshape=(16, 20, 6))
        elif self._tool == "blur":
            self._cur_item = self.cv.create_rectangle(
                x0, y0, e.x, e.y,
                outline="#00cfff", width=2, dash=(5, 3))

    def _on_release(self, e):
        if not self._drawing:
            return
        self._drawing = False
        x0, y0 = self._start
        x1, y1 = e.x, e.y

        if abs(x1 - x0) < 4 and abs(y1 - y0) < 4:
            if self._cur_item:
                self.cv.delete(self._cur_item)
            return

        ix0, iy0 = self._to_img(x0, y0)
        ix1, iy1 = self._to_img(x1, y1)
        c, w     = self._color, self._wvar.get()

        if self._tool == "blur":
            self.cv.delete(self._cur_item)
            region = (min(ix0, ix1), min(iy0, iy1),
                      max(ix0, ix1), max(iy0, iy1))
            self._annotations.append({"type": "blur", "region": region})
            self._refresh_base()
        else:
            self._annotations.append({
                "type":        self._tool,
                "coords":      (ix0, iy0, ix1, iy1),
                "color":       c,
                "width":       w,
                "canvas_item": self._cur_item,
            })
        self._cur_item = None

    # ── Text tool ────────────────────────────────────────────────────────────────────

    def _start_text(self, cx: int, cy: int):
        self._commit_text()
        fs  = max(11, int(13 * self._s))
        ent = tk.Entry(self.cv, fg=self._color,
                       font=("Segoe UI", fs, "bold"),
                       bg="#fffde7", relief=tk.SOLID, bd=1,
                       insertbackground=self._color)
        self.cv.create_window(cx, cy, window=ent, anchor="nw", tags="te")
        ent.focus_set()
        self._text_entry = ent
        self._text_pos   = (cx, cy)
        ent.bind("<Return>", lambda _: self._commit_text())
        ent.bind("<Escape>", lambda _: self._cancel_text())

    def _commit_text(self):
        if not self._text_entry:
            return
        txt = self._text_entry.get().strip()
        cx, cy = self._text_pos
        self.cv.delete("te")
        self._text_entry.destroy()
        self._text_entry = None
        if not txt:
            return
        fs_cv = max(11, int(13 * self._s))
        item  = self.cv.create_text(cx, cy, text=txt, fill=self._color,
                                     font=("Segoe UI", fs_cv, "bold"),
                                     anchor="nw")
        ix, iy = self._to_img(cx, cy)
        self._annotations.append({
            "type":        "text",
            "text":        txt,
            "coords":      (ix, iy),
            "color":       self._color,
            "font_size":   max(14, int(14 / self._s)),
            "canvas_item": item,
        })

    def _cancel_text(self):
        if self._text_entry:
            self.cv.delete("te")
            self._text_entry.destroy()
            self._text_entry = None

    # ── Undo ──────────────────────────────────────────────────────────────────────────────

    def _undo(self):
        if not self._annotations:
            return
        ann = self._annotations.pop()
        if ann.get("canvas_item"):
            self.cv.delete(ann["canvas_item"])
        if ann["type"] == "blur":
            self._refresh_base()

    # ── Export ──────────────────────────────────────────────────────────────────────────

    def _accept(self):
        self._commit_text()
        img  = self.orig_img.copy()
        draw = ImageDraw.Draw(img)

        for ann in self._annotations:
            t = ann["type"]

            if t == "blur":
                x0, y0, x1, y1 = ann["region"]
                region = img.crop((x0, y0, x1, y1))
                region = region.filter(ImageFilter.GaussianBlur(radius=18))
                img.paste(region, (x0, y0))
                draw = ImageDraw.Draw(img)

            elif t == "rect":
                x0, y0, x1, y1 = ann["coords"]
                draw.rectangle(
                    [min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)],
                    outline=ann["color"], width=ann["width"])

            elif t == "arrow":
                x0, y0, x1, y1 = ann["coords"]
                c, w = ann["color"], ann["width"]
                draw.line([x0, y0, x1, y1], fill=c, width=w)
                ang = math.atan2(y1 - y0, x1 - x0)
                hl  = max(16, w * 5)
                for a in (ang + math.pi - 0.45, ang + math.pi + 0.45):
                    draw.line([x1, y1,
                               x1 + int(hl * math.cos(a)),
                               y1 + int(hl * math.sin(a))],
                              fill=c, width=w)

            elif t == "text":
                x, y = ann["coords"]
                font = _load_font(ann["font_size"])
                draw.text((x, y), ann["text"],
                          fill=ann["color"], font=font)

        buf = BytesIO()
        img.save(buf, format="JPEG", quality=92)
        self.result_b64 = base64.b64encode(buf.getvalue()).decode()
        self.win.destroy()

    def _cancel(self):
        self._cancel_text()
        self.result_b64 = None
        self.win.destroy()


# ── Font helper ───────────────────────────────────────────────────────────────────────────

def _load_font(size: int) -> ImageFont.ImageFont:
    for path in [
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()
