"""
Screenshot-Annotationseditor
Werkzeuge: Auswahl (↖), Rahmen (▭), Pfeil (→), Text (T), Unschärfe (⊘)
"""

import math
import os
import base64
import tkinter as tk
from tkinter import ttk, colorchooser
from io import BytesIO
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageTk


_HANDLE_R = 5  # Griffpunkt-Radius in Canvas-Pixeln


def _point_to_segment_dist(px, py, ax, ay, bx, by) -> float:
    """Kürzester Abstand von Punkt (px,py) zum Segment (ax,ay)-(bx,by)."""
    dx, dy = bx - ax, by - ay
    if dx == dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


class AnnotationEditor:
    def __init__(self, parent, b64_image: str):
        self.result_b64: str | None = None

        self._tool        = "rect"
        self._color       = "#ff0000"
        self._annotations = []
        self._drawing     = False
        self._start       = (0, 0)
        self._cur_item    = None
        self._text_entry  = None
        self._text_pos    = (0, 0)
        self._edit_entry  = None
        self._edit_ann    = None

        # Select-Tool-Zustand
        self._selected_idx = -1
        self._sel_handles  = []
        self._drag_mode    = ""
        self._drag_start   = (0, 0)
        self._drag_orig    = None

        raw = base64.b64decode(b64_image)
        self.orig_img = Image.open(BytesIO(raw)).convert("RGB")
        self.img_w, self.img_h = self.orig_img.size

        sw = parent.winfo_screenwidth()
        sh = parent.winfo_screenheight()
        self._s  = min((sw * 0.85) / self.img_w, (sh * 0.80) / self.img_h, 1.0)
        self._dw = int(self.img_w * self._s)
        self._dh = int(self.img_h * self._s)

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

        self.cv.bind("<ButtonPress-1>",   self._on_press)
        self.cv.bind("<B1-Motion>",        self._on_drag)
        self.cv.bind("<ButtonRelease-1>",  self._on_release)
        self.cv.bind("<Motion>",           self._on_motion)
        self.cv.bind("<Double-Button-1>",  self._on_dblclick)
        self.win.bind("<Escape>",          lambda _: self._deselect())
        self.win.bind("<Delete>",          self._delete_selected)

        self.win.update_idletasks()
        x = (sw - self.win.winfo_width())  // 2
        y = (sh - self.win.winfo_height()) // 2
        self.win.geometry(f"+{x}+{y}")

        parent.wait_window(self.win)

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _build_toolbar(self):
        tb = tk.Frame(self.win, bg="#2d2d30", pady=5)
        tb.pack(fill=tk.X)

        self._tbtn: dict[str, tk.Button] = {}
        for tid, lbl in [("select", "↖  Auswahl"),
                         ("rect",   "▭  Rahmen"),
                         ("arrow",  "→  Pfeil"),
                         ("text",   "T  Text"),
                         ("number", "①  Nr."),
                         ("blur",   "⊘  Blur")]:
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

        tk.Button(tb, text="↩  Rückgängig", bg="#3c3c3c", fg="white",
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

        self._tool = "rect"
        for tid, btn in self._tbtn.items():
            btn.config(bg="#0078d4" if tid == "rect" else "#3c3c3c")

    def _select_tool(self, tool: str):
        self._tool = tool
        for tid, btn in self._tbtn.items():
            btn.config(bg="#0078d4" if tid == tool else "#3c3c3c")
        if not hasattr(self, "cv"):
            return
        if tool == "select":
            self.cv.config(cursor="arrow")
        else:
            self.cv.config(cursor="crosshair")
            self._deselect()

    def _pick_color(self):
        _, hex_color = colorchooser.askcolor(color=self._color, title="Farbe wählen")
        if hex_color:
            self._color = hex_color
            self._cbtn.config(bg=hex_color)

    # ── Koordinaten-Hilfe ─────────────────────────────────────────────────────

    def _to_img(self, cx: int, cy: int) -> tuple[int, int]:
        return int(cx / self._s), int(cy / self._s)

    def _to_cv(self, ix: int, iy: int) -> tuple[int, int]:
        return int(ix * self._s), int(iy * self._s)

    # ── Canvas-Basisbild ──────────────────────────────────────────────────────

    def _refresh_base(self):
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
        for ann in self._annotations:
            if ann["type"] == "blur":
                self._redraw_annotation_item(ann)

    # ── Griffpunkte ───────────────────────────────────────────────────────────

    def _get_handle_positions(self, ann: dict) -> dict[str, tuple[int, int]]:
        t = ann["type"]
        if t in ("rect", "blur"):
            x0, y0, x1, y1 = ann["coords"] if t == "rect" else ann["region"]
            cx0, cy0 = self._to_cv(x0, y0)
            cx1, cy1 = self._to_cv(x1, y1)
            lx, rx = min(cx0, cx1), max(cx0, cx1)
            ty, by = min(cy0, cy1), max(cy0, cy1)
            mx, my = (lx + rx) // 2, (ty + by) // 2
            return {
                "nw": (lx, ty), "n": (mx, ty), "ne": (rx, ty),
                "w":  (lx, my),                "e":  (rx, my),
                "sw": (lx, by), "s": (mx, by), "se": (rx, by),
                "move": (mx, my),
            }
        elif t == "arrow":
            ix0, iy0, ix1, iy1 = ann["coords"]
            s_cv = self._to_cv(ix0, iy0)
            e_cv = self._to_cv(ix1, iy1)
            return {
                "start": s_cv,
                "end":   e_cv,
                "move":  ((s_cv[0] + e_cv[0]) // 2, (s_cv[1] + e_cv[1]) // 2),
            }
        elif t == "text":
            cx, cy = self._to_cv(*ann["coords"])
            fs_cv = max(11, int(13 * self._s))
            est_w = max(20, int(len(ann.get("text", "")) * fs_cv * 0.65)) + 8
            est_h = int(fs_cv * 1.6) + 4
            return {"move": (cx + est_w // 2, cy + est_h // 2)}
        elif t == "number":
            cx, cy = self._to_cv(*ann["coords"])
            return {"move": (cx, cy)}
        return {}

    def _draw_selection(self, idx: int):
        self._clear_selection_handles()
        if idx < 0 or idx >= len(self._annotations):
            return
        ann     = self._annotations[idx]
        handles = self._get_handle_positions(ann)
        r = _HANDLE_R
        for key, (cx, cy) in handles.items():
            if key == "move":
                item = self.cv.create_oval(
                    cx - r, cy - r, cx + r, cy + r,
                    fill="#0078d4", outline="white", width=1.5, tags="sel_handle")
            else:
                item = self.cv.create_rectangle(
                    cx - r, cy - r, cx + r, cy + r,
                    fill="white", outline="#0078d4", width=1.5, tags="sel_handle")
            self._sel_handles.append((key, item))

    def _clear_selection_handles(self):
        for _, item in self._sel_handles:
            self.cv.delete(item)
        self._sel_handles = []

    def _deselect(self):
        self._selected_idx = -1
        self._clear_selection_handles()

    # ── Hit-Test ──────────────────────────────────────────────────────────────

    def _hit_handle(self, cx: int, cy: int) -> str:
        r = _HANDLE_R + 2
        for key, item in self._sel_handles:
            coords = self.cv.coords(item)
            if not coords:
                continue
            x0, y0, x1, y1 = coords
            if x0 - r <= cx <= x1 + r and y0 - r <= cy <= y1 + r:
                return key
        return ""

    def _hit_annotation(self, cx: int, cy: int) -> int:
        for i in range(len(self._annotations) - 1, -1, -1):
            ann = self._annotations[i]
            t   = ann["type"]
            if t in ("rect", "blur"):
                x0, y0, x1, y1 = ann["coords"] if t == "rect" else ann["region"]
                cvx0, cvy0 = self._to_cv(x0, y0)
                cvx1, cvy1 = self._to_cv(x1, y1)
                lx, rx = min(cvx0, cvx1) - 6, max(cvx0, cvx1) + 6
                ty, by = min(cvy0, cvy1) - 6, max(cvy0, cvy1) + 6
                if lx <= cx <= rx and ty <= cy <= by:
                    return i
            elif t == "arrow":
                ix0, iy0, ix1, iy1 = ann["coords"]
                acx0, acy0 = self._to_cv(ix0, iy0)
                acx1, acy1 = self._to_cv(ix1, iy1)
                if _point_to_segment_dist(cx, cy, acx0, acy0, acx1, acy1) < 10:
                    return i
            elif t == "text":
                tx, ty_img = ann["coords"]
                tcx, tcy = self._to_cv(tx, ty_img)
                fs_cv = max(11, int(13 * self._s))
                est_w = max(20, int(len(ann.get("text", "")) * fs_cv * 0.65)) + 8
                est_h = int(fs_cv * 1.6) + 4
                if tcx - 6 <= cx <= tcx + est_w and tcy - 6 <= cy <= tcy + est_h:
                    return i
            elif t == "number":
                ncx, ncy = self._to_cv(*ann["coords"])
                r = max(14, int(ann.get("radius", 14) * self._s)) + 6
                if math.hypot(cx - ncx, cy - ncy) <= r:
                    return i
        return -1

    # ── Cursor im Select-Modus ────────────────────────────────────────────────

    _RESIZE_CURSORS = {
        "nw": "size_nw_se", "se": "size_nw_se",
        "ne": "size_ne_sw", "sw": "size_ne_sw",
        "n":  "size_ns",    "s":  "size_ns",
        "w":  "size_we",    "e":  "size_we",
        "start": "crosshair", "end": "crosshair",
        "move": "fleur",
    }

    def _on_motion(self, e):
        if self._tool != "select":
            return
        if self._drag_mode:
            self.cv.config(cursor="fleur")
            return
        handle = self._hit_handle(e.x, e.y)
        if handle:
            self.cv.config(cursor=self._RESIZE_CURSORS.get(handle, "fleur"))
        elif self._hit_annotation(e.x, e.y) >= 0:
            self.cv.config(cursor="fleur")
        else:
            self.cv.config(cursor="arrow")

    # ── Maus-Events ───────────────────────────────────────────────────────────

    def _on_press(self, e):
        if self._tool == "select":
            self._sel_press(e.x, e.y)
            return
        if self._tool == "text":
            self._start_text(e.x, e.y)
            return
        if self._tool == "number":
            self._place_number(e.x, e.y)
            return
        self._deselect()
        self._drawing  = True
        self._start    = (e.x, e.y)
        self._cur_item = None

    def _on_drag(self, e):
        if self._tool == "select":
            self._sel_drag(e.x, e.y)
            return
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
                x0, y0, e.x, e.y, outline="#00cfff", width=2, dash=(5, 3))

    def _on_release(self, e):
        if self._tool == "select":
            self._sel_release()
            return
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
            self._annotations.append({"type": "blur", "region": region,
                                       "canvas_item": None})
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

    def _on_dblclick(self, e):
        """Canvas-level double-click: öffnet Inline-Edit für Text/Nummer."""
        if self._tool != "select" or self._selected_idx < 0:
            return
        ann = self._annotations[self._selected_idx]
        if ann["type"] in ("text", "number"):
            self._drag_mode = ""
            self._start_edit(ann)

    # ── Select-Drag-Logik ─────────────────────────────────────────────────────

    def _sel_press(self, cx: int, cy: int):
        if self._selected_idx >= 0:
            handle = self._hit_handle(cx, cy)
            if handle:
                self._drag_mode  = handle
                self._drag_start = (cx, cy)
                self._drag_orig  = self._get_coords(self._annotations[self._selected_idx])
                return
        idx = self._hit_annotation(cx, cy)
        if idx >= 0:
            self._selected_idx = idx
            self._drag_mode    = "move"
            self._drag_start   = (cx, cy)
            self._drag_orig    = self._get_coords(self._annotations[idx])
            self._draw_selection(idx)
        else:
            self._deselect()
            self._drag_mode = ""

    def _sel_drag(self, cx: int, cy: int):
        if self._selected_idx < 0 or not self._drag_mode:
            return
        ddx = (cx - self._drag_start[0]) / self._s
        ddy = (cy - self._drag_start[1]) / self._s
        ann = self._annotations[self._selected_idx]
        self._apply_transform(ann, ddx, ddy)
        self._redraw_annotation_item(ann)
        self._draw_selection(self._selected_idx)

    def _sel_release(self):
        if self._selected_idx >= 0:
            ann = self._annotations[self._selected_idx]
            if ann["type"] == "blur":
                self._refresh_base()
                self._draw_selection(self._selected_idx)
        self._drag_mode = ""
        self._drag_orig = None

    def _get_coords(self, ann: dict):
        t = ann["type"]
        if t in ("rect", "arrow"):
            return ann["coords"]
        if t == "blur":
            return ann["region"]
        if t in ("text", "number"):
            return ann["coords"]
        return None

    def _apply_transform(self, ann: dict, ddx: float, ddy: float):
        mode = self._drag_mode
        orig = self._drag_orig
        t    = ann["type"]

        if t in ("text", "number"):
            if mode == "move":
                ox, oy = orig
                ann["coords"] = (int(ox + ddx), int(oy + ddy))
            return

        key = "coords" if t in ("rect", "arrow") else "region"
        x0, y0, x1, y1 = orig
        dx, dy = int(ddx), int(ddy)

        if   mode == "move":  ann[key] = (x0+dx, y0+dy, x1+dx, y1+dy)
        elif mode == "nw":    ann[key] = (x0+dx, y0+dy, x1,    y1   )
        elif mode == "ne":    ann[key] = (x0,    y0+dy, x1+dx, y1   )
        elif mode == "sw":    ann[key] = (x0+dx, y0,    x1,    y1+dy)
        elif mode == "se":    ann[key] = (x0,    y0,    x1+dx, y1+dy)
        elif mode == "n":     ann[key] = (x0,    y0+dy, x1,    y1   )
        elif mode == "s":     ann[key] = (x0,    y0,    x1,    y1+dy)
        elif mode == "w":     ann[key] = (x0+dx, y0,    x1,    y1   )
        elif mode == "e":     ann[key] = (x0,    y0,    x1+dx, y1   )
        elif mode == "start": ann[key] = (x0+dx, y0+dy, x1,    y1   )
        elif mode == "end":   ann[key] = (x0,    y0,    x1+dx, y1+dy)

    def _redraw_annotation_item(self, ann: dict):
        if ann.get("canvas_item"):
            self.cv.delete(ann["canvas_item"])
            ann["canvas_item"] = None
        if ann.get("canvas_text"):
            self.cv.delete(ann["canvas_text"])
            ann["canvas_text"] = None
        t = ann["type"]
        if t == "rect":
            x0, y0, x1, y1 = ann["coords"]
            cx0, cy0 = self._to_cv(x0, y0)
            cx1, cy1 = self._to_cv(x1, y1)
            ann["canvas_item"] = self.cv.create_rectangle(
                cx0, cy0, cx1, cy1, outline=ann["color"], width=ann["width"])
        elif t == "arrow":
            x0, y0, x1, y1 = ann["coords"]
            cx0, cy0 = self._to_cv(x0, y0)
            cx1, cy1 = self._to_cv(x1, y1)
            ann["canvas_item"] = self.cv.create_line(
                cx0, cy0, cx1, cy1, fill=ann["color"], width=ann["width"],
                arrow=tk.LAST, arrowshape=(16, 20, 6))
        elif t == "text":
            x, y = ann["coords"]
            cx, cy = self._to_cv(x, y)
            fs_cv = max(11, int(13 * self._s))
            ann["canvas_item"] = self.cv.create_text(
                cx, cy, text=ann["text"], fill=ann["color"],
                font=("Segoe UI", fs_cv, "bold"), anchor="nw")
            self._bind_selectable_item(ann["canvas_item"], ann)
        elif t == "number":
            ix, iy = ann["coords"]
            cx, cy = self._to_cv(ix, iy)
            r = max(14, int(ann.get("radius", 14) * self._s))
            fs_cv = max(9, int(r * 0.9))
            ann["canvas_item"] = self.cv.create_oval(
                cx - r, cy - r, cx + r, cy + r,
                fill=ann["color"], outline=ann["color"])
            ann["canvas_text"] = self.cv.create_text(
                cx, cy, text=str(ann["number"]), fill="white",
                font=("Segoe UI", fs_cv, "bold"), anchor="center")
            self._bind_selectable_item(ann["canvas_item"], ann)
            self._bind_selectable_item(ann["canvas_text"], ann)
        elif t == "blur":
            x0, y0, x1, y1 = ann["region"]
            cx0, cy0 = self._to_cv(x0, y0)
            cx1, cy1 = self._to_cv(x1, y1)
            ann["canvas_item"] = self.cv.create_rectangle(
                cx0, cy0, cx1, cy1, outline="#00cfff", width=1, dash=(4, 3))

    # ── Nummerierungs-Tool ────────────────────────────────────────────────────

    def _place_number(self, cx: int, cy: int):
        ix, iy = self._to_img(cx, cy)
        num = self._next_number()
        ann = {
            "type":        "number",
            "number":      num,
            "coords":      (ix, iy),
            "radius":      int(14 / self._s),
            "color":       self._color,
            "canvas_item": None,
            "canvas_text": None,
        }
        self._annotations.append(ann)
        self._redraw_annotation_item(ann)

    def _next_number(self) -> int:
        existing = [a["number"] for a in self._annotations if a["type"] == "number"]
        return max(existing, default=0) + 1

    # ── Inline-Edit (Text & Nummer) ───────────────────────────────────────────

    def _start_edit(self, ann: dict):
        if self._edit_entry:
            return
        t = ann["type"]
        cx, cy = self._to_cv(*ann["coords"])
        if t == "number":
            r = max(14, int(ann.get("radius", 14) * self._s))
            cx -= r
            cy -= r
        fs = max(11, int(13 * self._s))
        initial = ann["text"] if t == "text" else str(ann["number"])
        ent = tk.Entry(self.cv, fg=ann["color"],
                       font=("Segoe UI", fs, "bold"),
                       bg="#fffde7", relief=tk.SOLID, bd=1,
                       insertbackground=ann["color"],
                       width=max(6, len(initial) + 2))
        ent.insert(0, initial)
        ent.select_range(0, tk.END)
        self.cv.create_window(cx, cy, window=ent, anchor="nw", tags="edit_te")
        ent.focus_set()
        self._edit_entry = ent
        self._edit_ann   = ann

        def commit(_=None):
            val = ent.get().strip()
            self.cv.delete("edit_te")
            ent.destroy()
            self._edit_entry = None
            self._edit_ann   = None
            if not val:
                return
            if t == "text":
                ann["text"] = val
            else:
                try:
                    ann["number"] = int(val)
                except ValueError:
                    return
            self._deselect()
            self._redraw_annotation_item(ann)

        def cancel(_=None):
            self.cv.delete("edit_te")
            ent.destroy()
            self._edit_entry = None
            self._edit_ann   = None

        ent.bind("<Return>", commit)
        ent.bind("<Escape>", cancel)

    # ── Löschen mit Entf-Taste ────────────────────────────────────────────────

    def _delete_selected(self, event=None):
        if self._tool != "select" or self._selected_idx < 0:
            return
        idx = self._selected_idx
        ann = self._annotations[idx]
        self._deselect()
        if ann.get("canvas_item"):
            self.cv.delete(ann["canvas_item"])
        if ann.get("canvas_text"):
            self.cv.delete(ann["canvas_text"])
        self._annotations.pop(idx)
        if ann["type"] == "blur":
            self._refresh_base()

    # ── Selektierbare Items binden ────────────────────────────────────────────

    def _bind_selectable_item(self, item_id: int, ann: dict):
        def on_press(e):
            if self._tool != "select":
                return
            if self._selected_idx >= 0 and self._hit_handle(e.x, e.y):
                return
            try:
                idx = next(i for i, a in enumerate(self._annotations) if a is ann)
            except StopIteration:
                return
            self._selected_idx = idx
            self._drag_mode    = "move"
            self._drag_start   = (e.x, e.y)
            self._drag_orig    = self._get_coords(ann)
            self._draw_selection(idx)
            return "break"

        self.cv.tag_bind(item_id, "<ButtonPress-1>", on_press)

    # ── Text-Tool ─────────────────────────────────────────────────────────────

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
                                     font=("Segoe UI", fs_cv, "bold"), anchor="nw")
        ix, iy = self._to_img(cx, cy)
        self._annotations.append({
            "type":        "text",
            "text":        txt,
            "coords":      (ix, iy),
            "color":       self._color,
            "font_size":   max(14, int(14 / self._s)),
            "canvas_item": item,
        })
        self._bind_selectable_item(item, self._annotations[-1])

    def _cancel_text(self):
        if self._text_entry:
            self.cv.delete("te")
            self._text_entry.destroy()
            self._text_entry = None

    # ── Rückgängig ────────────────────────────────────────────────────────────

    def _undo(self):
        if not self._annotations:
            return
        self._deselect()
        ann = self._annotations.pop()
        if ann.get("canvas_item"):
            self.cv.delete(ann["canvas_item"])
        if ann.get("canvas_text"):
            self.cv.delete(ann["canvas_text"])
        if ann["type"] == "blur":
            self._refresh_base()

    # ── Export ────────────────────────────────────────────────────────────────

    def _accept(self):
        self._commit_text()
        img  = self.orig_img.copy()
        draw = ImageDraw.Draw(img)

        for ann in self._annotations:
            t = ann["type"]

            if t == "blur":
                x0, y0, x1, y1 = ann["region"]
                x0, y0, x1, y1 = min(x0,x1), min(y0,y1), max(x0,x1), max(y0,y1)
                region = img.crop((x0, y0, x1, y1))
                region = region.filter(ImageFilter.GaussianBlur(radius=18))
                img.paste(region, (x0, y0))
                draw = ImageDraw.Draw(img)

            elif t == "rect":
                x0, y0, x1, y1 = ann["coords"]
                draw.rectangle(
                    [min(x0,x1), min(y0,y1), max(x0,x1), max(y0,y1)],
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
                draw.text((x, y), ann["text"], fill=ann["color"], font=font)

            elif t == "number":
                cx, cy = ann["coords"]
                r = ann.get("radius", 14)
                draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                             fill=ann["color"], outline=ann["color"])
                font = _load_font(max(10, int(r * 1.2)))
                draw.text((cx, cy), str(ann["number"]),
                          fill="white", font=font, anchor="mm")

        buf = BytesIO()
        img.save(buf, format="JPEG", quality=92)
        self.result_b64 = base64.b64encode(buf.getvalue()).decode()
        self.win.destroy()

    def _cancel(self):
        self._cancel_text()
        self.result_b64 = None
        self.win.destroy()


# ── Font-Hilfe ────────────────────────────────────────────────────────────────

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