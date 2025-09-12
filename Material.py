#!/usr/bin/env python3
from __future__ import annotations
import os, sys, sqlite3, hashlib, datetime, platform, subprocess, tempfile, threading, time
from dataclasses import dataclass
from typing import List, Optional, Tuple
import ctypes
import hmac  # NEW: for constant-time comparison
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from reportlab.lib.units import inch
from reportlab.pdfgen import canvas as pdfcanvas

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError:
    Image = ImageDraw = ImageFont = ImageTk = None

APP_TITLE = "Mva Label Printing Software"

# Print settings (PDF-only)
SUMATRA_PRINTER_NAME: str | None = None
SUMATRA_PAPER: str | None = None            # e.g., "Letter" to let Sumatra set host paper
SUMATRA_SCALE: str = "fit"                  # "noscale", "fit", "shrink"
SUMATRA_ORIENTATION: str | None = None      # "portrait", "landscape", or None

# Host page (when we wrap the 4x6 onto Letter/A4 ourselves)
HOST_PAPER_NAME: str = "Letter"
HOST_SCALE_MODE: str = "fit"

PERSIST_PDF_SECONDS: int = 25
USE_SUMATRA_32BIT_FIRST: bool = True
# How to send multiple copies via Sumatra:
#   'nx'   -> one job with the Nx token (e.g., '3x'). Some drivers ignore this.
#   'loop' -> send N separate one-copy jobs (most reliable across drivers).
SUMATRA_COPIES_MODE: str = "loop"

def _app_dir() -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "executable"):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def _resource_path(relpath: str) -> str:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, relpath)
    alt1 = os.path.join(_app_dir(), relpath)
    return alt1 if os.path.exists(alt1) else os.path.join(os.path.dirname(os.path.abspath(__file__)), relpath)

DB_PATH = os.path.join(_app_dir(), "labels.db")

ADMIN_PASSWORD_HASH = "16cb31feec45070c0f9c07e033a9aab6b57fc13925399fdee1637662623d6304"

# 4x6 content inches
LABEL_WIDTH_IN, LABEL_HEIGHT_IN, MARGIN_IN = 4.0, 6.0, 0.25

# Typography & layout
CODE_FONT_SIZE = 220
CODE_SIDE_MARGIN_IN = 0.08
MAX_DESC_FONT_SIZE = 52
MIN_DESC_FONT_SIZE = 28
UNDERLINE_GAP_PT = 3
UNDERLINE_STROKE_PT = 3
DESC_TOP_FRACTION = 0.60
DESC_LINE_SPACING = 1.22
DESC_ALIGN_CENTER = True

HOST_PAPER_SIZES = {
    "Letter": (8.5*inch, 11*inch),
    "A4":     (210/25.4*inch, 297/25.4*inch),
}

# Tiny base64 PNG fallback icon (16px). For a real icon, ship assets/app.ico
_APP_ICON_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAeFBMVEUA"
    "AAB/f3+AgICSkpKampqVlZWenp6RkZGurq6dnZ2Xl5e4uLi9vb3Pz8/Hx8e3t7eioqKtra2YmJi2tra5ubmnp6ezs7PExMS7u7vAwMDQ0NC+vr6jo6OJiYnFxcW/v7+qqqqUlJSP"
    "j49ra2s8PDzc3Nz////c53y2AAAAKXRSTlMAD6P0o8wK6p1fWgR8XH2mP3VQ6x4sKQGJqvU0GmXj9bQ3s1iY4B3bCkq8Lh1UZwAAABYSURBVBjTY2BgZGJmYGBgYGBQAkYGJgYm"
    "BkY2BlYQ0CqGAEYoAUYbCAGQ5QFQwYo6GAAEwKQwQwNQ0QwA4gGgYjAAE2QmGk0gWgC4mQKkc0E6QBi2pAxQFQjEAAI0gQy1mDq7AAAAAElFTkSuQmCC"
)

# ---------- Data ----------
@dataclass
class LabelRow:
    id: int
    name: str
    description: str
    created_at: str

# ---------- DB ----------
class LabelDB:
    def __init__(self, path: str):
        self.path = path
        self._init_db()

    def _connect(self):
        return sqlite3.connect(self.path)

    def _init_db(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with self._connect() as con:
            cur = con.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS labels (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    description TEXT NOT NULL,
                    created_at TEXT NOT NULL
                )
                """
            )
            con.commit()

    def add_label(self, name: str, description: str) -> int:
        with self._connect() as con:
            cur = con.cursor()
            cur.execute(
                "INSERT INTO labels(name, description, created_at) VALUES(?,?,?)",
                (name.strip(), description.strip(), datetime.datetime.now().isoformat(timespec="seconds")),
            )
            con.commit()
            return cur.lastrowid

    def get_all_labels(self) -> List[LabelRow]:
        with self._connect() as con:
            cur = con.cursor()
            cur.execute("SELECT id, name, description, created_at FROM labels ORDER BY name COLLATE NOCASE ASC")
            rows = cur.fetchall()
            return [LabelRow(*r) for r in rows]

    def get_label_by_name(self, name: str) -> Optional[LabelRow]:
        with self._connect() as con:
            cur = con.cursor()
            cur.execute("SELECT id, name, description, created_at FROM labels WHERE name = ?", (name.strip(),))
            r = cur.fetchone()
            return LabelRow(*r) if r else None

    def delete_by_ids(self, ids: list[int]) -> int:
        if not ids: return 0
        with self._connect() as con:
            cur = con.cursor()
            cur.executemany("DELETE FROM labels WHERE id = ?", [(i,) for i in ids])
            con.commit()
            return cur.rowcount

# ---------- PDF Renderer ----------
class LabelRenderer:
    def __init__(self, width_in: float, height_in: float, margin_in: float):
        self.content_w_pt = int(width_in * inch)
        self.content_h_pt = int(height_in * inch)
        self.margin_pt = int(margin_in * inch)
        self.code_side_margin_pt = int(CODE_SIDE_MARGIN_IN * inch)
        self.code_font_name = "Helvetica-Bold"
        self.desc_font_name = "Helvetica"

    def _fit_font(self, canv, text, target_width, max_size, min_size=18) -> int:
        size = max_size
        while size >= min_size:
            if canv.stringWidth(text, self.code_font_name, size) <= target_width:
                return size
            size -= 1
        return min_size

    def _wrap_text(self, canv, text, font_name, font_size, max_width) -> List[str]:
        words = text.replace("\r", "").split()
        if not words: return [""]
        lines, cur = [], words[0]
        for w in words[1:]:
            test = f"{cur} {w}"
            if canv.stringWidth(test, font_name, font_size) <= max_width:
                cur = test
            else:
                lines.append(cur); cur = w
        lines.append(cur)
        return lines

    def _fit_paragraph(self, canv, text, max_width, max_height) -> tuple[int, list[str]]:
        size = MAX_DESC_FONT_SIZE
        while size >= MIN_DESC_FONT_SIZE:
            lines = []
            for raw in (text.replace("\r", "").split("\n") or [""]):
                lines.extend(self._wrap_text(canv, raw, self.desc_font_name, size, max_width))
            line_h = size * DESC_LINE_SPACING
            if len(lines) * line_h <= max_height:
                return size, lines
            size -= 1
        lines = []
        for raw in (text.replace("\r", "").split("\n") or [""]):
            lines.extend(self._wrap_text(canv, raw, self.desc_font_name, MIN_DESC_FONT_SIZE, max_width))
        return MIN_DESC_FONT_SIZE, lines

    def _draw_label_content(self, c, code: str, description: str):
        w, h, m = self.content_w_pt, self.content_h_pt, self.margin_pt
        # Code
        max_code_width = w - (2 * self.code_side_margin_pt)
        code_size = self._fit_font(c, code, max_code_width, CODE_FONT_SIZE, min_size=28)
        c.setFont(self.code_font_name, code_size)
        ascent = 0.80 * code_size
        code_x, code_y = w / 2.0, h - 1 - ascent
        c.drawCentredString(code_x, code_y, code)
        # underline
        code_text_width = c.stringWidth(code, self.code_font_name, code_size)
        underline_y = code_y - UNDERLINE_GAP_PT
        c.setLineWidth(UNDERLINE_STROKE_PT)
        c.line(code_x - code_text_width / 2.0, underline_y, code_x + code_text_width / 2.0, underline_y)
        # Description
        block_w = w - 2 * m
        max_desc_height = (h * DESC_TOP_FRACTION) - m
        desc_size, lines = self._fit_paragraph(c, description, block_w, max_desc_height)
        c.setFont(self.desc_font_name, desc_size)
        line_h = desc_size * DESC_LINE_SPACING
        cur_y = h * DESC_TOP_FRACTION
        for ln in lines:
            (c.drawCentredString(w/2.0, cur_y, ln) if DESC_ALIGN_CENTER else c.drawString(m, cur_y, ln))
            cur_y -= line_h

    def render_pdf(self, code: str, description: str, out_path: str, host_wrap: bool, host_name: str, host_scale_mode: str):
        if not host_wrap:
            c = pdfcanvas.Canvas(out_path, pagesize=(self.content_w_pt, self.content_h_pt))
            self._draw_label_content(c, code, description); c.showPage(); c.save(); return
        host_w, host_h = HOST_PAPER_SIZES.get(host_name, HOST_PAPER_SIZES["Letter"])
        c = pdfcanvas.Canvas(out_path, pagesize=(host_w, host_h))
        host_margin_pt = 0.25 * inch
        avail_w = max(1, host_w - 2 * host_margin_pt)
        avail_h = max(1, host_h - 2 * host_margin_pt)
        cw, ch = self.content_w_pt, self.content_h_pt
        def best_fit(unrotated: bool):
            if unrotated:
                s = min(avail_w / cw, avail_h / ch) if host_scale_mode == "fit" else 1.0
                return s, (host_w - cw*s)/2.0, (host_h - ch*s)/2.0, cw*s, ch*s
            else:
                s = min(avail_w / ch, avail_h / cw) if host_scale_mode == "fit" else 1.0
                return s, (host_w - ch*s)/2.0, (host_h - cw*s)/2.0, ch*s, cw*s
        s0, x0, y0, w0, h0 = best_fit(True)
        s1, x1, y1, w1, h1 = best_fit(False)
        rotate = (w1 * h1) > (w0 * h0)
        if not rotate:
            c.saveState(); c.translate(x0, y0); c.scale(s0, s0); self._draw_label_content(c, code, description); c.restoreState()
        else:
            c.saveState(); c.translate(x1, y1); c.rotate(90); c.translate(0, -cw * s1); c.scale(s1, s1)
            self._draw_label_content(c, code, description); c.restoreState()
        c.showPage(); c.save()

# ---------- Helpers: DPI, AppID, Icon ----------
def _set_dpi_awareness():
    if platform.system().lower() == "windows":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try: ctypes.windll.user32.SetProcessDPIAware()
            except Exception: pass

def _set_appusermodel_id():
    if platform.system().lower() == "windows":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MaterialLabelPrinter.PhillipYoung.1.0")
        except Exception:
            pass

def _load_app_icon(root: tk.Tk):
    try:
        ico = _resource_path("assets/app.ico")
        if os.path.exists(ico):
            try:
                root.iconbitmap(default=ico)
                return
            except Exception:
                pass
    except Exception:
        pass
    try:
        img = tk.PhotoImage(data=_APP_ICON_B64)
        root.iconphoto(True, img)
    except Exception:
        pass

# ---------- App ----------
class App(tk.Tk):
    def __init__(self):
        _set_dpi_awareness()
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x680")
        self.minsize(980, 650)
        self.resizable(True, True)
        _load_app_icon(self)

        self.style = ttk.Style()
        self._apply_style(self.style)

        self.db = LabelDB(DB_PATH)
        self.renderer = LabelRenderer(LABEL_WIDTH_IN, LABEL_HEIGHT_IN, MARGIN_IN)

        self._build_ui()
        self._bind_shortcuts()
        self._reload_labels()

    def _apply_style(self, style: ttk.Style):
        try: style.theme_use("clam")
        except Exception: pass
        style.configure("TFrame", padding=0)
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Treeview", rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        style.configure("Card.TLabelframe.Label", font=("Segoe UI Semibold", 10))
        style.configure("Status.TLabel", foreground="#555")

    def _build_ui(self):
        pad = 12
        header = ttk.Frame(self); header.pack(fill=tk.X, padx=pad, pady=(pad, 8))
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI Semibold", 14)).pack(side=tk.LEFT)

        top = ttk.Frame(self); top.pack(fill=tk.X, padx=pad, pady=(0, 10))
        ttk.Label(top, text="Material:").pack(side=tk.LEFT, padx=(0,6))
        self.label_var = tk.StringVar()
        self.combo = ttk.Combobox(top, textvariable=self.label_var, state="readonly", width=50, height=15)
        self.combo.pack(side=tk.LEFT, padx=(0, 10))
        self.combo.bind("<<ComboboxSelected>>", lambda e: self._on_selection())

        ttk.Label(top, text="Copies:").pack(side=tk.LEFT)
        self.copies_var = tk.IntVar(value=1)
        self.copies_spin = ttk.Spinbox(top, from_=1, to=999, width=5, textvariable=self.copies_var, justify="center")
        self.copies_spin.pack(side=tk.LEFT, padx=(6, 10))

        ttk.Button(top, text="Print (Ctrl+P)", command=self._print_selected).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(top, text="Admin", command=self._admin_login).pack(side=tk.LEFT)

        mid = ttk.Frame(self); mid.pack(fill=tk.BOTH, expand=True, padx=pad, pady=0)

        prev_card = ttk.Labelframe(mid, text="Print Preview", style="Card.TLabelframe")
        prev_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10), pady=(0,8))
        prev_frame = ttk.Frame(prev_card); prev_frame.pack(padx=10, pady=10)
        self.preview_w, self.preview_h = 380, 570
        self.preview_canvas = tk.Canvas(prev_frame, width=self.preview_w, height=self.preview_h, bg="white",
                                        highlightthickness=1, highlightbackground="#ddd")
        self.preview_canvas.pack()

        right = ttk.Frame(mid); right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0,8))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, anchor="w", style="Status.TLabel").pack(fill=tk.X, padx=pad, pady=(6, pad))

    def _bind_shortcuts(self):
        self.bind_all("<Control-p>", lambda e: self._print_selected())
        self.bind_all("<Control-P>", lambda e: self._print_selected())

    # Data ops
    def _reload_labels(self):
        rows = self.db.get_all_labels()
        names = [r.name for r in rows]
        self.combo["values"] = tuple(names)
        if rows:
            cur = self.label_var.get()
            self.combo.set(cur if cur in names else names[0])
            self._on_selection()
        else:
            self.combo.set("")
            self.preview_canvas.delete("all")
            self.status_var.set("No labels yet.")

    # Admin
    def _admin_login(self):
        pw = simpledialog.askstring("Admin Login", "Enter admin password:", show='*', parent=self)
        if pw is None:
            return  # user canceled
        entered = hashlib.sha256(pw.encode("utf-8")).hexdigest()
        if hmac.compare_digest(entered, ADMIN_PASSWORD_HASH):
            self._open_admin_panel()
        else:
            messagebox.showerror("Access denied", "Incorrect password.")

    def _open_admin_panel(self):
        win = tk.Toplevel(self); win.title("Admin Panel"); win.geometry("400x320"); win.grab_set(); _load_app_icon(win)
        frm = ttk.Frame(win, padding=12); frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Admin actions", font=("Segoe UI Semibold", 11)).pack(anchor="w", pady=(0, 8))
        ttk.Button(frm, text="Add Single Label", command=self._open_add_single).pack(fill=tk.X, pady=6)
        ttk.Button(frm, text="View / Delete Labels", command=self._open_view_list).pack(fill=tk.X, pady=6)
        ttk.Button(frm, text="Bulk Import Labels", command=self._open_bulk_import).pack(fill=tk.X, pady=6)

    def _open_add_single(self):
        win = tk.Toplevel(self); win.title("Add Single Label"); win.geometry("560x380"); win.grab_set(); _load_app_icon(win)
        frm = ttk.Frame(win, padding=12); frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Label code (e.g., VG0100)").grid(row=0, column=0, sticky="w")
        code_var = tk.StringVar(); code_entry = ttk.Entry(frm, textvariable=code_var, width=44)
        code_entry.grid(row=1, column=0, sticky="we", pady=(0,10)); code_entry.focus_set()
        ttk.Label(frm, text="Description (multi-line)").grid(row=2, column=0, sticky="w")
        desc_txt = tk.Text(frm, height=10, wrap=tk.WORD, font=("Segoe UI", 10))
        desc_txt.grid(row=3, column=0, sticky="nsew"); frm.rowconfigure(3, weight=1); frm.columnconfigure(0, weight=1)
        btns = ttk.Frame(frm); btns.grid(row=4, column=0, sticky="e", pady=(12,0))
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side=tk.RIGHT, padx=(0,8))
        ttk.Button(btns, text="Save", command=lambda: self._save_single(code_var, desc_txt, win)).pack(side=tk.RIGHT)

    def _save_single(self, code_var, desc_txt, win):
        code = (code_var.get() or "").strip()
        desc = (desc_txt.get("1.0", tk.END) or "").strip()
        if not code:
            messagebox.showwarning("Required", "Please enter a label code."); return
        if not desc:
            messagebox.showwarning("Required", "Please enter a description."); return
        try:
            self.db.add_label(code, desc)
        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicate", f"A label named '{code}' already exists."); return
        self._reload_labels()
        messagebox.showinfo("Saved", f"Added label: {code}")
        win.destroy()

    def _open_view_list(self):
        win = tk.Toplevel(self); win.title("Current Labels — View / Delete"); win.geometry("1000x640"); win.grab_set(); _load_app_icon(win)
        outer = ttk.Frame(win, padding=12); outer.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(outer); top.pack(fill=tk.X, pady=(0,8))
        ttk.Label(top, text="Filter:").pack(side=tk.LEFT)
        qvar = tk.StringVar(); qentry = ttk.Entry(top, textvariable=qvar, width=50); qentry.pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Search", command=lambda: refresh()).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(top, text="Delete Selected", command=lambda: delete_selected()).pack(side=tk.LEFT, padx=(12, 0))

        cols = ("Code", "Description")
        tree = ttk.Treeview(outer, columns=cols, show="headings", selectmode="extended")
        tree.heading("Code", text="Code"); tree.heading("Description", text="Description")
        tree.column("Code", width=200, anchor="w"); tree.column("Description", width=760, anchor="w")
        tree.pack(fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set); yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        def refresh():
            for row in tree.get_children(): tree.delete(row)
            q = (qvar.get() or "").lower()
            for r in self.db.get_all_labels():
                if q and q not in r.name.lower() and q not in r.description.lower(): continue
                tree.insert("", tk.END, iid=str(r.id), values=(r.name, r.description))

        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Nothing selected", "Select one or more rows to delete."); return
            ids = []
            for iid in sel:
                try:
                    ids.append(int(iid))
                except ValueError:
                    vals = tree.item(iid, "values")
                    if vals:
                        name = str(vals[0]).strip(); row = self.db.get_label_by_name(name)
                        if row: ids.append(row.id)
            label = f"these {len(ids)} labels" if len(ids) > 1 else "this label"
            if not messagebox.askyesno("Confirm delete", f"Permanently delete {label}? This cannot be undone."): return
            deleted = self.db.delete_by_ids(ids)
            refresh()
            self._reload_labels()
            self.status_var.set(f"Deleted {deleted} item(s).")

        def on_key(event):
            if event.keysym == "Delete": delete_selected()

        tree.bind("<Key>", on_key)
        refresh()

    def _open_bulk_import(self):
        win = tk.Toplevel(self); win.title("Bulk Import Labels"); win.geometry("780x560"); win.grab_set(); _load_app_icon(win)
        outer = ttk.Frame(win, padding=12); outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(outer, text="Paste lines like:  VG0100  Description here").pack(anchor="w", pady=(0,6))
        txt = tk.Text(outer, wrap=tk.NONE, font=("Segoe UI", 10)); txt.pack(fill=tk.BOTH, expand=True)
        ttk.Button(outer, text="Import", command=lambda: do_import()).pack(pady=10, anchor="e")

        def do_import():
            pairs = self._parse_bulk_lines(txt.get("1.0", tk.END))
            added = skipped = 0
            for name, desc in pairs:
                try:
                    self.db.add_label(name, desc); added += 1
                except sqlite3.IntegrityError:
                    skipped += 1
            self._reload_labels()
            messagebox.showinfo("Import done", f"Added: {added}, Skipped: {skipped}")
            win.destroy()

    def _parse_bulk_lines(self, blob: str) -> List[Tuple[str, str]]:
        pairs: List[Tuple[str, str]] = []
        for raw in blob.splitlines():
            line = raw.strip()
            if not line: continue
            try: code, rest = line.split(None, 1)
            except ValueError: code, rest = line, ""
            pairs.append((code.strip(), rest.strip()))
        return pairs

    # ---- Preview ----
    def _on_selection(self):
        name = (self.label_var.get() or "").strip()
        if not name: return
        row = self.db.get_label_by_name(name)
        if row:
            self._render_preview(row.name, row.description)
            self.status_var.set(f"Loaded: {row.name}")

    def _pick_font(self, size: int, bold: bool = False):
        if ImageFont is None: return None
        candidates = [
            "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/segoeuib.ttf" if bold else None,
        ]
        for p in candidates:
            if p and os.path.exists(p):
                try: return ImageFont.truetype(p, size)
                except Exception: pass
        try: return ImageFont.load_default()
        except Exception: return None

    def _wrap_pillow(self, draw: 'ImageDraw.ImageDraw', text: str, font, max_width: int) -> List[str]:
        words = text.split()
        if not words: return [""]
        lines, line = [], words[0]
        for w in words[1:]:
            test = line + " " + w
            if draw.textlength(test, font=font) <= max_width: line = test
            else: lines.append(line); line = w
        lines.append(line); return lines

    def _render_preview(self, code: str, description: str):
        self.preview_canvas.delete("all")
        code_margin_px = int((CODE_SIDE_MARGIN_IN / LABEL_WIDTH_IN) * self.preview_w)

        if Image is None:
            self.preview_canvas.create_text(self.preview_w//2, 12, text=code, font=("Segoe UI", 36, "bold"),
                                            anchor="n", width=self.preview_w - 2*code_margin_px)
            self.preview_canvas.create_line(code_margin_px, 70, self.preview_w - code_margin_px, 70, width=3)
            usable_w = int(self.preview_w * 0.88); x_center = self.preview_w//2
            y = int(self.preview_h * DESC_TOP_FRACTION)
            for ln in (description.splitlines() or [""]):
                self.preview_canvas.create_text(x_center, y, text=ln, font=("Segoe UI", 24),
                                                width=usable_w, anchor="n")
                y += int(24 * DESC_LINE_SPACING)
            return

        img = Image.new("RGB", (self.preview_w, self.preview_h), "white")
        d = ImageDraw.Draw(img)
        start_guess = max(28, int(60 * (CODE_FONT_SIZE / 180.0)))
        code_font = self._pick_font(start_guess, bold=True)
        target_w = self.preview_w - 2 * code_margin_px
        if code_font:
            s = start_guess
            while s >= 18:
                f = self._pick_font(s, bold=True)
                if f and d.textlength(code, font=f) <= target_w: code_font = f; break
                s -= 1
        bbox = d.textbbox((0, 0), code, font=code_font)
        code_w, code_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        code_y_top = 10; code_x_left = max(code_margin_px, (self.preview_w - code_w) // 2)
        d.text((code_x_left, code_y_top), code, font=code_font, fill="black")
        ul_y = code_y_top + code_h + max(2, int((UNDERLINE_GAP_PT / 72.0) * 96))
        ul_left = max(code_margin_px, (self.preview_w - code_w) // 2)
        d.line([(ul_left, ul_y), (ul_left + code_w, ul_y)], width=max(2, int((UNDERLINE_STROKE_PT / 72.0) * 96)))
        lines: List[str] = []
        usable_w = int(self.preview_w * 0.88)
        desc_font = self._pick_font(32, bold=False)
        for raw in (description.splitlines() or [""]):
            lines.extend(self._wrap_pillow(d, raw, desc_font, usable_w))
        line_h = int(desc_font.size * DESC_LINE_SPACING)
        y = int(self.preview_h * DESC_TOP_FRACTION)
        for ln in lines:
            tw = d.textlength(ln, font=desc_font)
            x = (self.preview_w - tw) // 2
            d.text((x, y), ln, font=desc_font, fill="black")
            y += line_h
        self._preview_imgtk = ImageTk.PhotoImage(img)
        self.preview_canvas.create_image(0, 0, anchor="nw", image=self._preview_imgtk)

    # ---------- Sumatra (silent PDF printing) ----------
    def _find_sumatra(self) -> Optional[str]:
        cands = []
        if USE_SUMATRA_32BIT_FIRST:
            cands += [ _resource_path("assets/SumatraPDF-32.exe"),
                       _resource_path("assets/SumatraPDF.exe"),
                       r"C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe",
                       r"C:\\Program Files\\SumatraPDF\\SumatraPDF.exe" ]
        else:
            cands += [ _resource_path("assets/SumatraPDF.exe"),
                       _resource_path("assets/SumatraPDF-32.exe"),
                       r"C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
                       r"C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe" ]
        for p in cands:
            if os.path.isfile(p): return p
        return None

    def _silent_print_pdf(self, pdf_path: str, copies: int) -> bool:
        sp = self._find_sumatra()
        if not sp:
            self.status_var.set("SumatraPDF not found."); return False

        base_settings = []
        if SUMATRA_PAPER:
            base_settings.append(f"paper={SUMATRA_PAPER}")
        if SUMATRA_ORIENTATION in ("portrait", "landscape"):
            base_settings.append(SUMATRA_ORIENTATION)
        if SUMATRA_SCALE in ("noscale", "fit", "shrink"):
            base_settings.append(SUMATRA_SCALE)

        def launch_with_settings(settings_list: list[str]) -> int:
            settings_str = ",".join(settings_list) if settings_list else ""
            args = [sp]
            args += ["-print-to", SUMATRA_PRINTER_NAME] if SUMATRA_PRINTER_NAME else ["-print-to-default"]
            if settings_str:
                args += ["-print-settings", settings_str]
            args += [pdf_path]
            try:
                subprocess.Popen(
                    args + ["-silent", "-exit-on-print"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=0x08000000,  # CREATE_NO_WINDOW
                )
                return 0
            except Exception as e:
                self.status_var.set(f"Sumatra print failed: {e}")
                return 1

        c = max(1, int(copies or 1))
        if SUMATRA_COPIES_MODE == "nx":
            settings = [f"{c}x"] + base_settings
            rc = launch_with_settings(settings)
            return rc == 0
        else:
            for _ in range(c):
                rc = launch_with_settings(base_settings)
                if rc != 0:
                    return False
            return True

    # ---------- Print action (PDF only) ----------
    def _print_selected(self):
        if platform.system().lower() != "windows":
            messagebox.showerror("Unsupported OS", "This build is Windows-only."); return

        name = (self.label_var.get() or "").strip()
        if not name: return
        copies = max(1, int(self.copies_var.get() or 1))
        row = self.db.get_label_by_name(name)
        if not row: return

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = tmp.name

        try:
            host_wrap = (SUMATRA_PAPER is None)
            self.renderer.render_pdf(row.name, row.description, pdf_path,
                                     host_wrap=host_wrap, host_name=HOST_PAPER_NAME, host_scale_mode=HOST_SCALE_MODE)

            sent = self._silent_print_pdf(pdf_path, copies)
            if not sent:
                # Fallback: shell print (not silent)
                for _ in range(copies):
                    try:
                        os.startfile(pdf_path, "print")  # type: ignore[attr-defined]
                    except Exception:
                        subprocess.Popen(["cmd", "/c", "start", "", pdf_path])

            def _cleanup():
                deadline = time.time() + max(PERSIST_PDF_SECONDS, 25)
                while time.time() < deadline:
                    try:
                        os.remove(pdf_path); return
                    except Exception:
                        time.sleep(1.5)
                try: os.remove(pdf_path)
                except Exception: pass

            threading.Timer(25, _cleanup).start()
            self.status_var.set("Sent to printer (PDF).")
        except Exception as e:
            messagebox.showerror("Print error", f"Failed to print: {e}")
            try: os.remove(pdf_path)
            except Exception: pass

# ---------- main ----------
def main():
    _set_appusermodel_id()
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
