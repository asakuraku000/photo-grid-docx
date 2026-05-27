"""
A4 Photo Grid → DOCX Generator
Layouts: 1x1 (full page), 1x2, 1x3, 2x2, 3x3, 4x4
- Add images via file dialog or paste path
- Drag to reorder before generating
- Progress bar while building the DOCX
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import math

# ── Pillow (optional, for thumbnails) ────────────────────────────────────────
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ── python-docx ──────────────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Mm, Pt
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ─────────────────────────────────────────────────────────────────────────────
#  Layout definitions
# ─────────────────────────────────────────────────────────────────────────────

# label → (cols, rows)  — images per page = cols * rows
LAYOUTS = {
    "1×1  – Full page":   (1, 1),
    "1×2  – 2 portrait":  (1, 2),
    "1×3  – 3 portrait":  (1, 3),
    "2×2  – 4 per page":  (2, 2),
    "2×3  – 6 per page":  (2, 3),
    "3×3  – 9 per page":  (3, 3),
    "4×4  – 16 per page": (4, 4),
}
DEFAULT_LAYOUT = "2×2  – 4 per page"

A4_W_MM   = 210
A4_H_MM   = 297
MARGIN_MM = 10   # outer page margin
GAP_MM    = 3    # gap between cells


def cell_dims(cols: int, rows: int):
    """Return (cell_w_mm, cell_h_mm) for a given grid on A4."""
    usable_w = A4_W_MM - MARGIN_MM * 2 - GAP_MM * (cols - 1)
    usable_h = A4_H_MM - MARGIN_MM * 2 - GAP_MM * (rows - 1)
    return usable_w / cols, usable_h / rows


THUMB_SIZE = (120, 90)


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def make_thumbnail(path: str):
    if not HAS_PIL:
        return None
    try:
        img = Image.open(path)
        img.thumbnail(THUMB_SIZE, Image.LANCZOS)
        canvas = Image.new("RGB", THUMB_SIZE, (200, 200, 200))
        ox = (THUMB_SIZE[0] - img.width)  // 2
        oy = (THUMB_SIZE[1] - img.height) // 2
        canvas.paste(img, (ox, oy))
        return ImageTk.PhotoImage(canvas)
    except Exception:
        return None


def install_missing():
    import subprocess, sys
    missing = []
    if not HAS_PIL:   missing.append("Pillow")
    if not HAS_DOCX:  missing.append("python-docx")
    if not missing:
        return ""
    subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing,
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return "Installed: " + ", ".join(missing)


# ─────────────────────────────────────────────────────────────────────────────
#  DOCX builder
# ─────────────────────────────────────────────────────────────────────────────

def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc    = cell._tc
    tcPr  = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("start", start),
                      ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"),    str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color="FFFFFF", size=0):
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:val"),   "single")
        node.set(qn("w:sz"),    str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        tblBorders.append(node)
    tblPr.append(tblBorders)


def set_row_height(row, height_mm):
    tr   = row._tr
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    trHeight = OxmlElement("w:trHeight")
    # convert mm → twips (1 mm ≈ 56.69 twips)
    twips = int(height_mm * 56.69)
    trHeight.set(qn("w:val"),  str(twips))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)


def build_docx(image_paths: list, output_path: str,
               cols: int = 2, rows: int = 2,
               progress_cb=None, done_cb=None):
    """
    Build a DOCX with images in a cols×rows A4 grid.
    One page holds cols*rows images.
    """
    if not HAS_DOCX:
        if done_cb:
            done_cb("python-docx not installed. Run: pip install python-docx")
        return

    try:
        from docx.enum.text import WD_BREAK

        doc     = Document()
        section = doc.sections[0]
        section.page_width    = Mm(A4_W_MM)
        section.page_height   = Mm(A4_H_MM)
        section.top_margin    = Mm(MARGIN_MM)
        section.bottom_margin = Mm(MARGIN_MM)
        section.left_margin   = Mm(MARGIN_MM)
        section.right_margin  = Mm(MARGIN_MM)

        style = doc.styles["Normal"]
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after  = Pt(0)

        cell_w, cell_h = cell_dims(cols, rows)
        per_page       = cols * rows
        total          = len(image_paths)
        num_pages      = math.ceil(total / per_page)

        for page_idx in range(num_pages):
            chunk = image_paths[page_idx * per_page : (page_idx + 1) * per_page]

            table = doc.add_table(rows=rows, cols=cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(table)

            # set column widths
            for col_obj in table.columns:
                col_obj.width = Mm(cell_w)

            # set row heights
            for row_obj in table.rows:
                set_row_height(row_obj, cell_h)

            for i, img_path in enumerate(chunk):
                r_idx = i // cols
                c_idx = i % cols
                cell  = table.cell(r_idx, c_idx)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_margins(cell, 0, 0, 0, 0)

                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run  = para.add_run()

                if os.path.isfile(img_path):
                    run.add_picture(img_path, width=Mm(cell_w))

                if progress_cb:
                    progress_cb(int((page_idx * per_page + i + 1) / total * 90))

            # fill empty trailing cells
            for i in range(len(chunk), per_page):
                r_idx = i // cols
                c_idx = i % cols
                cell  = table.cell(r_idx, c_idx)
                set_cell_margins(cell, 0, 0, 0, 0)

            # page break (not after last page)
            if page_idx < num_pages - 1:
                pb_para = doc.add_paragraph()
                pb_para.paragraph_format.space_before = Pt(0)
                pb_para.paragraph_format.space_after  = Pt(0)
                pb_para.add_run().add_break(WD_BREAK.PAGE)

        if progress_cb:
            progress_cb(95)

        doc.save(output_path)

        if progress_cb:
            progress_cb(100)
        if done_cb:
            done_cb(None)

    except Exception as exc:
        if done_cb:
            done_cb(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
#  Draggable image card widget
# ─────────────────────────────────────────────────────────────────────────────

class ImageCard(tk.Frame):
    def __init__(self, master, path: str, index: int,
                 on_drag_end=None, on_delete=None, **kw):
        super().__init__(master, relief="raised", bd=1,
                         bg="#2b2b2b", cursor="hand2", **kw)
        self.path        = path
        self.index       = index
        self.on_drag_end = on_drag_end
        self.on_delete   = on_delete
        self._drag_start_y = 0

        # thumbnail
        self.thumb = make_thumbnail(path)
        if self.thumb:
            img_lbl = tk.Label(self, image=self.thumb, bg="#2b2b2b")
            img_lbl.pack(side="left", padx=4, pady=4)
        else:
            tk.Label(self, text="🖼", font=("Arial", 28), bg="#2b2b2b",
                     fg="#aaaaaa").pack(side="left", padx=8, pady=4)

        # info
        info = tk.Frame(self, bg="#2b2b2b")
        info.pack(side="left", fill="both", expand=True, padx=4)

        self.num_lbl = tk.Label(info, text=f"#{index+1}",
                                font=("Arial", 11, "bold"),
                                bg="#2b2b2b", fg="#f0c040")
        self.num_lbl.pack(anchor="w")

        name = os.path.basename(path)
        if len(name) > 34:
            name = name[:31] + "…"
        tk.Label(info, text=name, font=("Arial", 9),
                 bg="#2b2b2b", fg="#cccccc",
                 wraplength=200, justify="left").pack(anchor="w")
        tk.Label(info, text=path, font=("Arial", 7),
                 bg="#2b2b2b", fg="#888888",
                 wraplength=200, justify="left").pack(anchor="w")

        # right side: delete + drag handle
        right = tk.Frame(self, bg="#2b2b2b")
        right.pack(side="right", padx=6, pady=4)

        tk.Button(right, text="✕", font=("Arial", 10, "bold"),
                  bg="#c0392b", fg="white", relief="flat",
                  padx=6, pady=3, cursor="hand2",
                  activebackground="#e74c3c", activeforeground="white",
                  command=self._on_delete_click).pack(side="top", pady=(0, 4))

        tk.Label(right, text="⠿", font=("Arial", 18),
                 bg="#2b2b2b", fg="#555555").pack(side="top")

        # drag bindings
        drag_targets = [self, info]
        if self.thumb:
            drag_targets.append(img_lbl)
        for w in drag_targets:
            w.bind("<ButtonPress-1>",   self._on_press)
            w.bind("<B1-Motion>",       self._on_motion)
            w.bind("<ButtonRelease-1>", self._on_release)

    def _on_delete_click(self):
        if self.on_delete:
            self.on_delete(self)

    def _on_press(self, e):
        self._drag_start_y = e.y_root
        self.config(relief="sunken")

    def _on_motion(self, e):
        if abs(e.y_root - self._drag_start_y) > 10:
            self.config(relief="flat", bg="#3a3a3a")

    def _on_release(self, e):
        self.config(relief="raised", bg="#2b2b2b")
        if self.on_drag_end:
            self.on_drag_end(self, e.y_root)

    def set_index(self, idx: int):
        self.index = idx
        self.num_lbl.config(text=f"#{idx+1}")


# ─────────────────────────────────────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("A4 Photo Grid → DOCX Generator")
        self.configure(bg="#1e1e1e")
        self.minsize(720, 600)
        self.resizable(True, True)

        self.image_paths: list = []
        self.cards:       list = []

        self._build_ui()
        self._check_deps()

    # ── dependency check ─────────────────────────────────────────────────────
    def _check_deps(self):
        missing = []
        if not HAS_PIL:   missing.append("Pillow")
        if not HAS_DOCX:  missing.append("python-docx")
        if missing:
            self.status_var.set(f"⚠ Missing: {', '.join(missing)} — click Install")
            self.install_btn.config(state="normal")

    # ── UI ───────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── toolbar ──────────────────────────────────────────────────────────
        toolbar = tk.Frame(self, bg="#252525", pady=6)
        toolbar.pack(fill="x", side="top")

        btn = dict(font=("Arial", 10, "bold"), relief="flat",
                   padx=12, pady=6, cursor="hand2")

        tk.Button(toolbar, text="➕  Add Images",
                  bg="#4a90d9", fg="white",
                  command=self._add_images, **btn).pack(side="left", padx=6)

        tk.Button(toolbar, text="🗑  Clear All",
                  bg="#c0392b", fg="white",
                  command=self._clear_all, **btn).pack(side="left", padx=2)

        tk.Button(toolbar, text="📄  Generate DOCX",
                  bg="#27ae60", fg="white",
                  command=self._start_generate, **btn).pack(side="right", padx=6)

        self.install_btn = tk.Button(
            toolbar, text="⬇ Install deps",
            bg="#e67e22", fg="white",
            command=self._install_deps,
            state="disabled", **btn)
        self.install_btn.pack(side="right", padx=2)

        # ── layout dropdown ───────────────────────────────────────────────────
        layout_bar = tk.Frame(self, bg="#1e1e1e", pady=5)
        layout_bar.pack(fill="x", padx=10)

        tk.Label(layout_bar, text="Layout:", bg="#1e1e1e",
                 fg="#aaaaaa", font=("Arial", 10, "bold")).pack(side="left")

        self.layout_var = tk.StringVar(value=DEFAULT_LAYOUT)
        layout_menu = ttk.Combobox(
            layout_bar,
            textvariable=self.layout_var,
            values=list(LAYOUTS.keys()),
            state="readonly",
            font=("Arial", 10),
            width=22)
        layout_menu.pack(side="left", padx=8)
        layout_menu.bind("<<ComboboxSelected>>", lambda e: self._refresh_count())

        self.layout_hint = tk.Label(layout_bar, text="",
                                    bg="#1e1e1e", fg="#888888",
                                    font=("Arial", 9, "italic"))
        self.layout_hint.pack(side="left", padx=4)
        self._update_layout_hint()
        layout_menu.bind("<<ComboboxSelected>>",
                         lambda e: (self._refresh_count(),
                                    self._update_layout_hint()))

        # ── path input bar ────────────────────────────────────────────────────
        path_bar = tk.Frame(self, bg="#1e1e1e", pady=2)
        path_bar.pack(fill="x", padx=10)

        tk.Label(path_bar, text="Paste path:", bg="#1e1e1e",
                 fg="#aaaaaa", font=("Arial", 9)).pack(side="left")

        self.path_var = tk.StringVar()
        path_entry = tk.Entry(path_bar, textvariable=self.path_var,
                              bg="#333333", fg="white",
                              insertbackground="white",
                              relief="flat", font=("Arial", 10))
        path_entry.pack(side="left", fill="x", expand=True, padx=6, ipady=4)
        path_entry.bind("<Return>", lambda e: self._add_from_entry())

        tk.Button(path_bar, text="Add", bg="#4a90d9", fg="white",
                  relief="flat", padx=8, font=("Arial", 9, "bold"),
                  command=self._add_from_entry).pack(side="left")

        # ── separator ─────────────────────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=4)

        # ── count / hint row ──────────────────────────────────────────────────
        info_row = tk.Frame(self, bg="#1e1e1e")
        info_row.pack(fill="x", padx=12)

        self.count_lbl = tk.Label(info_row, text="No images added yet.",
                                  bg="#1e1e1e", fg="#888888",
                                  font=("Arial", 9, "italic"))
        self.count_lbl.pack(side="left")

        tk.Label(info_row, text="Drag cards to reorder",
                 bg="#1e1e1e", fg="#555555",
                 font=("Arial", 8)).pack(side="right")

        # ── scrollable card list ──────────────────────────────────────────────
        list_frame = tk.Frame(self, bg="#1e1e1e")
        list_frame.pack(fill="both", expand=True, padx=10, pady=4)

        canvas = tk.Canvas(list_frame, bg="#1e1e1e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                  command=canvas.yview)
        self.card_container = tk.Frame(canvas, bg="#1e1e1e")
        self.card_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=self.card_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))
        self._canvas = canvas

        # ── status / progress bar ─────────────────────────────────────────────
        bottom = tk.Frame(self, bg="#1a1a1a", pady=4)
        bottom.pack(fill="x", side="bottom")

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(bottom, textvariable=self.status_var,
                 bg="#1a1a1a", fg="#aaaaaa",
                 font=("Arial", 9)).pack(side="left", padx=10)

        self.progress = ttk.Progressbar(bottom, length=200, mode="determinate")
        self.progress.pack(side="right", padx=10)

    # ── layout helpers ────────────────────────────────────────────────────────
    def _current_layout(self):
        return LAYOUTS[self.layout_var.get()]

    def _update_layout_hint(self):
        cols, rows = self._current_layout()
        per = cols * rows
        w, h = cell_dims(cols, rows)
        self.layout_hint.config(
            text=f"{per} image{'s' if per>1 else ''}/page  •  ~{w:.0f}×{h:.0f} mm each")

    # ── card management ───────────────────────────────────────────────────────
    def _add_images(self):
        paths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[("Image files",
                        "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.webp"),
                       ("All files", "*.*")])
        for p in paths:
            self._add_path(p)

    def _add_from_entry(self):
        raw = self.path_var.get().strip().strip('"').strip("'")
        if not raw:
            return
        for p in raw.split(";"):
            p = p.strip()
            if p:
                self._add_path(p)
        self.path_var.set("")

    def _add_path(self, path: str):
        path = os.path.normpath(path)
        if not os.path.isfile(path):
            messagebox.showwarning("Not found", f"File not found:\n{path}")
            return
        if path in self.image_paths:
            messagebox.showinfo("Duplicate", f"Already added:\n{path}")
            return
        self.image_paths.append(path)
        self._add_card(path, len(self.image_paths) - 1)
        self._refresh_count()

    def _add_card(self, path: str, idx: int):
        card = ImageCard(self.card_container, path, idx,
                         on_drag_end=self._on_drag_end,
                         on_delete=self._delete_card)
        card.pack(fill="x", padx=4, pady=3)
        self.cards.append(card)

    def _delete_card(self, card):
        idx = self.cards.index(card)
        self.image_paths.pop(idx)
        self.cards.pop(idx)
        card.destroy()
        for i, c in enumerate(self.cards):
            c.set_index(i)
        self._refresh_count()

    def _clear_all(self):
        if not self.image_paths:
            return
        if messagebox.askyesno("Clear", "Remove all images?"):
            self.image_paths.clear()
            for c in self.cards:
                c.destroy()
            self.cards.clear()
            self._refresh_count()

    def _refresh_count(self):
        n = len(self.image_paths)
        if n == 0:
            self.count_lbl.config(text="No images added yet.")
            return
        cols, rows = self._current_layout()
        per_page = cols * rows
        pages = math.ceil(n / per_page)
        self.count_lbl.config(
            text=f"{n} image{'s' if n>1 else ''} → "
                 f"{pages} A4 page{'s' if pages>1 else ''} "
                 f"({per_page}/page)")

    # ── drag-to-reorder ───────────────────────────────────────────────────────
    def _on_drag_end(self, card, release_y_root: int):
        src_idx    = self.cards.index(card)
        target_idx = src_idx

        for i, c in enumerate(self.cards):
            try:
                cy = c.winfo_rooty()
                ch = c.winfo_height()
                if cy <= release_y_root <= cy + ch:
                    target_idx = i
                    break
            except Exception:
                pass

        if target_idx == src_idx:
            return

        self.image_paths.insert(target_idx, self.image_paths.pop(src_idx))
        self.cards.insert(target_idx, self.cards.pop(src_idx))

        for c in self.cards:
            c.pack_forget()
        for i, c in enumerate(self.cards):
            c.set_index(i)
            c.pack(fill="x", padx=4, pady=3)

    # ── generate DOCX ─────────────────────────────────────────────────────────
    def _start_generate(self):
        if not self.image_paths:
            messagebox.showwarning("No images", "Please add at least one image.")
            return
        if not HAS_DOCX:
            messagebox.showerror("Missing library",
                                 "python-docx is not installed.\n"
                                 "Click 'Install deps' first.")
            return

        out = filedialog.asksaveasfilename(
            title="Save DOCX as…",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")])
        if not out:
            return

        cols, rows = self._current_layout()
        self.progress["value"] = 0
        self.status_var.set(f"Building DOCX ({cols}×{rows})…")
        self.update_idletasks()

        paths_copy = list(self.image_paths)

        def _run():
            build_docx(
                paths_copy, out,
                cols=cols, rows=rows,
                progress_cb=self._update_progress,
                done_cb=self._on_done)

        threading.Thread(target=_run, daemon=True).start()

    def _update_progress(self, pct: int):
        self.after(0, lambda: self._set_progress(pct))

    def _set_progress(self, pct: int):
        self.progress["value"] = pct
        self.status_var.set(f"Processing… {pct}%")
        self.update_idletasks()

    def _on_done(self, err):
        if err:
            self.after(0, lambda: messagebox.showerror("Error", str(err)))
            self.after(0, lambda: self.status_var.set("❌ Failed"))
        else:
            self.after(0, lambda: messagebox.showinfo(
                "Done!", "DOCX generated successfully! ✅"))
            self.after(0, lambda: self.status_var.set("✅ Done!"))

    # ── install deps ──────────────────────────────────────────────────────────
    def _install_deps(self):
        self.status_var.set("Installing dependencies…")
        self.update_idletasks()
        try:
            msg = install_missing()
            messagebox.showinfo("Done",
                                (msg or "All dependencies already installed.") +
                                "\n\nPlease restart the app.")
            self.status_var.set("Restart required.")
        except Exception as e:
            messagebox.showerror("Install failed", str(e))
            self.status_var.set("Install failed.")


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
