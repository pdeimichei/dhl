#!/usr/bin/env python3
"""
DHL Spedizioni — Creatore documenti doganali CSV
Requires Python 3.9+  |  pip install reportlab
"""
import csv
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── ReportLab (PDF) ───────────────────────────────────────────────────────────
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    )
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ── Costanti ──────────────────────────────────────────────────────────────────
VERSION         = "1.0.4"
DELIMITER       = ";"
CONFIG_FILE     = Path.home() / ".dhl_spedizioni.json"
ANA_FILENAME    = "anagrafica_spedizioni.csv"

ANA_HEADERS = ["Rif.", "Tipo", "Descrizione", "Cod. Doganale", "U.M.", "Peso", "Origine"]
DOC_HEADERS = ["Rif.", "Tipo", "Descrizione", "Cod. Doganale",
               "Q.tà", "U.M.", "Prezzo", "Valuta",
               "Peso", "Blank", "Origine"]

ANA_WIDTHS  = [5, 10, 54, 14, 6, 8, 8]
DOC_WIDTHS  = [5, 10, 50, 14, 8, 6, 10, 7, 8, 6, 8]

VALUTE = ["EUR", "USD", "CHF"]

# Palette
C_HEADER_BG = "#2d3436"
C_HEADER_FG = "#dfe6e9"
C_ROW_EVEN  = "#f7f8fa"
C_ROW_ODD   = "#ffffff"
C_SELECTED  = "#d4e6f1"
C_TOOLBAR   = "#2d3436"
C_LOCKED    = "#f0f0f0"
C_UNLOCKED  = "#fffde7"
C_BG        = "#ffffff"
C_SUBTEXT   = "#636e72"

# Button colours → style names
_BTN_COLORS = {
    "#2980b9": "Blue",
    "#27ae60": "Green",
    "#c0392b": "Red",
    "#8e44ad": "Purple",
    "#7f8c8d": "Gray",
    "#d35400": "Orange",
    "#2d3436": "Dark",
    "#ecf0f1": "Light",
}


# ── Utilità ───────────────────────────────────────────────────────────────────
def _darken(hex_color: str, factor: float = 0.82) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#{:02x}{:02x}{:02x}".format(int(r * factor), int(g * factor), int(b * factor))


def _setup_styles():
    """Define all ttk styles. Call once after Tk() is created."""
    s = ttk.Style()
    try:
        s.theme_use("clam")
    except tk.TclError:
        pass

    s.configure("TScrollbar", troughcolor="#ecf0f1", background="#bdc3c7",
                borderwidth=0, relief="flat")
    s.configure("TCombobox", padding=2, relief="flat", borderwidth=1)
    s.map("TCombobox",
          fieldbackground=[("readonly", "white"), ("disabled", C_LOCKED)])

    # Toolbar / dialog buttons
    for color_hex, name in _BTN_COLORS.items():
        fg  = "#555555" if name == "Light" else "white"
        drk = _darken(color_hex)
        sn  = f"{name}.TButton"
        s.configure(sn, background=color_hex, foreground=fg,
                    font=("Arial", 9, "bold"), borderwidth=0, relief="flat",
                    padding=(10, 5), focusthickness=0, focuscolor="none")
        s.map(sn,
              background=[("active", drk), ("pressed", _darken(color_hex, 0.72))],
              foreground=[("active", fg),  ("pressed", fg)],
              relief=[("active", "flat"),  ("pressed", "flat")])

    # Big home-screen buttons
    for color_hex, name in (("#2980b9", "BigBlue"), ("#27ae60", "BigGreen")):
        drk = _darken(color_hex)
        sn  = f"{name}.TButton"
        s.configure(sn, background=color_hex, foreground="white",
                    font=("Arial", 12, "bold"), borderwidth=0, relief="flat",
                    padding=(28, 13), focusthickness=0, focuscolor="none")
        s.map(sn,
              background=[("active", drk), ("pressed", _darken(color_hex, 0.72))],
              foreground=[("active", "white"), ("pressed", "white")],
              relief=[("active", "flat"), ("pressed", "flat")])

    # Inline link-style button (settings)
    s.configure("Link.TButton", background=C_BG, foreground="#2980b9",
                font=("Arial", 8, "underline"), borderwidth=0, relief="flat",
                padding=(4, 2), focusthickness=0, focuscolor="none")
    s.map("Link.TButton",
          background=[("active", C_BG)],
          foreground=[("active", _darken("#2980b9"))])


def _btn(parent, text, color, cmd, side="left", padx=4) -> ttk.Button:
    name = _BTN_COLORS.get(color, "Blue")
    b = ttk.Button(parent, text=text, command=cmd, style=f"{name}.TButton")
    b.pack(side=side, padx=padx, pady=2)
    return b


def open_file(path: str):
    if platform.system() == "Darwin":
        subprocess.call(["open", path])
    elif platform.system() == "Windows":
        os.startfile(path)
    else:
        subprocess.call(["xdg-open", path])


# ── Gestione configurazione ───────────────────────────────────────────────────
class ConfigManager:
    def __init__(self):
        self._data: dict = {}
        self._load()

    def _load(self):
        if CONFIG_FILE.exists():
            try:
                self._data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            except Exception:
                self._data = {}

    def _save(self):
        CONFIG_FILE.write_text(json.dumps(self._data, indent=2), encoding="utf-8")

    def get_folder(self) -> str | None:
        return self._data.get("folder")

    def set_folder(self, path: str):
        self._data["folder"] = path
        self._save()


# ── Finestra Impostazioni ─────────────────────────────────────────────────────
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.title("Impostazioni")
        self.geometry("560x210")
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=C_BG)
        self._config = config
        self._build()
        self.transient(parent)

    def _build(self):
        tk.Label(self, text="Cartella di lavoro", font=("Arial", 10, "bold"),
                 bg=C_BG, fg=C_HEADER_BG, padx=20, pady=(15, 4)).pack(anchor="w")

        row = tk.Frame(self, bg=C_BG, padx=20, pady=4)
        row.pack(fill="x")

        self._folder_var = tk.StringVar(value=self._config.get_folder() or "")
        tk.Entry(row, textvariable=self._folder_var, width=48,
                 state="readonly", font=("Arial", 9),
                 readonlybackground="#f7f8fa", relief="flat",
                 highlightthickness=1, highlightbackground="#dfe6e9").pack(side="left", padx=(0, 6))
        _btn(row, "Sfoglia…", "#2980b9", self._browse)

        tk.Label(self, text="Qui vengono salvati l'anagrafica e i documenti esportati.",
                 font=("Arial", 8), fg=C_SUBTEXT, bg=C_BG,
                 padx=20, pady=(0, 12)).pack(anchor="w")

        sep = tk.Frame(self, bg="#dfe6e9", height=1)
        sep.pack(fill="x", padx=0)

        foot = tk.Frame(self, bg=C_BG, pady=8, padx=14)
        foot.pack(fill="x")
        tk.Label(foot, text=f"DHL Spedizioni  v{VERSION}", font=("Arial", 8),
                 fg="#b2bec3", bg=C_BG).pack(side="left")
        _btn(foot, "Chiudi", "#7f8c8d", self.destroy, side="right", padx=0)

    def _browse(self):
        path = filedialog.askdirectory(
            title="Seleziona cartella di lavoro",
            initialdir=self._config.get_folder() or str(Path.home()))
        if path:
            self._config.set_folder(path)
            self._folder_var.set(path)


# ── Finestra Anagrafica ───────────────────────────────────────────────────────
class AnagraficaWindow(tk.Toplevel):

    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.title("Anagrafica Spedizioni")
        self.geometry("1100x560")
        self.minsize(800, 400)
        self.configure(bg=C_BG)
        self._config = config
        self._rows: list[list[str]] = []
        self._row_widgets: list[dict] = []
        self._selected: int | None = None
        self._build_ui()
        self._load()
        self.transient(parent)

    def _build_ui(self):
        # Toolbar
        bar = tk.Frame(self, bg=C_TOOLBAR, pady=8, padx=12)
        bar.pack(side="top", fill="x")
        tk.Label(bar, text="Anagrafica Prodotti", bg=C_TOOLBAR, fg="white",
                 font=("Arial", 12, "bold")).pack(side="left", padx=(0, 16))
        _btn(bar, "＋  Aggiungi prodotto", "#2980b9", self._add_row)
        _btn(bar, "✕  Elimina prodotto",  "#c0392b", self._del_row)
        _btn(bar, "←  Chiudi",            "#7f8c8d", self.destroy, side="right")

        # Status bar
        self._status = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status, anchor="w",
                 bg="#f0f2f5", fg=C_SUBTEXT, font=("Arial", 9),
                 padx=12, pady=4).pack(side="bottom", fill="x")

        # Canvas + scrollbars
        outer = tk.Frame(self, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=8, pady=(6, 4))

        self._canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical",   command=self._canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=C_BG)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _e: self._canvas.configure(
                             scrollregion=self._canvas.bbox("all")))
        for w in (self._canvas, self._inner):
            w.bind("<MouseWheel>", self._on_scroll)
            w.bind("<Button-4>",  lambda _e: self._canvas.yview_scroll(-1, "units"))
            w.bind("<Button-5>",  lambda _e: self._canvas.yview_scroll( 1, "units"))

    def _on_scroll(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── CSV I/O ───────────────────────────────────────────────────────────────
    def _csv_path(self) -> str | None:
        folder = self._config.get_folder()
        return os.path.join(folder, ANA_FILENAME) if folder else None

    def _load(self):
        path = self._csv_path()
        if not path:
            messagebox.showwarning("Cartella non configurata",
                                   "Configura prima la cartella di lavoro nelle Impostazioni.",
                                   parent=self)
            return
        if not os.path.exists(path):
            with open(path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f, delimiter=DELIMITER).writerow(ANA_HEADERS)
            self._rows = []
        else:
            with open(path, newline="", encoding="utf-8-sig") as f:
                all_rows = list(csv.reader(f, delimiter=DELIMITER))
            # Skip header row by checking first cell only (robust across schema changes)
            data = all_rows[1:] if all_rows and all_rows[0] and all_rows[0][0] == ANA_HEADERS[0] else all_rows
            n = len(ANA_HEADERS)
            self._rows = [r[:n] + [""] * max(0, n - len(r)) for r in data if any(c.strip() for c in r)]
        self._rebuild_table()
        self._update_status()

    def _save(self):
        path = self._csv_path()
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=DELIMITER)
            w.writerow(ANA_HEADERS)
            for row in self._rows:
                w.writerow(row)

    # ── Tabella ───────────────────────────────────────────────────────────────
    def _rebuild_table(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._row_widgets = []
        self._selected = None

        # Header row
        hf = tk.Frame(self._inner, bg=C_HEADER_BG)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="#", width=3, bg=C_HEADER_BG, fg=C_HEADER_FG,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, padx=(6, 2), pady=5)
        for j, (h, w) in enumerate(zip(ANA_HEADERS, ANA_WIDTHS)):
            tk.Label(hf, text=h, width=w, bg=C_HEADER_BG, fg=C_HEADER_FG,
                     font=("Arial", 9, "bold"), anchor="center"
                     ).grid(row=0, column=j + 1, padx=2, pady=5)

        for i, row_data in enumerate(self._rows):
            self._build_ana_row(i, row_data)

    def _build_ana_row(self, i: int, row_data: list[str]):
        bg = C_ROW_EVEN if i % 2 == 0 else C_ROW_ODD
        frame = tk.Frame(self._inner, bg=bg)
        frame.grid(row=i + 1, column=0, sticky="ew")

        num = tk.Label(frame, text=str(i + 1), width=3,
                       bg=bg, fg="#aaa", font=("Arial", 9), cursor="hand2")
        num.grid(row=0, column=0, padx=(6, 2))
        num.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))

        vars_: list[tk.StringVar] = []
        entries: list[tk.Widget]  = []

        for j, (val, w) in enumerate(zip(row_data, ANA_WIDTHS)):
            var = tk.StringVar(value=val)
            vars_.append(var)
            e = tk.Entry(frame, textvariable=var, width=w,
                         bg=C_LOCKED, relief="flat", font=("Arial", 10),
                         state="readonly", cursor="arrow",
                         readonlybackground=C_LOCKED,
                         highlightthickness=1, highlightbackground="#e8eaed")
            e.grid(row=0, column=j + 1, padx=2, pady=3, sticky="ew")
            e.bind("<Button-1>", lambda event, idx=i, col=j: self._try_unlock(idx, col))
            entries.append(e)

        self._row_widgets.append({"frame": frame, "vars": vars_, "entries": entries, "bg": bg})

    # ── Blocco/sblocco ────────────────────────────────────────────────────────
    def _try_unlock(self, row_idx: int, col_idx: int):
        entry: tk.Entry = self._row_widgets[row_idx]["entries"][col_idx]
        if entry.cget("state") == "normal":
            return
        if not messagebox.askyesno(
                "Modifica campo",
                "Sei sicuro di voler modificare questo campo?\n"
                "La modifica sarà salvata immediatamente.",
                parent=self):
            return
        entry.configure(state="normal", bg=C_UNLOCKED,
                        readonlybackground=C_UNLOCKED, cursor="xterm",
                        highlightbackground="#f9ca24")
        entry.focus_set()
        entry.bind("<Return>",   lambda _e, r=row_idx, c=col_idx: self._lock_and_save(r, c))
        entry.bind("<FocusOut>", lambda _e, r=row_idx, c=col_idx: self._lock_and_save(r, c))

    def _lock_and_save(self, row_idx: int, col_idx: int):
        winfo = self._row_widgets[row_idx]
        entry: tk.Entry = winfo["entries"][col_idx]
        if entry.cget("state") == "readonly":
            return
        for j, var in enumerate(winfo["vars"]):
            self._rows[row_idx][j] = var.get()
        entry.configure(state="readonly", bg=C_LOCKED,
                        readonlybackground=C_LOCKED, cursor="arrow",
                        highlightbackground="#e8eaed")
        self._save()
        self._update_status("Salvato")

    # ── Selezione riga ────────────────────────────────────────────────────────
    def _select_row(self, idx: int):
        if self._selected is not None and self._selected < len(self._row_widgets):
            self._set_row_bg(self._selected, self._row_widgets[self._selected]["bg"])
        self._selected = idx
        self._set_row_bg(idx, C_SELECTED)
        self._update_status()

    def _set_row_bg(self, idx: int, color: str):
        if idx is None or idx >= len(self._row_widgets):
            return
        frame = self._row_widgets[idx]["frame"]
        frame.configure(bg=color)
        for child in frame.winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass

    # ── Aggiungi / Elimina ────────────────────────────────────────────────────
    def _add_row(self):
        new_row = ["1", "INV_ITEM", "", "", "PCS", "", "IT"]  # Rif.|Tipo|Desc|Cod|UM|Peso|Origine
        self._rows.append(new_row)
        i = len(self._rows) - 1
        self._build_ana_row(i, new_row)
        winfo = self._row_widgets[i]
        for j, entry in enumerate(winfo["entries"]):
            entry.configure(state="normal", bg=C_UNLOCKED,
                            readonlybackground=C_UNLOCKED, cursor="xterm")
            entry.bind("<Return>",   lambda _e, r=i, c=j: self._lock_and_save(r, c))
            entry.bind("<FocusOut>", lambda _e, r=i, c=j: self._lock_and_save(r, c))
        winfo["entries"][2].focus_set()
        self._canvas.after(60, lambda: self._canvas.yview_moveto(1.0))
        self._update_status()

    def _del_row(self):
        if self._selected is None:
            messagebox.showinfo("Nessuna selezione",
                                "Clicca sul numero di riga per selezionarla.",
                                parent=self)
            return
        if not messagebox.askyesno("Elimina prodotto",
                                   f"Eliminare il prodotto alla riga {self._selected + 1}?\n"
                                   "L'operazione non è reversibile.",
                                   icon="warning", parent=self):
            return
        self._rows.pop(self._selected)
        self._save()
        self._rebuild_table()
        self._update_status("Prodotto eliminato")

    def _update_status(self, extra: str = ""):
        parts = [f"{len(self._rows)} prodotti"]
        if self._selected is not None:
            parts.append(f"Riga {self._selected + 1} selezionata")
        if extra:
            parts.append(extra)
        self._status.set("  ·  ".join(parts))

    def get_products(self) -> list[list[str]]:
        return [list(r) for r in self._rows]


# ── Finestra Crea Documento ───────────────────────────────────────────────────
class DocumentoWindow(tk.Toplevel):

    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.title("Crea Documento")
        self.geometry("1200x640")
        self.minsize(900, 480)
        self.configure(bg=C_BG)
        self._config = config
        self._products: list[list[str]] = []
        self._doc_rows: list[dict] = []
        self._selected_row: int | None = None
        self._created_at: datetime | None = None
        self._last_saved: datetime | None = None
        self._autosave_pending = False
        self._load_products()
        self._build_ui()
        self.transient(parent)

    def _load_products(self):
        folder = self._config.get_folder()
        if not folder:
            return
        path = os.path.join(folder, ANA_FILENAME)
        if not os.path.exists(path):
            return
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter=DELIMITER))
        if rows and rows[0] and rows[0][0] == ANA_HEADERS[0]:
            rows = rows[1:]
        n = len(ANA_HEADERS)
        self._products = [r[:n] + [""] * max(0, n - len(r)) for r in rows if any(c.strip() for c in r)]

    def _build_ui(self):
        # Toolbar
        bar = tk.Frame(self, bg=C_TOOLBAR, pady=8, padx=12)
        bar.pack(side="top", fill="x")
        tk.Label(bar, text="Crea Documento", bg=C_TOOLBAR, fg="white",
                 font=("Arial", 12, "bold")).pack(side="left", padx=(0, 16))
        _btn(bar, "＋  Aggiungi riga",  "#2980b9", self._add_row)
        _btn(bar, "✕  Elimina riga",   "#c0392b", self._del_row)
        _btn(bar, "💾  Salva CSV",      "#27ae60", self._save_csv)
        _btn(bar, "🖨  Stampa PDF",     "#8e44ad", self._export_pdf)
        _btn(bar, "←  Chiudi",         "#7f8c8d", self.destroy, side="right")

        # Header: nome file
        hdr = tk.Frame(self, bg="#f0f2f5", padx=14, pady=10)
        hdr.pack(side="top", fill="x")
        tk.Label(hdr, text="Nome file:", bg="#f0f2f5", fg=C_HEADER_BG,
                 font=("Arial", 10, "bold")).pack(side="left")
        self._filename_var = tk.StringVar()
        self._filename_var.trace_add("write", self._on_filename_change)
        tk.Entry(hdr, textvariable=self._filename_var, width=24,
                 font=("Arial", 11, "bold"), relief="flat",
                 highlightthickness=1, highlightbackground="#b2bec3",
                 bg="white").pack(side="left", padx=(8, 2), ipady=3)
        tk.Label(hdr, text=".csv", bg="#f0f2f5", fg=C_SUBTEXT,
                 font=("Arial", 10)).pack(side="left")
        tk.Label(hdr, text="  es. 2025-600", bg="#f0f2f5", fg="#b2bec3",
                 font=("Arial", 9)).pack(side="left")
        self._hdr_status = tk.StringVar(value="")
        tk.Label(hdr, textvariable=self._hdr_status, bg="#f0f2f5",
                 fg="#27ae60", font=("Arial", 9, "bold")).pack(side="right", padx=8)

        # Status bar
        self._status = tk.StringVar(value="Aggiungi righe al documento")
        tk.Label(self, textvariable=self._status, anchor="w",
                 bg="#f0f2f5", fg=C_SUBTEXT, font=("Arial", 9),
                 padx=12, pady=4).pack(side="bottom", fill="x")

        # Scrollable table
        outer = tk.Frame(self, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=8, pady=(4, 4))

        self._canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical",   command=self._canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=C_BG)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _e: self._canvas.configure(
                             scrollregion=self._canvas.bbox("all")))
        for w in (self._canvas, self._inner):
            w.bind("<MouseWheel>", self._on_scroll)
            w.bind("<Button-4>",  lambda _e: self._canvas.yview_scroll(-1, "units"))
            w.bind("<Button-5>",  lambda _e: self._canvas.yview_scroll( 1, "units"))

        self._build_doc_header()

    def _on_scroll(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_doc_header(self):
        hf = tk.Frame(self._inner, bg=C_HEADER_BG)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="#", width=3, bg=C_HEADER_BG, fg=C_HEADER_FG,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, padx=(6, 2), pady=5)
        for j, (h, w) in enumerate(zip(DOC_HEADERS, DOC_WIDTHS)):
            label = "" if h == "Blank" else h   # blank column has no visible header
            tk.Label(hf, text=label, width=w, bg=C_HEADER_BG, fg=C_HEADER_FG,
                     font=("Arial", 9, "bold"), anchor="center"
                     ).grid(row=0, column=j + 1, padx=2, pady=5)

    # ── Righe documento ───────────────────────────────────────────────────────
    def _add_row(self):
        if not self._config.get_folder():
            messagebox.showwarning("Cartella non configurata",
                                   "Configura la cartella di lavoro nelle Impostazioni.",
                                   parent=self)
            return
        i = len(self._doc_rows)
        bg = C_ROW_EVEN if i % 2 == 0 else C_ROW_ODD

        frame = tk.Frame(self._inner, bg=bg)
        frame.grid(row=i + 1, column=0, sticky="ew")

        num = tk.Label(frame, text=str(i + 1), width=3,
                       bg=bg, fg="#aaa", font=("Arial", 9), cursor="hand2")
        num.grid(row=0, column=0, padx=(6, 2))
        num.bind("<Button-1>", lambda _e, idx=i: self._select_doc_row(idx))

        vars_ = {h: tk.StringVar() for h in DOC_HEADERS}
        vars_["Rif."].set("1")
        vars_["Tipo"].set("INV_ITEM")
        vars_["Valuta"].set("EUR")
        vars_["Blank"].set("")  # always empty per customer upload requirement

        widgets = {}

        def make_trace(v):
            v.trace_add("write", self._schedule_autosave)
        for h in DOC_HEADERS:
            make_trace(vars_[h])

        col_idx = 1

        # Lockable (pre-filled from anagrafica) — editable with confirmation,
        # changes apply only to this document, not back to the anagrafica.
        def lockable_entry(var, width, col):
            e = tk.Entry(frame, textvariable=var, width=width,
                         font=("Arial", 10), bg=C_LOCKED, state="readonly",
                         readonlybackground=C_LOCKED, relief="flat",
                         highlightthickness=1, highlightbackground="#e8eaed")
            e.grid(row=0, column=col, padx=2, pady=3)
            def on_click(_evt, w=e):
                self._select_doc_row(i)
                frame.after(10, lambda: self._try_unlock_doc_field(w))
            e.bind("<Button-1>", on_click)
            return e

        # Freely editable — user input fields (no confirmation needed)
        def rw_entry(var, width, col):
            e = tk.Entry(frame, textvariable=var, width=width,
                         font=("Arial", 10), bg="white", relief="flat",
                         highlightthickness=1, highlightbackground="#b2bec3")
            e.grid(row=0, column=col, padx=2, pady=3)
            e.bind("<Button-1>", lambda _e, idx=i: self._select_doc_row(idx))
            return e

        # DOC_HEADERS order: Rif. | Tipo | Descrizione | Cod. Doganale |
        #                    Q.tà | Peso | U.M. | Prezzo | Valuta |
        #                    Altro Prezzo | (vuoto) | Origine
        widgets["Rif."]          = lockable_entry(vars_["Rif."],          DOC_WIDTHS[0],  col_idx);  col_idx += 1
        widgets["Tipo"]          = lockable_entry(vars_["Tipo"],           DOC_WIDTHS[1],  col_idx);  col_idx += 1

        # Descrizione — combobox (always freely editable)
        desc_values = [p[2] for p in self._products]
        cb_desc = ttk.Combobox(frame, textvariable=vars_["Descrizione"],
                               values=desc_values, width=DOC_WIDTHS[2] - 2,
                               font=("Arial", 10))
        cb_desc.grid(row=0, column=col_idx, padx=2, pady=3);  col_idx += 1
        cb_desc.bind("<<ComboboxSelected>>", lambda _e, idx=i: self._on_desc_selected(idx))
        cb_desc.bind("<Button-1>", lambda _e, idx=i: self._select_doc_row(idx))
        widgets["Descrizione"] = cb_desc

        widgets["Cod. Doganale"] = lockable_entry(vars_["Cod. Doganale"], DOC_WIDTHS[3],  col_idx);  col_idx += 1
        widgets["Q.tà"]          = rw_entry(vars_["Q.tà"],                DOC_WIDTHS[4],  col_idx);  col_idx += 1
        widgets["U.M."]          = lockable_entry(vars_["U.M."],           DOC_WIDTHS[5],  col_idx);  col_idx += 1
        widgets["Prezzo"]        = rw_entry(vars_["Prezzo"],               DOC_WIDTHS[6],  col_idx);  col_idx += 1

        # Valuta — combobox (always freely editable)
        cb_val = ttk.Combobox(frame, textvariable=vars_["Valuta"],
                              values=VALUTE, width=DOC_WIDTHS[8] - 2,
                              font=("Arial", 10), state="readonly")
        cb_val.grid(row=0, column=col_idx, padx=2, pady=3);  col_idx += 1
        cb_val.bind("<Button-1>", lambda _e, idx=i: self._select_doc_row(idx))
        widgets["Valuta"] = cb_val

        # Blank column: always empty, not editable, no label
        _blank_e = tk.Entry(frame, textvariable=vars_["Blank"], width=DOC_WIDTHS[9],
                            font=("Arial", 10), bg=C_LOCKED, state="readonly",
                            readonlybackground=C_LOCKED, relief="flat",
                            highlightthickness=0)
        _blank_e.grid(row=0, column=col_idx, padx=2, pady=3)
        _blank_e.bind("<Button-1>", lambda _e, idx=i: self._select_doc_row(idx))
        widgets["Peso"]          = lockable_entry(vars_["Peso"],           DOC_WIDTHS[8],  col_idx);  col_idx += 1
        widgets["Blank"] = _blank_e;  col_idx += 1
        widgets["Origine"]       = lockable_entry(vars_["Origine"],        DOC_WIDTHS[10], col_idx)

        self._doc_rows.append({"frame": frame, "vars": vars_, "widgets": widgets, "bg": bg})
        self._canvas.after(60, lambda: self._canvas.yview_moveto(1.0))
        widgets["Descrizione"].focus_set()
        self._update_status()

    def _on_desc_selected(self, row_idx: int):
        row = self._doc_rows[row_idx]
        desc = row["vars"]["Descrizione"].get()
        product = next((p for p in self._products if p[2] == desc), None)
        if product is None:
            return
        # ANA_HEADERS: Rif. | Tipo | Descrizione | Cod. Doganale | U.M. | Peso | Origine
        rif, tipo, _, cod, um, peso, origine = (product + [""] * 7)[:7]
        for key, val in [("Rif.", rif), ("Tipo", tipo),
                         ("Cod. Doganale", cod), ("U.M.", um),
                         ("Peso", peso), ("Origine", origine)]:
            row["vars"][key].set(val)

    # ── Blocco/sblocco campi documento ───────────────────────────────────────
    def _try_unlock_doc_field(self, entry: tk.Entry):
        """Chiede conferma prima di rendere editabile un campo pre-compilato.
        Le modifiche rimangono solo nel documento corrente, non nell'anagrafica."""
        if entry.cget("state") == "normal":
            return  # già in modifica
        if not messagebox.askyesno(
                "Modifica campo",
                "Sei sicuro di voler modificare questo campo?\n\n"
                "La modifica sarà applicata solo a questo documento.\n"
                "L'anagrafica prodotti non verrà modificata.",
                parent=self):
            return
        entry.configure(state="normal", bg=C_UNLOCKED,
                        readonlybackground=C_UNLOCKED, cursor="xterm",
                        highlightbackground="#f9ca24")
        entry.focus_set()
        entry.bind("<Return>",   lambda _e, w=entry: self._lock_doc_field(w))
        entry.bind("<FocusOut>", lambda _e, w=entry: self._lock_doc_field(w))

    def _lock_doc_field(self, entry: tk.Entry):
        if entry.cget("state") == "readonly":
            return
        entry.configure(state="readonly", bg=C_LOCKED,
                        readonlybackground=C_LOCKED, cursor="arrow",
                        highlightbackground="#e8eaed")
        self._schedule_autosave()

    # ── Selezione riga ────────────────────────────────────────────────────────
    def _select_doc_row(self, idx: int):
        if self._selected_row is not None and self._selected_row < len(self._doc_rows):
            old = self._doc_rows[self._selected_row]
            self._set_doc_row_bg(self._selected_row, old["bg"])
        self._selected_row = idx
        self._set_doc_row_bg(idx, C_SELECTED)
        self._update_status()

    def _set_doc_row_bg(self, idx: int, color: str):
        if idx is None or idx >= len(self._doc_rows):
            return
        frame = self._doc_rows[idx]["frame"]
        frame.configure(bg=color)
        for child in frame.winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass

    # ── Elimina riga ──────────────────────────────────────────────────────────
    def _del_row(self):
        if self._selected_row is None:
            messagebox.showinfo("Nessuna selezione",
                                "Clicca sul numero di riga per selezionarla.",
                                parent=self)
            return
        if not messagebox.askyesno("Elimina riga",
                                   f"Eliminare la riga {self._selected_row + 1}?",
                                   icon="warning", parent=self):
            return
        self._doc_rows[self._selected_row]["frame"].destroy()
        self._doc_rows.pop(self._selected_row)
        self._selected_row = None
        for i, r in enumerate(self._doc_rows):
            r["bg"] = C_ROW_EVEN if i % 2 == 0 else C_ROW_ODD
            r["frame"].grid(row=i + 1, column=0, sticky="ew")
            for child in r["frame"].winfo_children():
                if isinstance(child, tk.Label):
                    child.configure(text=str(i + 1))
                    break
        self._schedule_autosave()
        self._update_status()

    # ── Nome file & salvataggio ───────────────────────────────────────────────
    def _on_filename_change(self, *_):
        self._schedule_autosave()

    def _get_output_path(self) -> str | None:
        folder = self._config.get_folder()
        name   = self._filename_var.get().strip()
        if not folder or not name:
            return None
        return os.path.join(folder, name if name.endswith(".csv") else name + ".csv")

    def _schedule_autosave(self, *_):
        if self._autosave_pending:
            return
        self._autosave_pending = True
        self.after(600, self._autosave)

    def _autosave(self):
        self._autosave_pending = False
        path = self._get_output_path()
        if not path:
            return
        self._write_csv(path)
        now = datetime.now()
        if self._created_at is None:
            self._created_at = now
        self._last_saved = now
        self._hdr_status.set(f"✓ Salvato automaticamente — {now.strftime('%H:%M:%S')}")

    def _save_csv(self):
        path = self._get_output_path()
        if not path:
            messagebox.showwarning("Nome file mancante",
                                   "Inserisci un nome file prima di salvare.",
                                   parent=self)
            return
        self._write_csv(path)
        now = datetime.now()
        if self._created_at is None:
            self._created_at = now
        self._last_saved = now
        self._hdr_status.set(f"✓ Salvato — {now.strftime('%H:%M:%S')}")
        messagebox.showinfo("Salvato", f"Documento salvato in:\n{path}", parent=self)

    def _write_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            csv.writer(f, delimiter=DELIMITER).writerows(
                [[r["vars"][h].get() for h in DOC_HEADERS] for r in self._doc_rows]
            )

    # ── Esporta PDF ───────────────────────────────────────────────────────────
    def _export_pdf(self):
        if not REPORTLAB_OK:
            messagebox.showerror(
                "ReportLab non disponibile",
                "La libreria 'reportlab' non è installata.\nEsegui:  pip install reportlab",
                parent=self)
            return

        filename    = self._filename_var.get().strip() or "documento"
        folder      = self._config.get_folder() or tempfile.gettempdir()
        pdf_name    = filename if filename.endswith(".pdf") else filename + ".pdf"
        pdf_path    = os.path.join(folder, pdf_name)
        created_str = (self._created_at or datetime.now()).strftime("%d/%m/%Y %H:%M")
        modified_str= (self._last_saved  or datetime.now()).strftime("%d/%m/%Y %H:%M")

        doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4),
                                leftMargin=1.5*cm, rightMargin=1.5*cm,
                                topMargin=1.5*cm,  bottomMargin=1.5*cm)
        styles   = getSampleStyleSheet()
        t_style  = ParagraphStyle("T", parent=styles["Heading1"], fontSize=16,
                                  textColor=colors.HexColor(C_HEADER_BG), spaceAfter=4)
        s_style  = ParagraphStyle("S", parent=styles["Normal"], fontSize=9,
                                  textColor=colors.HexColor(C_SUBTEXT), spaceAfter=12)
        c_style  = ParagraphStyle("C",  parent=styles["Normal"], fontSize=8, leading=10)
        ch_style = ParagraphStyle("CH", parent=styles["Normal"], fontSize=8, leading=10,
                                  textColor=colors.white)

        story = [
            Paragraph(filename, t_style),
            Paragraph(f"Data creazione: <b>{created_str}</b> &nbsp;&nbsp; "
                      f"Ultima modifica: <b>{modified_str}</b>", s_style),
        ]

        vis_headers = [h for h in DOC_HEADERS if h != ""]
        vis_indices  = [i for i, h in enumerate(DOC_HEADERS) if h != ""]

        table_data = [[Paragraph(f"<b>{h}</b>", ch_style) for h in vis_headers]]
        for row in self._doc_rows:
            table_data.append([
                Paragraph(row["vars"][DOC_HEADERS[i]].get(), c_style)
                for i in vis_indices
            ])

        page_w = landscape(A4)[0] - 3 * cm
        weights = [1, 2, 8, 2.5, 2, 1.5, 2, 1.5, 2, 1.5]
        col_widths = [page_w * (w / sum(weights)) for w in weights]

        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",     (0, 0), (-1, 0),  colors.HexColor(C_HEADER_BG)),
            ("TEXTCOLOR",      (0, 0), (-1, 0),  colors.whitesmoke),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor(C_ROW_EVEN), colors.white]),
            ("GRID",           (0, 0), (-1, -1), 0.4, colors.HexColor("#dfe6e9")),
            ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",     (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
            ("LEFTPADDING",    (0, 0), (-1, -1), 4),
        ]))

        story.append(tbl)
        doc.build(story)
        open_file(pdf_path)
        self._update_status(f"PDF salvato: {pdf_name}")

    # ── Status ────────────────────────────────────────────────────────────────
    def _update_status(self, extra: str = ""):
        parts = [f"{len(self._doc_rows)} righe"]
        if self._selected_row is not None:
            parts.append(f"Riga {self._selected_row + 1} selezionata")
        path = self._get_output_path()
        if path:
            parts.append(path)
        if extra:
            parts.append(extra)
        self._status.set("  ·  ".join(parts))


# ── Schermata principale ──────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"DHL Spedizioni  v{VERSION}")
        self.geometry("460x340")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        _setup_styles()
        self._config = ConfigManager()
        self._build_ui()
        self._ensure_folder()

    def _build_ui(self):
        # Top accent band
        tk.Frame(self, bg=C_TOOLBAR, height=5).pack(fill="x")

        # Main content
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=32, pady=(24, 8))

        # Title
        tk.Label(body, text="DHL Spedizioni", font=("Arial", 22, "bold"),
                 fg=C_HEADER_BG, bg=C_BG).pack(anchor="w")
        tk.Label(body, text="Gestione documenti doganali", font=("Arial", 10),
                 fg=C_SUBTEXT, bg=C_BG).pack(anchor="w", pady=(2, 20))

        # Buttons
        ttk.Button(body, text="📋   Visualizza Anagrafica",
                   command=self._open_anagrafica,
                   style="BigBlue.TButton").pack(fill="x", pady=(0, 8))
        ttk.Button(body, text="📄   Crea Documento",
                   command=self._open_documento,
                   style="BigGreen.TButton").pack(fill="x")

        # Bottom bar
        sep = tk.Frame(self, bg="#e8eaed", height=1)
        sep.pack(fill="x", pady=(16, 0))

        bottom = tk.Frame(self, bg="#f7f8fa", pady=6, padx=14)
        bottom.pack(side="bottom", fill="x")

        self._folder_label = tk.StringVar()
        self._refresh_folder_label()
        tk.Label(bottom, textvariable=self._folder_label,
                 bg="#f7f8fa", fg="#636e72", font=("Arial", 10),
                 anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="⚙  Impostazioni",
                   command=self._open_settings,
                   style="Link.TButton").pack(side="right")

    def _refresh_folder_label(self):
        folder = self._config.get_folder()
        self._folder_label.set(f"📁  {folder}" if folder else "Cartella non configurata")

    def _ensure_folder(self):
        if not self._config.get_folder():
            messagebox.showinfo(
                "Benvenuto in DHL Spedizioni",
                "Seleziona la cartella di lavoro dove verranno salvati\n"
                "l'anagrafica prodotti e i documenti esportati.")
            path = filedialog.askdirectory(title="Seleziona cartella di lavoro",
                                           initialdir=str(Path.home()))
            if path:
                self._config.set_folder(path)
                self._refresh_folder_label()
                self._copy_default_anagrafica(path)

    def _copy_default_anagrafica(self, folder: str):
        dest = os.path.join(folder, ANA_FILENAME)
        if os.path.exists(dest):
            return
        candidates = [
            os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), ANA_FILENAME),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), ANA_FILENAME),
        ]
        if getattr(sys, "frozen", False):
            candidates.insert(0, os.path.join(sys._MEIPASS, ANA_FILENAME))
        for src in candidates:
            if os.path.exists(src):
                shutil.copy2(src, dest)
                return

    def _open_anagrafica(self):
        if not self._config.get_folder():
            self._open_settings()
            return
        AnagraficaWindow(self, self._config)

    def _open_documento(self):
        if not self._config.get_folder():
            self._open_settings()
            return
        DocumentoWindow(self, self._config)

    def _open_settings(self):
        SettingsDialog(self, self._config)
        self._refresh_folder_label()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
