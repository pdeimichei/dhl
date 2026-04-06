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
DELIMITER       = ";"
CONFIG_FILE     = Path.home() / ".dhl_spedizioni.json"
ANA_FILENAME    = "anagrafica_spedizioni.csv"

# Colonne dell'anagrafica (con intestazioni)
ANA_HEADERS = ["Rif.", "Tipo", "Descrizione", "Cod. Doganale", "U.M.", "Altro Prezzo", "Origine"]

# Colonne del documento di spedizione
DOC_HEADERS = ["Rif.", "Tipo", "Descrizione", "Cod. Doganale",
               "Q.tà / Peso", "U.M.", "Prezzo", "Valuta",
               "Altro Prezzo", "", "Origine"]

# Larghezze colonne anagrafica (caratteri)
ANA_WIDTHS  = [5, 10, 54, 14, 6, 12, 8]
# Larghezze colonne documento
DOC_WIDTHS  = [5, 10, 50, 14, 10, 6, 10, 7, 12, 4, 8]

VALUTE      = ["EUR", "USD", "CHF"]

# Colori
C_HEADER_BG = "#2d3436"
C_HEADER_FG = "#dfe6e9"
C_ROW_EVEN  = "#f5f6fa"
C_ROW_ODD   = "#ffffff"
C_SELECTED  = "#d4e6f1"
C_TOOLBAR   = "#2d3436"
C_LOCKED    = "#f0f0f0"   # sfondo celle bloccate
C_UNLOCKED  = "#fffde7"   # sfondo cella in modifica


# ── Utilità ───────────────────────────────────────────────────────────────────
def _darken(hex_color: str, factor: float = 0.85) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#{:02x}{:02x}{:02x}".format(int(r * factor), int(g * factor), int(b * factor))


def _btn(parent, text, color, cmd, side="left", padx=4):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=color, fg="white", relief="flat",
                  font=("Arial", 9, "bold"), padx=10, pady=4,
                  cursor="hand2", activeforeground="white",
                  activebackground=color, bd=0)
    b.pack(side=side, padx=padx)
    darker = _darken(color)
    b.bind("<Enter>", lambda _e: b.configure(bg=darker))
    b.bind("<Leave>", lambda _e: b.configure(bg=color))
    return b


def open_file(path: str):
    """Apre un file col visualizzatore predefinito del sistema."""
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
        self.resizable(False, False)
        self.grab_set()
        self._config = config
        self._build()
        self.transient(parent)

    def _build(self):
        tk.Label(self, text="Cartella di lavoro", font=("Arial", 10, "bold"),
                 padx=20, pady=(15, 4)).pack(anchor="w")

        row = tk.Frame(self, padx=20, pady=4)
        row.pack(fill="x")

        self._folder_var = tk.StringVar(value=self._config.get_folder() or "")
        tk.Entry(row, textvariable=self._folder_var, width=48,
                 state="readonly", font=("Arial", 9)).pack(side="left", padx=(0, 6))
        _btn(row, "Sfoglia…", "#2980b9", self._browse)

        tk.Label(self, text="Qui viene salvata l'anagrafica e i documenti esportati.",
                 font=("Arial", 8), fg="#888", padx=20, pady=(0, 12)).pack(anchor="w")

        _btn(self, "Chiudi", "#7f8c8d", self.destroy, side="right", padx=20)

    def _browse(self):
        path = filedialog.askdirectory(title="Seleziona cartella di lavoro",
                                       initialdir=self._config.get_folder() or str(Path.home()))
        if path:
            self._config.set_folder(path)
            self._folder_var.set(path)


# ── Finestra Anagrafica ───────────────────────────────────────────────────────
class AnagraficaWindow(tk.Toplevel):
    """
    Mostra e permette di modificare anagrafica_spedizioni.csv con meccanismo di blocco
    per campo: ogni cella è readonly finché l'utente non conferma la modifica.
    Salvataggio immediato su disco a ogni modifica confermata.
    """

    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.title("Anagrafica Spedizioni")
        self.geometry("1100x560")
        self.minsize(800, 400)
        self._config = config
        self._rows: list[list[str]] = []
        self._row_widgets: list[dict] = []   # [{vars, entries, frame}, …]
        self._selected: int | None = None
        self._apply_style()
        self._build_ui()
        self._load()
        self.transient(parent)

    def _apply_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TScrollbar", troughcolor="#ecf0f1", background="#bdc3c7")

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        bar = tk.Frame(self, bg=C_TOOLBAR, pady=7, padx=10)
        bar.pack(side="top", fill="x")

        tk.Label(bar, text="Anagrafica Prodotti", bg=C_TOOLBAR, fg="white",
                 font=("Arial", 12, "bold")).pack(side="left", padx=(0, 20))
        _btn(bar, "＋  Aggiungi prodotto", "#2980b9", self._add_row)
        _btn(bar, "✕  Elimina prodotto",  "#c0392b", self._del_row)
        _btn(bar, "←  Chiudi",            "#7f8c8d", self.destroy, side="right")

        self._status = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status, anchor="w",
                 bg="#ecf0f1", fg="#555", font=("Arial", 9),
                 padx=10, pady=3).pack(side="bottom", fill="x")

        # Canvas + scrollbar
        outer = tk.Frame(self, bg="white")
        outer.pack(fill="both", expand=True, padx=8, pady=(6, 4))

        self._canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical",   command=self._canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg="white")
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
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
        if not folder:
            return None
        return os.path.join(folder, ANA_FILENAME)

    def _load(self):
        path = self._csv_path()
        if not path:
            messagebox.showwarning("Cartella non configurata",
                                   "Configura prima la cartella di lavoro nelle Impostazioni.",
                                   parent=self)
            return
        if not os.path.exists(path):
            # Crea file vuoto con intestazioni
            with open(path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f, delimiter=DELIMITER).writerow(ANA_HEADERS)
            self._rows = []
        else:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=DELIMITER)
                all_rows = list(reader)
            # Salta intestazione se presente
            if all_rows and all_rows[0] == ANA_HEADERS:
                data = all_rows[1:]
            else:
                data = all_rows
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

        # Intestazione
        hf = tk.Frame(self._inner, bg=C_HEADER_BG)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="#", width=3, bg=C_HEADER_BG, fg=C_HEADER_FG,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, padx=(4, 1), pady=3)
        for j, (h, w) in enumerate(zip(ANA_HEADERS, ANA_WIDTHS)):
            tk.Label(hf, text=h, width=w, bg=C_HEADER_BG, fg=C_HEADER_FG,
                     font=("Arial", 10, "bold"), justify="center",
                     anchor="center").grid(row=0, column=j + 1, padx=1, pady=3)

        for i, row_data in enumerate(self._rows):
            self._build_ana_row(i, row_data)

    def _build_ana_row(self, i: int, row_data: list[str]):
        bg = C_ROW_EVEN if i % 2 == 0 else C_ROW_ODD
        frame = tk.Frame(self._inner, bg=bg)
        frame.grid(row=i + 1, column=0, sticky="ew")

        num = tk.Label(frame, text=str(i + 1), width=3,
                       bg=bg, fg="#888", font=("Arial", 9), cursor="hand2")
        num.grid(row=0, column=0, padx=(4, 1))
        num.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))

        vars_: list[tk.StringVar] = []
        entries: list[tk.Widget]  = []

        for j, (val, w) in enumerate(zip(row_data, ANA_WIDTHS)):
            var = tk.StringVar(value=val)
            vars_.append(var)
            e = tk.Entry(frame, textvariable=var, width=w,
                         bg=C_LOCKED, relief="flat", font=("Arial", 10),
                         state="readonly", cursor="arrow",
                         readonlybackground=C_LOCKED)
            e.grid(row=0, column=j + 1, padx=1, pady=2, sticky="ew")
            e.bind("<Button-1>", lambda event, idx=i, col=j: self._try_unlock(idx, col))
            entries.append(e)

        self._row_widgets.append({"frame": frame, "vars": vars_, "entries": entries, "bg": bg})

    # ── Blocco/sblocco ────────────────────────────────────────────────────────
    def _try_unlock(self, row_idx: int, col_idx: int):
        entry: tk.Entry = self._row_widgets[row_idx]["entries"][col_idx]
        if entry.cget("state") == "normal":
            return  # già in modifica
        if not messagebox.askyesno(
                "Modifica campo",
                "Sei sicuro di voler modificare questo campo?\n"
                "La modifica sarà salvata immediatamente.",
                parent=self):
            return
        entry.configure(state="normal", bg=C_UNLOCKED,
                        readonlybackground=C_UNLOCKED, cursor="xterm")
        entry.focus_set()
        entry.bind("<Return>",   lambda _e, r=row_idx, c=col_idx: self._lock_and_save(r, c))
        entry.bind("<FocusOut>", lambda _e, r=row_idx, c=col_idx: self._lock_and_save(r, c))

    def _lock_and_save(self, row_idx: int, col_idx: int):
        winfo = self._row_widgets[row_idx]
        entry: tk.Entry = winfo["entries"][col_idx]
        if entry.cget("state") == "readonly":
            return  # già salvato (FocusOut può sparare due volte)
        # Aggiorna dati
        for j, var in enumerate(winfo["vars"]):
            self._rows[row_idx][j] = var.get()
        entry.configure(state="readonly", bg=C_LOCKED,
                        readonlybackground=C_LOCKED, cursor="arrow")
        self._save()
        self._update_status("Salvato")

    # ── Selezione riga ────────────────────────────────────────────────────────
    def _select_row(self, idx: int):
        if self._selected is not None and self._selected < len(self._row_widgets):
            old_bg = self._row_widgets[self._selected]["bg"]
            self._set_row_bg(self._selected, old_bg)
        self._selected = idx
        self._set_row_bg(idx, C_SELECTED)
        self._update_status()

    def _set_row_bg(self, idx: int, color: str):
        if idx is None or idx >= len(self._row_widgets):
            return
        winfo = self._row_widgets[idx]
        winfo["frame"].configure(bg=color)
        for child in winfo["frame"].winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass

    # ── Aggiungi / Elimina riga ───────────────────────────────────────────────
    def _add_row(self):
        new_row = ["1", "INV_ITEM", "", "", "PCS", "", "IT"]
        self._rows.append(new_row)
        i = len(self._rows) - 1
        self._build_ana_row(i, new_row)
        # Sblocca subito tutti i campi della nuova riga
        winfo = self._row_widgets[i]
        for j, entry in enumerate(winfo["entries"]):
            entry.configure(state="normal", bg=C_UNLOCKED,
                            readonlybackground=C_UNLOCKED, cursor="xterm")
            entry.bind("<Return>",   lambda _e, r=i, c=j: self._lock_and_save(r, c))
            entry.bind("<FocusOut>", lambda _e, r=i, c=j: self._lock_and_save(r, c))
        winfo["entries"][2].focus_set()  # focus su Descrizione
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

    # ── Status ────────────────────────────────────────────────────────────────
    def _update_status(self, extra: str = ""):
        parts = [f"{len(self._rows)} prodotti"]
        if self._selected is not None:
            parts.append(f"Riga {self._selected + 1} selezionata")
        if extra:
            parts.append(extra)
        self._status.set("  |  ".join(parts))

    # ── Esportazione dati per DocumentoWindow ─────────────────────────────────
    def get_products(self) -> list[list[str]]:
        return [list(r) for r in self._rows]


# ── Finestra Crea Documento ───────────────────────────────────────────────────
class DocumentoWindow(tk.Toplevel):
    """
    Crea un documento di spedizione: l'utente sceglie un nome file,
    aggiunge righe selezionando prodotti dall'anagrafica, compila quantità
    e prezzo. Il documento si auto-salva a ogni modifica (se il nome file
    è stato inserito). Può esportare un PDF.
    """

    def __init__(self, parent, config: ConfigManager):
        super().__init__(parent)
        self.title("Crea Documento")
        self.geometry("1200x640")
        self.minsize(900, 480)
        self._config = config
        self._products: list[list[str]] = []   # da anagrafica
        self._doc_rows: list[dict] = []        # ogni dict: {vars, widgets, selected}
        self._selected_row: int | None = None
        self._created_at: datetime | None = None
        self._last_saved: datetime | None = None
        self._autosave_pending = False
        self._apply_style()
        self._load_products()
        self._build_ui()
        self.transient(parent)

    def _apply_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

    def _load_products(self):
        folder = self._config.get_folder()
        if not folder:
            return
        path = os.path.join(folder, ANA_FILENAME)
        if not os.path.exists(path):
            return
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=DELIMITER)
            rows = list(reader)
        if rows and rows[0] == ANA_HEADERS:
            rows = rows[1:]
        n = len(ANA_HEADERS)
        self._products = [r[:n] + [""] * max(0, n - len(r)) for r in rows if any(c.strip() for c in r)]

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Toolbar superiore ─────────────────────────────────────────────────
        bar = tk.Frame(self, bg=C_TOOLBAR, pady=7, padx=10)
        bar.pack(side="top", fill="x")

        tk.Label(bar, text="Crea Documento", bg=C_TOOLBAR, fg="white",
                 font=("Arial", 12, "bold")).pack(side="left", padx=(0, 20))
        _btn(bar, "＋  Aggiungi riga",  "#2980b9", self._add_row)
        _btn(bar, "✕  Elimina riga",   "#c0392b", self._del_row)
        _btn(bar, "💾  Salva CSV",      "#27ae60", self._save_csv)
        _btn(bar, "🖨  Stampa PDF",     "#8e44ad", self._export_pdf)
        _btn(bar, "←  Chiudi",         "#7f8c8d", self.destroy, side="right")

        # ── Header nome file ──────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#ecf0f1", padx=14, pady=8)
        hdr.pack(side="top", fill="x")

        tk.Label(hdr, text="Nome file:", bg="#ecf0f1",
                 font=("Arial", 10, "bold")).pack(side="left")
        self._filename_var = tk.StringVar()
        self._filename_var.trace_add("write", self._on_filename_change)
        filename_entry = tk.Entry(hdr, textvariable=self._filename_var,
                                  width=24, font=("Arial", 11, "bold"))
        filename_entry.pack(side="left", padx=(6, 2))
        tk.Label(hdr, text=".csv", bg="#ecf0f1",
                 font=("Arial", 10), fg="#666").pack(side="left")

        tk.Label(hdr, text="  (es. 2025-600)", bg="#ecf0f1",
                 font=("Arial", 9), fg="#aaa").pack(side="left")

        self._hdr_status = tk.StringVar(value="")
        tk.Label(hdr, textvariable=self._hdr_status, bg="#ecf0f1",
                 font=("Arial", 9), fg="#27ae60").pack(side="right", padx=8)

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = tk.StringVar(value="Aggiungi righe al documento")
        tk.Label(self, textvariable=self._status, anchor="w",
                 bg="#ecf0f1", fg="#555", font=("Arial", 9),
                 padx=10, pady=3).pack(side="bottom", fill="x")

        # ── Tabella scrollabile ───────────────────────────────────────────────
        outer = tk.Frame(self, bg="white")
        outer.pack(fill="both", expand=True, padx=8, pady=(4, 4))

        self._canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical",   command=self._canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg="white")
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _e: self._canvas.configure(
                             scrollregion=self._canvas.bbox("all")))
        for w in (self._canvas, self._inner):
            w.bind("<MouseWheel>", self._on_scroll)
            w.bind("<Button-4>",  lambda _e: self._canvas.yview_scroll(-1, "units"))
            w.bind("<Button-5>",  lambda _e: self._canvas.yview_scroll( 1, "units"))

        self._build_header_row()

    def _on_scroll(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_header_row(self):
        hf = tk.Frame(self._inner, bg=C_HEADER_BG)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="#", width=3, bg=C_HEADER_BG, fg=C_HEADER_FG,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, padx=(4, 1), pady=3)
        for j, (h, w) in enumerate(zip(DOC_HEADERS, DOC_WIDTHS)):
            tk.Label(hf, text=h, width=w, bg=C_HEADER_BG, fg=C_HEADER_FG,
                     font=("Arial", 10, "bold"), justify="center",
                     anchor="center").grid(row=0, column=j + 1, padx=1, pady=3)

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

        # Numero riga
        num = tk.Label(frame, text=str(i + 1), width=3,
                       bg=bg, fg="#888", font=("Arial", 9), cursor="hand2")
        num.grid(row=0, column=0, padx=(4, 1))
        num.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))

        # ── Variabili per ogni colonna ─────────────────────────────────────
        # Ordine DOC_HEADERS: Rif. | Tipo | Descrizione | Cod. Doganale |
        #                     Q.tà/Peso | U.M. | Prezzo | Valuta |
        #                     Altro Prezzo | (vuoto) | Origine
        vars_ = {h: tk.StringVar() for h in DOC_HEADERS}

        # Valori predefiniti
        vars_["Rif."].set("1")
        vars_["Tipo"].set("INV_ITEM")
        vars_["Valuta"].set("EUR")

        widgets = {}

        def make_trace(v):
            v.trace_add("write", self._schedule_autosave)

        for h in DOC_HEADERS:
            make_trace(vars_[h])

        col_idx = 1

        # Rif. — readonly (pre-filled da anagrafica)
        e_rif = tk.Entry(frame, textvariable=vars_["Rif."],
                         width=DOC_WIDTHS[0], font=("Arial", 10),
                         bg=C_LOCKED, state="readonly",
                         readonlybackground=C_LOCKED, relief="flat")
        e_rif.grid(row=0, column=col_idx, padx=1, pady=2)
        e_rif.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Rif."] = e_rif
        col_idx += 1

        # Tipo — readonly
        e_tipo = tk.Entry(frame, textvariable=vars_["Tipo"],
                          width=DOC_WIDTHS[1], font=("Arial", 10),
                          bg=C_LOCKED, state="readonly",
                          readonlybackground=C_LOCKED, relief="flat")
        e_tipo.grid(row=0, column=col_idx, padx=1, pady=2)
        e_tipo.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Tipo"] = e_tipo
        col_idx += 1

        # Descrizione — Combobox selezionabile/digitabile
        desc_values = [p[2] for p in self._products] if self._products else []
        cb_desc = ttk.Combobox(frame, textvariable=vars_["Descrizione"],
                               values=desc_values, width=DOC_WIDTHS[2] - 2,
                               font=("Arial", 10))
        cb_desc.grid(row=0, column=col_idx, padx=1, pady=2)
        cb_desc.bind("<<ComboboxSelected>>",
                     lambda _e, idx=i: self._on_desc_selected(idx))
        cb_desc.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Descrizione"] = cb_desc
        col_idx += 1

        # Cod. Doganale — readonly
        e_cod = tk.Entry(frame, textvariable=vars_["Cod. Doganale"],
                         width=DOC_WIDTHS[3], font=("Arial", 10),
                         bg=C_LOCKED, state="readonly",
                         readonlybackground=C_LOCKED, relief="flat")
        e_cod.grid(row=0, column=col_idx, padx=1, pady=2)
        e_cod.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Cod. Doganale"] = e_cod
        col_idx += 1

        # Q.tà / Peso — editabile
        e_qty = tk.Entry(frame, textvariable=vars_["Q.tà / Peso"],
                         width=DOC_WIDTHS[4], font=("Arial", 10), relief="flat",
                         bg="white")
        e_qty.grid(row=0, column=col_idx, padx=1, pady=2)
        e_qty.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Q.tà / Peso"] = e_qty
        col_idx += 1

        # U.M. — readonly
        e_um = tk.Entry(frame, textvariable=vars_["U.M."],
                        width=DOC_WIDTHS[5], font=("Arial", 10),
                        bg=C_LOCKED, state="readonly",
                        readonlybackground=C_LOCKED, relief="flat")
        e_um.grid(row=0, column=col_idx, padx=1, pady=2)
        e_um.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["U.M."] = e_um
        col_idx += 1

        # Prezzo — editabile
        e_prezzo = tk.Entry(frame, textvariable=vars_["Prezzo"],
                            width=DOC_WIDTHS[6], font=("Arial", 10), relief="flat",
                            bg="white")
        e_prezzo.grid(row=0, column=col_idx, padx=1, pady=2)
        e_prezzo.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Prezzo"] = e_prezzo
        col_idx += 1

        # Valuta — Combobox EUR/USD/CHF
        cb_val = ttk.Combobox(frame, textvariable=vars_["Valuta"],
                              values=VALUTE, width=DOC_WIDTHS[7] - 2,
                              font=("Arial", 10), state="readonly")
        cb_val.grid(row=0, column=col_idx, padx=1, pady=2)
        cb_val.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Valuta"] = cb_val
        col_idx += 1

        # Altro Prezzo — readonly
        e_altro = tk.Entry(frame, textvariable=vars_["Altro Prezzo"],
                           width=DOC_WIDTHS[8], font=("Arial", 10),
                           bg=C_LOCKED, state="readonly",
                           readonlybackground=C_LOCKED, relief="flat")
        e_altro.grid(row=0, column=col_idx, padx=1, pady=2)
        e_altro.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Altro Prezzo"] = e_altro
        col_idx += 1

        # (vuoto)
        e_vuoto = tk.Entry(frame, textvariable=vars_[""],
                           width=DOC_WIDTHS[9], font=("Arial", 10),
                           bg=C_LOCKED, state="readonly",
                           readonlybackground=C_LOCKED, relief="flat")
        e_vuoto.grid(row=0, column=col_idx, padx=1, pady=2)
        widgets[""] = e_vuoto
        col_idx += 1

        # Origine — readonly
        e_orig = tk.Entry(frame, textvariable=vars_["Origine"],
                          width=DOC_WIDTHS[10], font=("Arial", 10),
                          bg=C_LOCKED, state="readonly",
                          readonlybackground=C_LOCKED, relief="flat")
        e_orig.grid(row=0, column=col_idx, padx=1, pady=2)
        e_orig.bind("<Button-1>", lambda _e, idx=i: self._select_row(idx))
        widgets["Origine"] = e_orig

        self._doc_rows.append({"frame": frame, "vars": vars_,
                                "widgets": widgets, "bg": bg})
        self._canvas.after(60, lambda: self._canvas.yview_moveto(1.0))
        widgets["Descrizione"].focus_set()
        self._update_status()

    def _on_desc_selected(self, row_idx: int):
        """Quando l'utente sceglie una descrizione, pre-popola i campi dall'anagrafica."""
        row = self._doc_rows[row_idx]
        desc = row["vars"]["Descrizione"].get()
        product = next((p for p in self._products if p[2] == desc), None)
        if product is None:
            return
        # ANA_HEADERS: Rif. | Tipo | Descrizione | Cod. Doganale | U.M. | Altro Prezzo | Origine
        rif, tipo, _, cod, um, altro, origine = (product + [""] * 7)[:7]
        for key, val in [("Rif.", rif), ("Tipo", tipo),
                         ("Cod. Doganale", cod), ("U.M.", um),
                         ("Altro Prezzo", altro), ("Origine", origine)]:
            w = row["widgets"][key]
            row["vars"][key].set(val)
            if isinstance(w, tk.Entry):
                w.configure(state="readonly", readonlybackground=C_LOCKED)

    # ── Selezione riga ────────────────────────────────────────────────────────
    def _select_row(self, idx: int):
        if self._selected_row is not None and self._selected_row < len(self._doc_rows):
            old = self._doc_rows[self._selected_row]
            self._set_row_bg_doc(self._selected_row, old["bg"])
        self._selected_row = idx
        self._set_row_bg_doc(idx, C_SELECTED)
        self._update_status()

    def _set_row_bg_doc(self, idx: int, color: str):
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
        # Rinumera
        for i, r in enumerate(self._doc_rows):
            r["bg"] = C_ROW_EVEN if i % 2 == 0 else C_ROW_ODD
            r["frame"].grid(row=i + 1, column=0, sticky="ew")
            for child in r["frame"].winfo_children():
                if isinstance(child, tk.Label):
                    child.configure(text=str(i + 1))
                    break
        self._schedule_autosave()
        self._update_status()

    # ── Nome file ─────────────────────────────────────────────────────────────
    def _on_filename_change(self, *_):
        self._schedule_autosave()

    def _get_output_path(self) -> str | None:
        folder = self._config.get_folder()
        name = self._filename_var.get().strip()
        if not folder or not name:
            return None
        if not name.endswith(".csv"):
            name += ".csv"
        return os.path.join(folder, name)

    # ── Auto-salvataggio ──────────────────────────────────────────────────────
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
        self._hdr_status.set(f"Salvato automaticamente — {now.strftime('%H:%M:%S')}")

    # ── Salva CSV manuale ─────────────────────────────────────────────────────
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
        self._hdr_status.set(f"Salvato — {now.strftime('%H:%M:%S')}")
        messagebox.showinfo("Salvato", f"Documento salvato in:\n{path}", parent=self)

    def _write_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=DELIMITER)
            for row in self._doc_rows:
                w.writerow([row["vars"][h].get() for h in DOC_HEADERS])

    def _collect_rows(self) -> list[list[str]]:
        return [[r["vars"][h].get() for h in DOC_HEADERS] for r in self._doc_rows]

    # ── Esporta PDF ───────────────────────────────────────────────────────────
    def _export_pdf(self):
        if not REPORTLAB_OK:
            messagebox.showerror(
                "ReportLab non disponibile",
                "La libreria 'reportlab' non è installata.\n"
                "Esegui:  pip install reportlab",
                parent=self)
            return

        filename = self._filename_var.get().strip() or "documento"
        folder   = self._config.get_folder() or tempfile.gettempdir()
        pdf_name = filename if filename.endswith(".pdf") else filename + ".pdf"
        pdf_path = os.path.join(folder, pdf_name)

        created_str  = (self._created_at or datetime.now()).strftime("%d/%m/%Y %H:%M")
        modified_str = (self._last_saved  or datetime.now()).strftime("%d/%m/%Y %H:%M")

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=landscape(A4),
            leftMargin=1.5 * cm, rightMargin=1.5 * cm,
            topMargin=1.5 * cm,  bottomMargin=1.5 * cm,
        )

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "DocTitle", parent=styles["Heading1"],
            fontSize=16, textColor=colors.HexColor("#2d3436"), spaceAfter=4
        )
        sub_style = ParagraphStyle(
            "DocSub", parent=styles["Normal"],
            fontSize=9, textColor=colors.HexColor("#636e72"), spaceAfter=12
        )
        cell_style = ParagraphStyle(
            "Cell", parent=styles["Normal"],
            fontSize=8, leading=10
        )

        story = [
            Paragraph(filename, title_style),
            Paragraph(
                f"Data creazione: <b>{created_str}</b> &nbsp;&nbsp; "
                f"Ultima modifica: <b>{modified_str}</b>",
                sub_style
            ),
        ]

        # Intestazione tabella (escludi colonna vuota "")
        visible_headers = [h for h in DOC_HEADERS if h != ""]
        col_indices      = [i for i, h in enumerate(DOC_HEADERS) if h != ""]

        table_data = [[Paragraph(f"<b>{h}</b>", cell_style) for h in visible_headers]]
        for row in self._doc_rows:
            vals = [row["vars"][DOC_HEADERS[i]].get() for i in col_indices]
            table_data.append([Paragraph(v, cell_style) for v in vals])

        # Larghezze colonne PDF (proporzionali)
        page_w = landscape(A4)[0] - 3 * cm
        weights = [1, 2, 8, 2.5, 2, 1.5, 2, 1.5, 2, 1.5]  # 10 colonne visibili
        total_w = sum(weights)
        col_widths = [page_w * (w / total_w) for w in weights]

        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0),  colors.HexColor("#2d3436")),
            ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.whitesmoke),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.HexColor("#f5f6fa"), colors.white]),
            ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#dfe6e9")),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",  (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
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
        self._status.set("  |  ".join(parts))


# ── Schermata principale ──────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DHL Spedizioni")
        self.geometry("480x300")
        self.resizable(False, False)
        self._config = ConfigManager()
        self._apply_style()
        self._build_ui()
        self._ensure_folder()

    def _apply_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

    def _build_ui(self):
        tk.Frame(self, bg=C_TOOLBAR, height=6).pack(fill="x")

        tk.Label(self, text="DHL Spedizioni", font=("Arial", 20, "bold"),
                 fg=C_TOOLBAR, pady=24).pack()

        btn_frame = tk.Frame(self, pady=4)
        btn_frame.pack()

        self._make_big_btn(btn_frame, "📋  Visualizza Anagrafica",
                           "#2980b9", self._open_anagrafica).pack(pady=6)
        self._make_big_btn(btn_frame, "📄  Crea Documento",
                           "#27ae60", self._open_documento).pack(pady=6)

        self._folder_label = tk.StringVar(value="")
        self._refresh_folder_label()

        bottom = tk.Frame(self, bg="#ecf0f1", pady=5, padx=10)
        bottom.pack(side="bottom", fill="x")
        tk.Label(bottom, textvariable=self._folder_label,
                 bg="#ecf0f1", fg="#888", font=("Arial", 8),
                 anchor="w").pack(side="left", fill="x", expand=True)
        tk.Button(bottom, text="⚙ Impostazioni", font=("Arial", 8),
                  bg="#ecf0f1", fg="#555", relief="flat", cursor="hand2",
                  command=self._open_settings).pack(side="right")

    @staticmethod
    def _make_big_btn(parent, text, color, cmd):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg="white", relief="flat",
                      font=("Arial", 11, "bold"), padx=24, pady=10,
                      width=28, cursor="hand2",
                      activeforeground="white", activebackground=color, bd=0)
        darker = _darken(color)
        b.bind("<Enter>", lambda _e: b.configure(bg=darker))
        b.bind("<Leave>", lambda _e: b.configure(bg=color))
        return b

    def _refresh_folder_label(self):
        folder = self._config.get_folder()
        self._folder_label.set(f"Cartella: {folder}" if folder else "Cartella non configurata")

    def _ensure_folder(self):
        """Al primo avvio chiede la cartella di lavoro; copia l'anagrafica di default se mancante."""
        if not self._config.get_folder():
            messagebox.showinfo(
                "Benvenuto",
                "Seleziona la cartella di lavoro dove verranno salvati\n"
                "l'anagrafica prodotti e i documenti esportati.",
            )
            path = filedialog.askdirectory(title="Seleziona cartella di lavoro",
                                           initialdir=str(Path.home()))
            if path:
                self._config.set_folder(path)
                self._refresh_folder_label()
                self._copy_default_anagrafica(path)

    def _copy_default_anagrafica(self, folder: str):
        """Copia l'anagrafica di default nella cartella di lavoro se non esiste già."""
        dest = os.path.join(folder, ANA_FILENAME)
        if os.path.exists(dest):
            return
        # Cerca il file accanto all'eseguibile o allo script
        candidates = [
            os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), ANA_FILENAME),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), ANA_FILENAME),
        ]
        # Supporto PyInstaller bundle
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
