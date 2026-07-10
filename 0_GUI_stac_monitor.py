"""
0_GUI_stac_monitor.py  –  STAC Monitoring-Tool (read-only)

Zeigt Items und Assets der Collection "ch.swisstopo.spezialbefliegungen"
in einer Baumansicht. Funktionen:
  - Asset-Status-Prüfung via HEAD (HTTP-Code, Dateigrösse, Last-Modified)
  - Statistik: OK / Fehler / Gesamtgrösse
  - Export Download-Links (JSON für Kunden)
  - Export Tabelle (CSV für interne Auswertung)
  - Item-JSON Detailansicht (Doppelklick oder Rechtsklick)
  - URL in Zwischenablage kopieren, im Browser öffnen

Credentials: secrets/stac_credentials.json
Format:      {"INT": {"username": "...", "password": "..."}, "PROD": {...}}
"""

import csv
import ctypes
import io
import json
import threading
import concurrent.futures
import webbrowser
from datetime import datetime
from email.utils import parsedate
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

from stac_api import (
    COLLECTION_ID, ENVIRONMENTS, AUFTRAGSTYPEN, EXT_PRESETS,
    get_item_direct, get_collection_items, filter_items,
    check_asset_info, browser_url, asset_area,
    stac_item_year, stac_item_area, stac_item_acq_date,
    build_stac_item, is_cog_asset, map_viewer_url,
)


# ─── Farbpaletten ─────────────────────────────────────────────────────────────

LIGHT = {
    "root":       "#f0f0f0",
    "panel":      "#f5f5f5",
    "input":      "#ffffff",
    "fg":         "#1a1a1a",
    "fg_dim":     "#666666",
    "accent":     "#0063b1",
    "hdr_bg":     "#1a3a5c",
    "hdr_fg":     "#ffffff",
    "btn":        "#e1e1e1",
    "btn_hover":  "#c8c8c8",
    "list":       "#ffffff",
    "log_bg":     "#1e1e1e",
    "log_fg":     "#d4d4d4",
    "sep":        "#c0c0c0",
    "sel_bg":     "#0078d4",
    "sel_fg":     "#ffffff",
    "ok":         "#2e7d32",
    "err":        "#c62828",
    "warn":       "#8a6f2e",
    "tree_item":  "#0063b1",
    "tree_ok":    "#2e7d32",
    "tree_err":   "#c62828",
    "tree_warn":  "#8a6f2e",
    "tree_dim":   "#888888",
}

DARK = {
    "root":       "#1e1e1e",
    "panel":      "#252526",
    "input":      "#3c3c3c",
    "fg":         "#cccccc",
    "fg_dim":     "#7a7a7a",
    "accent":     "#4fc3f7",
    "hdr_bg":     "#1a1a1a",
    "hdr_fg":     "#cccccc",
    "btn":        "#3c3c3c",
    "btn_hover":  "#505050",
    "list":       "#2d2d30",
    "log_bg":     "#1e1e1e",
    "log_fg":     "#d4d4d4",
    "sep":        "#3c3c3c",
    "sel_bg":     "#094771",
    "sel_fg":     "#cccccc",
    "ok":         "#66bb6a",
    "err":        "#ef5350",
    "warn":       "#c9a84c",
    "tree_item":  "#4fc3f7",
    "tree_ok":    "#66bb6a",
    "tree_err":   "#ef5350",
    "tree_warn":  "#c9a84c",
    "tree_dim":   "#7a7a7a",
}


# ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

def _fmt_size(size_bytes: Optional[int]) -> str:
    if size_bytes is None:
        return "–"
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 ** 2:
        return f"{size_bytes / 1024:.1f} KB"
    if size_bytes < 1024 ** 3:
        return f"{size_bytes / 1024 ** 2:.1f} MB"
    return f"{size_bytes / 1024 ** 3:.2f} GB"


def _fmt_date(lm_str: Optional[str]) -> str:
    """Parst HTTP Last-Modified-Header auf YYYY-MM-DD."""
    if not lm_str:
        return "–"
    try:
        t = parsedate(lm_str)
        if t:
            return f"{t[0]}-{t[1]:02d}-{t[2]:02d}"
    except Exception:
        pass
    return lm_str[:10] if len(lm_str) >= 10 else lm_str


def _status_label(sc: Optional[int]) -> Tuple[str, str]:
    """Gibt (Anzeigetext, Tag-Name) für einen HTTP-Statuscode zurück."""
    if sc is None:
        return "–", "asset_dim"
    if sc == 200:
        return "✓  200", "asset_ok"
    if sc > 0:
        return f"✗  {sc}", "asset_err"
    if sc == -2:
        return "✗  timeout", "asset_warn"
    return "✗  err", "asset_warn"


# ─── Item-JSON Popup ──────────────────────────────────────────────────────────

class ItemJsonDialog(tk.Toplevel):
    def __init__(self, parent, item: Dict, dark: bool):
        super().__init__(parent)
        T = DARK if dark else LIGHT
        self.title(f"STAC Item  —  {item.get('id', '')}")
        self.configure(bg=T["root"])
        self.minsize(720, 520)

        txt = scrolledtext.ScrolledText(
            self, font=("Cascadia Mono", 9),
            bg=T["log_bg"], fg=T["log_fg"],
            insertbackground=T["log_fg"],
        )
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", json.dumps(item, indent=2, ensure_ascii=False))
        txt.config(state="disabled")

        btn_row = tk.Frame(self, bg=T["root"])
        btn_row.pack(fill="x", padx=8, pady=(0, 8))

        def _copy_id():
            self.clipboard_clear()
            self.clipboard_append(item.get("id", ""))

        tk.Button(btn_row, text="Item-ID kopieren",
                  bg=T["btn"], fg=T["fg"], relief="flat", padx=10, pady=4,
                  command=_copy_id).pack(side="left")
        tk.Button(btn_row, text="Schliessen",
                  bg=T["btn"], fg=T["fg"], relief="flat", padx=10, pady=4,
                  command=self.destroy).pack(side="right")

        self.transient(parent)
        self.grab_set()


# ─── Export-Vorschau (Anzeigen statt sofort speichern) ───────────────────────

class ExportPreviewDialog(tk.Toplevel):
    """Zeigt generierten Export-Inhalt an; Speichern erfolgt erst auf Wunsch."""

    def __init__(self, parent, dark: bool, title: str, content: str,
                 initialfile: str, filetypes: List[Tuple[str, str]],
                 defaultextension: str, encoding: str = "utf-8",
                 write_newline: str = None, on_saved=None):
        super().__init__(parent)
        T = DARK if dark else LIGHT
        self.title(title)
        self.configure(bg=T["root"])
        self.minsize(720, 520)

        txt = scrolledtext.ScrolledText(
            self, font=("Cascadia Mono", 9),
            bg=T["log_bg"], fg=T["log_fg"],
            insertbackground=T["log_fg"],
        )
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", content)
        txt.config(state="disabled")

        btn_row = tk.Frame(self, bg=T["root"])
        btn_row.pack(fill="x", padx=8, pady=(0, 8))

        def _save():
            path = filedialog.asksaveasfilename(
                defaultextension=defaultextension,
                filetypes=filetypes,
                title=title,
                initialfile=initialfile,
            )
            if not path:
                return
            try:
                with open(path, "w", newline=write_newline, encoding=encoding) as f:
                    f.write(content)
                if on_saved:
                    on_saved(path)
            except Exception as exc:
                messagebox.showerror("Export-Fehler", str(exc))

        tk.Button(btn_row, text="Speichern unter...",
                  bg=T["btn"], fg=T["fg"], relief="flat", padx=10, pady=4,
                  command=_save).pack(side="left")
        tk.Button(btn_row, text="Schliessen",
                  bg=T["btn"], fg=T["fg"], relief="flat", padx=10, pady=4,
                  command=self.destroy).pack(side="right")

        self.transient(parent)
        self.grab_set()


# ─── Haupt-Applikation ────────────────────────────────────────────────────────

class StacMonitorApp(tk.Tk):

    _COLS      = ("sel", "area", "status", "typ", "groesse", "geaendert")
    _COL_HEADS = {"sel": "Auswahl", "area": "Area", "status": "Status", "typ": "Typ / Ext.",
                  "groesse": "Grösse", "geaendert": "Geändert"}
    _COL_W     = {"sel": 60, "area": 90, "status": 100, "typ": 90,
                  "groesse": 90, "geaendert": 105}

    # Kreis-Glyphen aus dem Unicode-BMP-Bereich (Geometric Shapes, U+25xx/U+2Bxx),
    # etwas grösser als die ursprünglichen ●/○-Zeichen. Farbige Emoji-Kreise
    # (z.B. 🟢/🟡, U+1F7Ex) liegen ausserhalb der BMP (>U+FFFF) und lassen
    # manche Tcl/Tk-Builds (z.B. Python 3.6 auf Windows) mit
    # "character U+... is above the range (U+0000-U+FFFF) allowed by Tcl"
    # abstürzen – deshalb ausschliesslich BMP-Zeichen. Die Amber-Einfärbung bei
    # Auswahl erfolgt stattdessen über Zeilen-Tags (siehe _asset_tag/_item_tag),
    # da ttk.Treeview keine Einzelzell-Farbe kennt, nur zeilenweise Tags.
    _CHK_ON      = "⬤"
    _CHK_OFF     = "◯"
    _CHK_PARTIAL = "◐"

    _LOAD_BTN_LABEL    = "ITEM-Liste laden"
    _RELOAD_BTN_LABEL  = "ITEM-Liste aktualisieren"
    _SPINNER_FRAMES    = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"

    def __init__(self):
        super().__init__()
        self.title("STAC Monitor  —  ch.swisstopo.spezialbefliegungen")
        self.minsize(1040, 720)

        self._dark: bool = True
        self._auth: Optional[Tuple] = None
        self._base_url: str = ""

        self._all_items: List[Dict] = []
        self._visible_items: List[Dict] = []
        self._items_loaded_once: bool = False

        # Baum-Metadaten: tree_iid → dict mit kind/item_id/asset_key/href/item
        self._nodes: Dict[str, Dict] = {}
        # Prüfergebnisse: {item_id: {asset_key: {status, size_bytes, last_modified}}}
        self._asset_info: Dict[str, Dict[str, Dict]] = {}
        # Export-Auswahl je Asset-Knoten (tree_iid → bool). Fehlender Eintrag = gewählt.
        self._checked: Dict[str, bool] = {}

        # Lade-Spinner im "ITEM-Liste laden"-Button
        self._spinner_job: Optional[str] = None
        self._spinner_idx: int = 0

        self._build_ui()
        self._apply_theme(True)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        self._hdr = tk.Frame(self, height=52)
        self._hdr.pack(fill="x")
        self._hdr.pack_propagate(False)
        self._hdr_lbl = tk.Label(
            self._hdr,
            text="  STAC Monitor  —  ch.swisstopo.spezialbefliegungen  [read-only]",
            font=("Segoe UI", 13, "bold"),
        )
        self._hdr_lbl.pack(side="left", padx=16, pady=10)
        self._theme_btn = tk.Button(
            self._hdr, text="Hell", relief="flat", borderwidth=0,
            font=("Segoe UI", 9), cursor="hand2", padx=10, pady=4,
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side="right", padx=12)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=12, pady=8)

        self._build_credentials(main)
        self._build_filters(main)
        self._build_actions(main)
        self._build_stac_functions(main)
        self._build_tree(main)
        self._build_stats(main)
        self._build_log(main)

    def _build_credentials(self, parent):
        sec = ttk.LabelFrame(parent, text="1   Umgebung & Credentials",
                             padding=8, style="Section.TLabelframe")
        sec.pack(fill="x", pady=(0, 4))

        ttk.Label(sec, text="Umgebung:").pack(side="left", padx=(0, 6))
        self._env_var = tk.StringVar(value="INT")
        for env in ("INT", "PROD"):
            ttk.Radiobutton(sec, text=env, variable=self._env_var, value=env,
                            command=self._on_env_change).pack(side="left", padx=4)

        self._url_lbl = ttk.Label(sec, text=ENVIRONMENTS["INT"],
                                   font=("Segoe UI", 8), style="Dim.TLabel")
        self._url_lbl.pack(side="left", padx=12)

        ttk.Button(sec, text="STAC Browser öffnen",
                   command=self._open_stac_browser).pack(side="left", padx=(0, 12))

        self._cred_btn = ttk.Button(sec, text="Credentials laden",
                                     command=self._load_credentials, style="Amber.TButton")
        self._cred_btn.pack(side="left", padx=(12, 6))

        self._cred_lbl = ttk.Label(sec, text="nicht geladen",
                                    font=("Segoe UI", 9, "italic"), style="Dim.TLabel")
        self._cred_lbl.pack(side="left")

    def _build_filters(self, parent):
        sec = ttk.LabelFrame(parent, text="2   Filter",
                             padding=8, style="Section.TLabelframe")
        sec.pack(fill="x", pady=(0, 4))
        sec.columnconfigure(4, weight=1)

        # Auftragstyp
        ttk.Label(sec, text="Auftragstyp:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._auftragstyp_var = tk.StringVar(value=list(AUFTRAGSTYPEN.keys())[0])
        col = 1
        for typ in AUFTRAGSTYPEN:
            ttk.Radiobutton(sec, text=typ, variable=self._auftragstyp_var, value=typ,
                            command=self._on_auftragstyp_change).grid(
                row=0, column=col, sticky="w", padx=(0, 12))
            col += 1

        # Jahr + Suche
        ttk.Label(sec, text="Jahr [optional]:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self._year_var = tk.StringVar()
        self._year_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._year_var, width=8).grid(
            row=1, column=1, sticky="w", pady=(6, 0))

        ttk.Label(sec, text="Item-ID / Suche [optional]:").grid(
            row=1, column=2, sticky="w", padx=(16, 6), pady=(6, 0))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._search_var, width=34).grid(
            row=1, column=3, sticky="w", pady=(6, 0))
        ttk.Label(sec, text="Teilstring genügt  (für direkten Abruf: vollständige ID)",
                  font=("Segoe UI", 8), style="Dim.TLabel").grid(
            row=1, column=4, sticky="w", padx=(8, 0), pady=(6, 0))

        # Suchfeld gleich mit dem Default-Auftragstyp vorbefüllen (Radiobutton-Command
        # feuert sonst erst bei einem tatsächlichen Klick, nicht bei der Vorauswahl).
        self._on_auftragstyp_change()

        # Dateiendung
        ttk.Label(sec, text="Dateiendung:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ext_frame = ttk.Frame(sec)
        ext_frame.grid(row=2, column=1, columnspan=4, sticky="w", pady=(6, 0))
        self._ext_vars: List[Tuple[tk.BooleanVar, List[str]]] = []
        for label, exts in EXT_PRESETS:
            var = tk.BooleanVar(value=False)
            var.trace_add("write", lambda *_: self._apply_filters())
            self._ext_vars.append((var, exts))
            ttk.Checkbutton(ext_frame, text=label, variable=var).pack(side="left", padx=(0, 10))
        ttk.Label(ext_frame, text="freies Suffix [optional]:").pack(side="left", padx=(6, 4))
        self._ext_custom_var = tk.StringVar()
        self._ext_custom_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(ext_frame, textvariable=self._ext_custom_var, width=14).pack(side="left")

    def _build_actions(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 4))

        self._load_btn = ttk.Button(
            row, text=self._LOAD_BTN_LABEL, command=self._load, state="disabled",
            style="AmberBold.TButton")
        self._load_btn.pack(side="left", padx=(0, 16))

        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=(0, 16))

        self._expand_btn = ttk.Button(
            row, text="Alle aufklappen", command=self._expand_all, state="disabled")
        self._expand_btn.pack(side="left", padx=(0, 4))

        self._collapse_btn = ttk.Button(
            row, text="Alle einklappen", command=self._collapse_all, state="disabled")
        self._collapse_btn.pack(side="left", padx=(0, 16))

    def _build_stac_functions(self, parent):
        sec = ttk.LabelFrame(parent, text="STAC-Funktionen",
                             padding=8, style="Section.TLabelframe")
        sec.pack(fill="x", pady=(0, 4))

        self._check_btn = ttk.Button(
            sec, text="Assets prüfen  (HEAD)", command=self._check_assets, state="disabled")
        self._check_btn.pack(side="left", padx=(0, 16))

        ttk.Separator(sec, orient="vertical").pack(side="left", fill="y", padx=(0, 16))

        self._export_json_btn = ttk.Button(
            sec, text="Export JSON (Kunden-Links)",
            command=self._export_json, state="disabled")
        self._export_json_btn.pack(side="left", padx=(0, 4))

        self._export_csv_btn = ttk.Button(
            sec, text="Export CSV", command=self._export_csv, state="disabled")
        self._export_csv_btn.pack(side="left", padx=(0, 4))

        self._export_links_btn = ttk.Button(
            sec, text="Item - STAC Browser Links",
            command=self._export_stac_browser_links, state="disabled")
        self._export_links_btn.pack(side="left", padx=(0, 4))

        self._map_viewer_btn = ttk.Button(
            sec, text="Link auf Kartenviewer",
            command=self._open_map_viewer, state="disabled")
        self._map_viewer_btn.pack(side="left")

    def _build_tree(self, parent):
        frame = ttk.LabelFrame(parent, text="3   Items & Assets",
                               padding=4, style="Section.TLabelframe")
        frame.pack(fill="both", expand=True, pady=(0, 4))
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(frame)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        self._select_all_btn = ttk.Button(
            toolbar, text="Alle auswählen", command=self._select_all, state="disabled")
        self._select_all_btn.pack(side="left", padx=(0, 4))

        self._deselect_all_btn = ttk.Button(
            toolbar, text="Alles abwählen", command=self._deselect_all, state="disabled")
        self._deselect_all_btn.pack(side="left")

        self._tree = ttk.Treeview(
            frame, columns=self._COLS, show="tree headings", selectmode="browse")

        self._tree.column("#0", width=320, minwidth=200, stretch=False)
        self._tree.heading("#0", text="Item / Asset")
        for col in self._COLS:
            self._tree.column(col, width=self._COL_W[col],
                              minwidth=55, stretch=False, anchor="center")
            self._tree.heading(col, text=self._COL_HEADS[col])

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")

        self._ctx = tk.Menu(self, tearoff=0)
        self._tree.bind("<Button-3>",  self._on_right_click)
        self._tree.bind("<Double-1>",  self._on_double_click)
        self._tree.bind("<Button-1>",  self._on_tree_click)

    def _build_stats(self, parent):
        self._stats_outer = tk.Frame(parent)
        self._stats_outer.pack(fill="x", pady=(0, 2))
        self._stats_lbl = tk.Label(
            self._stats_outer, text="Keine Daten geladen.",
            font=("Segoe UI", 9), anchor="w")
        self._stats_lbl.pack(side="left", padx=4)

    def _build_log(self, parent):
        frm = ttk.LabelFrame(parent, text="Log",
                              padding=4, style="Section.TLabelframe")
        frm.pack(fill="x")
        self._log = scrolledtext.ScrolledText(
            frm, height=5, state="disabled",
            font=("Cascadia Mono", 8), wrap="word")
        self._log.pack(fill="both")

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _toggle_theme(self):
        self._apply_theme(not self._dark)

    def _apply_theme(self, dark: bool):
        self._dark = dark
        T = DARK if dark else LIGHT

        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",
            background=T["panel"], foreground=T["fg"],
            fieldbackground=T["input"],
            selectbackground=T["sel_bg"], selectforeground=T["sel_fg"],
            bordercolor=T["sep"], lightcolor=T["panel"], darkcolor=T["sep"],
            insertcolor=T["fg"], troughcolor=T["root"])
        s.configure("TFrame",    background=T["panel"])
        s.configure("TLabelframe",
            background=T["panel"], bordercolor=T["sep"])
        s.configure("TLabelframe.Label",
            background=T["panel"], foreground=T["fg"], font=("Segoe UI", 9, "bold"))
        s.configure("Section.TLabelframe",
            background=T["panel"], bordercolor=T["sep"])
        s.configure("Section.TLabelframe.Label",
            background=T["panel"], foreground=T["accent"], font=("Segoe UI", 10, "bold"))
        s.configure("TLabel",    background=T["panel"], foreground=T["fg"])
        s.configure("Dim.TLabel", background=T["panel"], foreground=T["fg_dim"])
        s.configure("TButton",
            background=T["btn"], foreground=T["fg"],
            bordercolor=T["sep"], relief="flat", padding=(8, 4), focuscolor=T["panel"])
        s.map("TButton",
            background=[("active", T["btn_hover"]), ("pressed", T["sep"])],
            foreground=[("active", T["fg"])],
            relief=[("pressed", "flat")])
        s.configure("Amber.TButton",
            background=T["btn"], foreground=T["warn"],
            bordercolor=T["sep"], relief="flat", padding=(8, 4), focuscolor=T["panel"])
        s.map("Amber.TButton",
            background=[("active", T["btn_hover"]), ("pressed", T["sep"])],
            foreground=[("active", T["warn"])],
            relief=[("pressed", "flat")])
        s.configure("AmberBold.TButton",
            background=T["btn"], foreground=T["warn"],
            bordercolor=T["sep"], relief="flat", padding=(8, 4), focuscolor=T["panel"],
            font=("Segoe UI", 9, "bold"))
        s.map("AmberBold.TButton",
            background=[("active", T["btn_hover"]), ("pressed", T["sep"])],
            foreground=[("active", T["warn"])],
            relief=[("pressed", "flat")])
        s.configure("TRadiobutton",
            background=T["panel"], foreground=T["fg"], focuscolor=T["panel"])
        s.map("TRadiobutton",
            background=[("active", T["panel"])], foreground=[("active", T["fg"])])
        s.configure("TCheckbutton",
            background=T["panel"], foreground=T["fg"], focuscolor=T["panel"])
        s.map("TCheckbutton",
            background=[("active", T["panel"])], foreground=[("active", T["fg"])])
        s.configure("TEntry",
            fieldbackground=T["input"], foreground=T["fg"],
            bordercolor=T["sep"], insertcolor=T["fg"],
            selectbackground=T["sel_bg"], selectforeground=T["sel_fg"])
        s.configure("Vertical.TScrollbar",
            background=T["btn"], troughcolor=T["root"],
            bordercolor=T["sep"], arrowcolor=T["fg"])
        s.configure("Horizontal.TScrollbar",
            background=T["btn"], troughcolor=T["root"],
            bordercolor=T["sep"], arrowcolor=T["fg"])
        s.configure("Treeview",
            background=T["list"], foreground=T["fg"],
            fieldbackground=T["list"], rowheight=22, bordercolor=T["sep"])
        s.configure("Treeview.Heading",
            background=T["btn"], foreground=T["fg"],
            relief="flat", padding=(4, 4))
        s.map("Treeview",
            background=[("selected", T["sel_bg"])],
            foreground=[("selected", T["sel_fg"])])
        s.map("Treeview.Heading",
            background=[("active", T["btn_hover"])])

        self._tree.tag_configure("item",
            foreground=T["tree_item"], font=("Segoe UI", 9, "bold"))
        # Amber (bestehende warn-Akzentfarbe), sobald vollständig ausgewählt –
        # bei Assets nur solange noch kein HTTP-Prüfergebnis vorliegt, damit die
        # aussagekräftigere ok/err/warn-Statusfarbe nach der Prüfung erhalten bleibt.
        self._tree.tag_configure("item_selected",
            foreground=T["tree_warn"], font=("Segoe UI", 9, "bold"))
        self._tree.tag_configure("asset_selected", foreground=T["tree_warn"])
        self._tree.tag_configure("asset_ok",   foreground=T["tree_ok"])
        self._tree.tag_configure("asset_err",  foreground=T["tree_err"])
        self._tree.tag_configure("asset_warn", foreground=T["tree_warn"])
        self._tree.tag_configure("asset_dim",  foreground=T["tree_dim"])

        self.configure(bg=T["root"])
        self._hdr.configure(bg=T["hdr_bg"])
        self._hdr_lbl.configure(bg=T["hdr_bg"], fg=T["hdr_fg"])
        self._theme_btn.configure(
            bg=T["hdr_bg"], fg=T["hdr_fg"],
            activebackground=T["btn"], activeforeground=T["fg"],
            text="Hell" if dark else "Dark")
        self._log.configure(bg=T["log_bg"], fg=T["log_fg"],
                             insertbackground=T["log_fg"])
        self._stats_outer.configure(bg=T["panel"])
        self._stats_lbl.configure(bg=T["panel"], fg=T["fg_dim"])
        self._ctx.configure(
            bg=T["btn"], fg=T["fg"],
            activebackground=T["sel_bg"], activeforeground=T["sel_fg"])

        self._set_titlebar_dark(dark)

    def _set_titlebar_dark(self, dark: bool):
        if not self.winfo_ismapped():
            self.after(50, lambda: self._set_titlebar_dark(dark))
            return
        try:
            hwnd  = int(self.wm_frame(), 16)
            value = ctypes.c_int(1 if dark else 0)
            for attr in (20, 19):
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                    break
            ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0027)
        except Exception:
            pass

    # ── Event Handler ─────────────────────────────────────────────────────────

    def _on_env_change(self):
        self._url_lbl.configure(text=ENVIRONMENTS[self._env_var.get()])
        self._auth = None
        self._cred_lbl.configure(text="nicht geladen")
        self._cred_btn.configure(style="Amber.TButton")
        self._load_btn.config(state="disabled")
        self._load_btn.configure(style="AmberBold.TButton")

    def _open_stac_browser(self, item_id: Optional[str] = None):
        url = browser_url(self._env_var.get(), item_id)
        webbrowser.open(url)
        self.clipboard_clear()
        self.clipboard_append(url)
        self._log_write(f"[STAC Browser] geöffnet & kopiert: {url}\n")

    def _on_auftragstyp_change(self):
        typ     = self._auftragstyp_var.get()
        suggest = AUFTRAGSTYPEN[typ]
        known   = set(AUFTRAGSTYPEN.values())
        cur     = self._search_var.get().strip()
        if not cur or cur in known:
            self._search_var.set(suggest)

    def _load_credentials(self):
        env = self._env_var.get()
        try:
            cfg_path = Path(__file__).parent / "secrets" / "stac_credentials.json"
            with open(cfg_path, encoding="utf-8") as f:
                cfg = json.load(f)
            creds          = cfg[env]
            self._auth     = (creds["username"], creds["password"])
            self._base_url = ENVIRONMENTS[env]
            T = DARK if self._dark else LIGHT
            self._cred_lbl.configure(
                text=f"Geladen: {creds['username']}", foreground=T["ok"])
            self._cred_btn.configure(style="TButton")
            self._load_btn.config(state="normal")
            self._log_write(f"[Credentials] {env} – {creds['username']}\n")
        except Exception as exc:
            T = DARK if self._dark else LIGHT
            self._cred_lbl.configure(text="Fehler!", foreground=T["err"])
            messagebox.showerror("Credentials-Fehler", str(exc))

    # ── Laden ─────────────────────────────────────────────────────────────────

    def _load(self):
        if not self._auth:
            return
        self._load_btn.configure(style="TButton")
        self._all_items.clear()
        self._asset_info.clear()
        self._visible_items = []
        self._populate_tree([], [])  # Bestehende Liste sofort leeren, bevor neu geladen wird
        self._set_busy(True)
        search = self._search_var.get().strip()
        threading.Thread(target=self._worker_load, args=(search,), daemon=True).start()

    def _worker_load(self, search: str):
        try:
            if search:
                self._log_write(f"[Laden] Prüfe exakte Item-ID: {search} …\n")
                item = get_item_direct(self._base_url, self._auth, search)
                if item is not None:
                    self._all_items = [item]
                    self._log_write(f"[OK] {item['id']} geladen (Direct-Lookup).\n")
                    self.after(0, self._apply_filters)
                    return
                self._log_write("[Info] Keine exakte Übereinstimmung – "
                                "lade gesamte Collection …\n")
            else:
                self._log_write("[Laden] Hole alle Items der Collection …\n")
            items = get_collection_items(self._base_url, self._auth, self._log_write)
            self._all_items = items
            self._log_write(f"[Laden] {len(items)} Items geladen.\n")
            self.after(0, self._apply_filters)
        except Exception as exc:
            self._log_write(f"[FEHLER] {exc}\n")
            self.after(0, lambda: messagebox.showerror("Fehler", str(exc)))
        finally:
            self.after(0, lambda: self._set_busy(False))

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self._load_btn.config(state=state if self._auth else "disabled")
        if busy:
            self._check_btn.config(state="disabled")
            self._export_json_btn.config(state="disabled")
            self._export_csv_btn.config(state="disabled")
            self._export_links_btn.config(state="disabled")
            self._map_viewer_btn.config(state="disabled")
            self._start_load_spinner()
        else:
            self._stop_load_spinner()

    def _start_load_spinner(self):
        self._spinner_idx = 0
        self._animate_load_spinner()

    def _animate_load_spinner(self):
        frame = self._SPINNER_FRAMES[self._spinner_idx % len(self._SPINNER_FRAMES)]
        self._load_btn.config(text=f"{frame}  Lade Items …")
        self._spinner_idx += 1
        self._spinner_job = self.after(120, self._animate_load_spinner)

    def _stop_load_spinner(self):
        if self._spinner_job is not None:
            self.after_cancel(self._spinner_job)
            self._spinner_job = None
        label = self._RELOAD_BTN_LABEL if self._items_loaded_once else self._LOAD_BTN_LABEL
        self._load_btn.config(text=label)

    # ── Filter + Treeview ─────────────────────────────────────────────────────

    def _active_extensions(self) -> List[str]:
        result = []
        for var, exts in self._ext_vars:
            if var.get():
                result.extend(exts)
        for part in self._ext_custom_var.get().replace(",", " ").split():
            result.append(part if part.startswith(".") else f".{part}")
        return result

    def _apply_filters(self):
        if not self._all_items:
            return
        self._items_loaded_once = True
        year   = self._year_var.get().strip()
        search = self._search_var.get().strip()
        exts   = self._active_extensions()

        items = self._all_items
        if search:
            items = filter_items(items, search)
        if year:
            items = [it for it in items if stac_item_year(it) == year]
        if exts:
            def _has_match(it):
                for k, v in it.get("assets", {}).items():
                    href = v.get("href", "")
                    if any(href.lower().endswith(e) or k.lower().endswith(e) for e in exts):
                        return True
                return False
            items = [it for it in items if _has_match(it)]

        self._visible_items = items
        self._populate_tree(items, exts)

    def _populate_tree(self, items: List[Dict], exts: List[str]):
        self._tree.delete(*self._tree.get_children())
        self._nodes.clear()

        if not items:
            self._stats_lbl.configure(text="Keine Items nach aktuellem Filter.")
            self._toggle_tree_buttons(False)
            return

        sorted_items = sorted(items, key=stac_item_acq_date, reverse=True)
        total_assets = 0
        _pfx         = COLLECTION_ID + "_"

        for item in sorted_items:
            iid     = item["id"]
            area    = stac_item_area(item)
            acq     = stac_item_acq_date(item)
            display = iid[len(_pfx):] if iid.startswith(_pfx) else iid

            assets = item.get("assets", {})
            if exts:
                asset_keys = [
                    k for k, v in assets.items()
                    if any(v.get("href", "").lower().endswith(e) or k.lower().endswith(e)
                           for e in exts)
                ]
            else:
                asset_keys = list(assets.keys())

            total_assets += len(asset_keys)

            meta = "  ".join(p for p in [area, acq] if p)
            label = display + (f"   [{meta}]" if meta else "")

            asset_node_ids = [f"asset::{iid}::{ak}" for ak in asset_keys]

            node_id    = f"item::{iid}"
            item_glyph = self._item_check_glyph(asset_node_ids)
            self._tree.insert("", "end", iid=node_id,
                              text=f"  {label}",
                              values=(item_glyph, area,
                                      "", "", f"{len(asset_keys)} Assets", ""),
                              tags=(self._item_tag(item_glyph == self._CHK_ON),), open=True)
            self._nodes[node_id] = {"kind": "item", "item_id": iid, "item": item}

            item_info = self._asset_info.get(iid, {})
            for ak in sorted(asset_keys):
                aval     = assets.get(ak, {})
                href     = aval.get("href", "")
                atype    = aval.get("type", "")
                ext      = Path(href).suffix if href else ""
                a_area   = asset_area(aval)
                info     = item_info.get(ak)
                sc       = info.get("status")   if info else None
                sz       = info.get("size_bytes") if info else None
                lm       = info.get("last_modified") if info else None
                stxt, tg = _status_label(sc)

                anid = f"asset::{iid}::{ak}"
                self._tree.insert(node_id, "end", iid=anid,
                                  text=f"        {ak}",
                                  values=(self._chk_glyph(anid), a_area, stxt,
                                          ext or atype[:22], _fmt_size(sz), _fmt_date(lm)),
                                  tags=(self._asset_tag(self._is_checked(anid), tg),))
                self._nodes[anid] = {
                    "kind": "asset", "item_id": iid, "asset_key": ak,
                    "href": href, "item": item,
                }

        n = len(sorted_items)
        self._stats_lbl.configure(
            text=f"{n} Item(s)  |  {total_assets} Asset(s)  "
                 f"(Gesamtcollection: {len(self._all_items)} Items)")
        self._toggle_tree_buttons(True)

    def _toggle_tree_buttons(self, on: bool):
        state = "normal" if on else "disabled"
        self._check_btn.config(state=state)
        self._export_json_btn.config(state=state)
        self._export_csv_btn.config(state=state)
        self._export_links_btn.config(state=state)
        self._map_viewer_btn.config(state=state)
        self._expand_btn.config(state=state)
        self._collapse_btn.config(state=state)
        self._select_all_btn.config(state=state)
        self._deselect_all_btn.config(state=state)

    def _expand_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=True)

    def _collapse_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=False)

    # ── Export-Auswahl (Checkboxen) ───────────────────────────────────────────

    def _is_checked(self, asset_nid: str) -> bool:
        # Default: abgewählt – Nutzer wählt Assets/Items bewusst für Export/
        # Kartenviewer aus, statt sie aktiv abzuwählen.
        return self._checked.get(asset_nid, False)

    def _chk_glyph(self, asset_nid: str) -> str:
        return self._CHK_ON if self._is_checked(asset_nid) else self._CHK_OFF

    def _item_asset_nids(self, item_id: str) -> List[str]:
        return [nid for nid, d in self._nodes.items()
                if d["kind"] == "asset" and d["item_id"] == item_id]

    def _item_check_glyph(self, asset_nids: List[str]) -> str:
        if not asset_nids:
            return self._CHK_OFF
        states = [self._is_checked(n) for n in asset_nids]
        if all(states):
            return self._CHK_ON
        if not any(states):
            return self._CHK_OFF
        return self._CHK_PARTIAL

    def _asset_status_tag(self, asset_nid: str) -> str:
        """HTTP-Prüfstatus-Tag eines Assets, unabhängig vom Auswahlstatus."""
        d    = self._nodes.get(asset_nid, {})
        info = self._asset_info.get(d.get("item_id"), {}).get(d.get("asset_key"))
        sc   = info.get("status") if info else None
        _, tag = _status_label(sc)
        return tag

    def _asset_tag(self, checked: bool, status_tag: str) -> str:
        """Zeilen-Tag für ein Asset: amber, wenn ausgewählt und noch nicht
        HTTP-geprüft; sonst der Prüfstatus-Tag unverändert."""
        return "asset_selected" if checked and status_tag == "asset_dim" else status_tag

    def _item_tag(self, all_checked: bool) -> str:
        """Zeilen-Tag für ein Item: amber, wenn alle Assets ausgewählt sind,
        sonst die bisherige Item-Kennfarbe."""
        return "item_selected" if all_checked else "item"

    def _refresh_item_glyph(self, item_id: str):
        item_nid = f"item::{item_id}"
        if not self._tree.exists(item_nid):
            return
        glyph = self._item_check_glyph(self._item_asset_nids(item_id))
        vals  = list(self._tree.item(item_nid, "values"))
        vals[0] = glyph
        self._tree.item(item_nid, values=vals,
                        tags=(self._item_tag(glyph == self._CHK_ON),))

    def _on_tree_click(self, event):
        if self._tree.identify_region(event.x, event.y) != "cell":
            return
        if self._tree.identify_column(event.x) != "#1":  # "sel"-Spalte
            return
        row = self._tree.identify_row(event.y)
        d = self._nodes.get(row)
        if not d:
            return

        if d["kind"] == "asset":
            self._checked[row] = not self._is_checked(row)
            vals = list(self._tree.item(row, "values"))
            vals[0] = self._chk_glyph(row)
            row_tag = self._asset_tag(self._checked[row], self._asset_status_tag(row))
            self._tree.item(row, values=vals, tags=(row_tag,))
            self._refresh_item_glyph(d["item_id"])
        else:  # item: alle zugehörigen Assets gemeinsam (de)selektieren
            asset_nids = self._item_asset_nids(d["item_id"])
            new_state  = self._item_check_glyph(asset_nids) != self._CHK_ON
            for nid in asset_nids:
                self._checked[nid] = new_state
                vals = list(self._tree.item(nid, "values"))
                vals[0] = self._chk_glyph(nid)
                row_tag = self._asset_tag(new_state, self._asset_status_tag(nid))
                self._tree.item(nid, values=vals, tags=(row_tag,))
            self._refresh_item_glyph(d["item_id"])
        return "break"

    def _select_all(self):
        self._set_all_checked(True)

    def _deselect_all(self):
        self._set_all_checked(False)

    def _set_all_checked(self, state: bool):
        for nid, d in self._nodes.items():
            if d["kind"] != "asset":
                continue
            self._checked[nid] = state
            if self._tree.exists(nid):
                vals = list(self._tree.item(nid, "values"))
                vals[0] = self._chk_glyph(nid)
                row_tag = self._asset_tag(state, self._asset_status_tag(nid))
                self._tree.item(nid, values=vals, tags=(row_tag,))
        for nid, d in self._nodes.items():
            if d["kind"] == "item":
                self._refresh_item_glyph(d["item_id"])

    # ── HEAD-Prüfung ──────────────────────────────────────────────────────────

    def _check_assets(self):
        tasks = [
            (d["item_id"], d["asset_key"], d["href"])
            for nid, d in self._nodes.items()
            if d["kind"] == "asset" and d.get("href") and self._is_checked(nid)
        ]
        if not tasks:
            self._log_write("[Prüfung] Keine ausgewählten Assets mit URL.\n")
            return

        self._check_btn.config(state="disabled")
        self._log_write(f"[Prüfung] {len(tasks)} ausgewählte Assets …\n")

        # Spinner setzen
        for iid, ak, _ in tasks:
            nid = f"asset::{iid}::{ak}"
            if self._tree.exists(nid):
                cur = self._tree.item(nid, "values")
                self._tree.item(nid, values=(cur[0], cur[1], "⟳", cur[3], "–", "–"),
                                tags=("asset_dim",))

        threading.Thread(target=self._worker_check, args=(tasks,), daemon=True).start()

    def _worker_check(self, tasks: List[Tuple[str, str, str]]):
        ok_cnt  = err_cnt = 0
        tot_sz  = 0

        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            future_map = {
                pool.submit(check_asset_info, href, self._auth): (iid, ak)
                for iid, ak, href in tasks
            }
            for fut in concurrent.futures.as_completed(future_map):
                iid, ak = future_map[fut]
                try:
                    info = fut.result()
                except Exception:
                    info = {"status": -3, "size_bytes": None, "last_modified": None}

                self._asset_info.setdefault(iid, {})[ak] = info

                sc       = info.get("status")
                sz       = info.get("size_bytes")
                lm       = info.get("last_modified")
                stxt, tg = _status_label(sc)

                if sc == 200:
                    ok_cnt += 1
                    tot_sz += sz or 0
                elif sc is not None:
                    err_cnt += 1

                nid = f"asset::{iid}::{ak}"
                cur_sel  = self._chk_glyph(nid)
                cur_typ  = ""
                cur_area = ""
                if self._tree.exists(nid):
                    cur_vals = self._tree.item(nid, "values")
                    cur_area, cur_typ = cur_vals[1], cur_vals[3]

                self.after(0, lambda n=nid, sel=cur_sel, s=stxt, t=cur_typ, ar=cur_area,
                           sz_=_fmt_size(sz), lm_=_fmt_date(lm), tag=tg:
                           self._tree.exists(n) and
                           self._tree.item(n, values=(sel, ar, s, t, sz_, lm_), tags=(tag,)))

                self._log_write(f"  {ak}  →  {stxt}  {_fmt_size(sz)}\n")

        self._log_write(
            f"[Prüfung] Fertig: ✓ {ok_cnt}  ✗ {err_cnt}  "
            f"|  Gesamtgrösse (200 OK): {_fmt_size(tot_sz)}\n")
        self.after(0, lambda: self._check_btn.config(state="normal"))
        self.after(0, lambda: self._refresh_stats(ok_cnt, err_cnt, tot_sz))

    def _refresh_stats(self, ok: int, err: int, total_bytes: int):
        n_items  = len(self._visible_items)
        n_assets = sum(len(v) for v in self._asset_info.values())
        self._stats_lbl.configure(
            text=(f"{n_items} Item(s)  |  {n_assets} Asset(s) geprüft  |  "
                  f"✓ {ok} OK   ✗ {err} Fehler  |  "
                  f"Gesamtgrösse: {_fmt_size(total_bytes)}"))

    # ── Kontextmenü / Doppelklick ─────────────────────────────────────────────

    def _on_right_click(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        self._tree.selection_set(row)
        d = self._nodes.get(row, {})
        self._ctx.delete(0, "end")

        if d.get("kind") == "asset":
            href = d.get("href", "")
            iid  = d.get("item_id", "")
            if href:
                self._ctx.add_command(
                    label="URL kopieren",
                    command=lambda h=href: self._clip(h))
                self._ctx.add_command(
                    label="Im Browser öffnen",
                    command=lambda h=href: webbrowser.open(h))
                self._ctx.add_separator()
            self._ctx.add_command(
                label="Item-ID kopieren",
                command=lambda i=iid: self._clip(i))

        if d.get("kind") in ("asset", "item"):
            item = d.get("item")
            if item:
                self._ctx.add_command(
                    label="Item-JSON anzeigen",
                    command=lambda it=item: ItemJsonDialog(self, it, self._dark))
                self._ctx.add_command(
                    label="Im STAC Browser öffnen",
                    command=lambda i=d.get("item_id"): self._open_stac_browser(i))

        self._ctx.tk_popup(event.x_root, event.y_root)

    def _on_double_click(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        d = self._nodes.get(row, {})
        if d.get("kind") == "item":
            self._open_stac_browser(d.get("item_id"))
        elif d.get("kind") == "asset":
            href = d.get("href", "")
            if href:
                webbrowser.open(href)
        # Verhindert das Standard-Auf-/Zuklappen der Treeview bei Doppelklick
        # (Ein-/Ausklappen soll nur über den Button oder das Dreieck erfolgen)
        return "break"

    def _clip(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self._log_write(f"[Clipboard] {text}\n")

    # ── Export ────────────────────────────────────────────────────────────────

    def _export_json(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return

        exts       = self._active_extensions()
        items_out  = []
        asset_count = 0
        for item in self._visible_items:
            iid    = item["id"]
            assets = item.get("assets", {})
            assets_out: Dict = {}
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if exts and not any(href.lower().endswith(e) or ak.lower().endswith(e)
                                    for e in exts):
                    continue
                if not self._is_checked(f"asset::{iid}::{ak}"):
                    continue
                assets_out[ak] = aval
            if not assets_out:
                continue
            items_out.append(build_stac_item(item, assets_out))
            asset_count += len(assets_out)

        # STAC-ItemCollection: Standardformat für einen Export mehrerer valider
        # STAC-1.0.0-Items (analog zur Struktur, die auch die STAC-API selbst
        # bei /items bzw. /search zurückgibt).
        output = {
            "type":     "FeatureCollection",
            "features": items_out,
        }
        content = json.dumps(output, indent=2, ensure_ascii=False)

        def _on_saved(path):
            self._log_write(f"[Export] JSON: {path}\n")
            messagebox.showinfo("Export erfolgreich",
                                f"{len(items_out)} Items  |  "
                                f"{asset_count} Assets\n{path}")

        ExportPreviewDialog(
            self, self._dark, "Download-Links exportieren (JSON)", content,
            initialfile=f"stac_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            filetypes=[("JSON", "*.json"), ("Alle Dateien", "*.*")],
            defaultextension=".json", on_saved=_on_saved,
        )

    def _export_csv(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return

        exts = self._active_extensions()
        rows = []
        for item in self._visible_items:
            iid    = item["id"]
            year   = stac_item_year(item)
            area   = stac_item_area(item)
            acq    = stac_item_acq_date(item)
            assets = item.get("assets", {})
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if exts and not any(href.lower().endswith(e) or ak.lower().endswith(e)
                                    for e in exts):
                    continue
                if not self._is_checked(f"asset::{iid}::{ak}"):
                    continue
                info = self._asset_info.get(iid, {}).get(ak, {})
                rows.append({
                    "item_id":       iid,
                    "year":          year,
                    "area":          area,
                    "acq_date":      acq,
                    "asset_key":     ak,
                    "extension":     Path(href).suffix if href else "",
                    "media_type":    aval.get("type", ""),
                    "http_status":   info.get("status", ""),
                    "size_bytes":    info.get("size_bytes", ""),
                    "size_human":    _fmt_size(info.get("size_bytes")),
                    "last_modified": _fmt_date(info.get("last_modified")),
                    "href":          href,
                })
        if not rows:
            messagebox.showwarning("Keine Daten", "Keine Assets nach Filter.")
            return

        sio = io.StringIO(newline="")
        writer = csv.DictWriter(sio, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
        content = sio.getvalue()

        def _on_saved(path):
            self._log_write(f"[Export] CSV: {path}\n")
            messagebox.showinfo("Export erfolgreich",
                                f"{len(rows)} Zeilen exportiert.\n{path}")

        ExportPreviewDialog(
            self, self._dark, "Export CSV", content,
            initialfile=f"stac_monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv"), ("Alle Dateien", "*.*")],
            defaultextension=".csv", encoding="utf-8-sig",
            write_newline="", on_saved=_on_saved,
        )

    def _export_stac_browser_links(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return

        env    = self._env_var.get()
        exts   = self._active_extensions()
        _pfx   = COLLECTION_ID + "_"
        blocks = []
        for item in self._visible_items:
            iid     = item["id"]
            display = iid[len(_pfx):] if iid.startswith(_pfx) else iid
            assets  = item.get("assets", {})
            asset_entries = []
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if exts and not any(href.lower().endswith(e) or ak.lower().endswith(e)
                                    for e in exts):
                    continue
                if not self._is_checked(f"asset::{iid}::{ak}"):
                    continue
                asset_entries.append((ak, href))
            if not asset_entries:
                continue
            asset_entries.sort(key=lambda e: e[0])
            lines = [
                f"item: {display};",
                f"- {browser_url(env, iid, include_lang=False)}",
                "asset: ",
            ]
            for ak, href in asset_entries:
                lines.append(ak)
                lines.append(f"- {href}")
            blocks.append("\n".join(lines))

        if not blocks:
            messagebox.showwarning("Keine Auswahl", "Keine ausgewählten Assets nach Filter.")
            return

        content = "\n\n\n".join(blocks) + "\n"

        def _on_saved(path):
            self._log_write(f"[Export] STAC-Browser-Links: {path}\n")
            messagebox.showinfo("Export erfolgreich",
                                f"{len(blocks)} Item(s) exportiert.\n{path}")

        ExportPreviewDialog(
            self, self._dark, "Item - STAC Browser Links exportieren", content,
            initialfile=f"item_STAC-Browser-Links_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.txt",
            filetypes=[("Textdatei", "*.txt"), ("Alle Dateien", "*.*")],
            defaultextension=".txt", on_saved=_on_saved,
        )

    def _open_map_viewer(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return

        exts    = self._active_extensions()
        hrefs   = []
        skipped = 0
        for item in self._visible_items:
            iid    = item["id"]
            assets = item.get("assets", {})
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if exts and not any(href.lower().endswith(e) or ak.lower().endswith(e)
                                    for e in exts):
                    continue
                if not self._is_checked(f"asset::{iid}::{ak}"):
                    continue
                if not is_cog_asset(href):
                    skipped += 1
                    continue
                hrefs.append(href)

        if not hrefs:
            messagebox.showwarning(
                "Keine COG-Assets",
                "Keine ausgewählten Assets sind GeoTIFF (.tif/.tiff) und damit "
                "als Layer im Kartenviewer darstellbar.")
            return

        url = map_viewer_url(hrefs)
        webbrowser.open(url)
        hinweis = f"  ({skipped} nicht-COG Asset(s) übersprungen)" if skipped else ""
        self._log_write(f"[Kartenviewer] {len(hrefs)} Layer geöffnet{hinweis}\n{url}\n")

    # ── Logging ───────────────────────────────────────────────────────────────

    def _log_write(self, text: str):
        def _do():
            self._log.config(state="normal")
            self._log.insert("end", text)
            self._log.see("end")
            self._log.config(state="disabled")
        self.after(0, _do)


# ─── Einstiegspunkt ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = StacMonitorApp()
    app.mainloop()
