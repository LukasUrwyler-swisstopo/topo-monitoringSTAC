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
import importlib
import io
import json
import os
import site
import subprocess
import sys
import sysconfig
import tempfile
import threading
import time
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
    build_stac_item, is_cog_asset, is_ebo_ebn_asset, ebo_ebn_kml_item_id,
    is_thumbnail_asset, map_viewer_url, embed_viewer_url, union_bbox,
)

# Firmenproxy für pip, falls die direkte Verbindung zu PyPI im Bundesnetz
# fehlschlägt (analog zum Proxy-Fallback in stac_api.py).
_PIP_PROXY = "http://proxy-bvcol.admin.ch:8080"


def _ensure_win32_modules():
    """Stellt win32con/win32gui/win32process bereit (für das angedockte
    Viewer-Fenster). Fehlt pywin32, wird es beim ersten Start automatisch per
    pip nachinstalliert – kein manuelles "pip install pywin32" nötig. Ein
    reines "try import, pip install, import" reicht dafür nicht: pywin32 hängt
    seine Unterordner über eine .pth-Datei in sys.path, die von site.py nur
    beim Interpreter-Start verarbeitet wird. site.addsitedir() stösst diese
    Verarbeitung zur Laufzeit erneut an, sodass der Import ohne Neustart des
    Tools klappt."""
    try:
        import win32con, win32gui, win32process
        return win32con, win32gui, win32process, True
    except ImportError:
        pass

    for extra_args in ([], ["--proxy", _PIP_PROXY]):
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--quiet", "pywin32"] + extra_args,
                timeout=120)
            for p in {sysconfig.get_path("purelib"), sysconfig.get_path("platlib")}:
                site.addsitedir(p)
            importlib.invalidate_caches()
            import win32con, win32gui, win32process
            return win32con, win32gui, win32process, True
        except Exception:
            continue
    return None, None, None, False


win32con, win32gui, win32process, _WIN32_AVAILABLE = _ensure_win32_modules()

# App-Modus-fähige Chromium-Browser fürs Viewer-Fenster: Chrome bevorzugt,
# Edge als Fallback (Firmenumgebung hat oft nur Edge, kein separates Chrome).
# Beide unterstützen dieselben --app/--window-size/--window-position-Flags
# und (Chromium-Basis) dieselbe Fensterklasse "Chrome_WidgetWin_1".
_BROWSER_CANDIDATES = (
    ("chrome.exe", "Google\\Chrome", (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    )),
    ("msedge.exe", "Microsoft\\Edge", (
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    )),
)


def _find_browser_exe() -> Optional[str]:
    """Sucht eine app-modus-fähige Browser.exe (Chrome vor Edge). Prüft
    Standardpfade, das lokale AppData-Verzeichnis (Pro-User-Installation)
    sowie die Windows "App Paths"-Registry."""
    for exe_name, vendor_subdir, candidates in _BROWSER_CANDIDATES:
        for p in candidates:
            if Path(p).is_file():
                return p
        local = os.environ.get("LOCALAPPDATA")
        if local:
            p = str(Path(local) / vendor_subdir / "Application" / exe_name)
            if Path(p).is_file():
                return p
        try:
            import winreg
            key = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as k:
                p = winreg.QueryValueEx(k, "")[0]
                if p and Path(p).is_file():
                    return p
        except OSError:
            pass
    return None


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
    _COL_W     = {"sel": 90, "area": 180, "status": 100, "typ": 90,
                  "groesse": 90, "geaendert": 105}


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
        # Wird nach dem ersten "Assets prüfen (HEAD)"-Lauf True -> schaltet den
        # "Nur Fehler-Assets anzeigen"-Filter frei.
        self._assets_checked_once: bool = False

        # Lade-Spinner im "ITEM-Liste laden"-Button
        self._spinner_job: Optional[str] = None
        self._spinner_idx: int = 0

        # Angedocktes Viewer-Fenster (Chrome im App-Modus, rechts neben dem
        # Hauptfenster) – Prozess/Fensterhandle des aktuell offenen Viewers.
        self._viewer_proc: Optional[subprocess.Popen] = None
        self._viewer_hwnd: Optional[int] = None
        self._viewer_shown_key: Optional[Tuple[str, str]] = None
        self._viewer_profile_dir: Optional[str] = None
        self._reposition_job: Optional[str] = None

        self._build_ui()
        self._apply_theme(True)

        self.bind("<Configure>", self._on_main_window_configure)
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)

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

        row1 = ttk.Frame(sec)
        row1.pack(side="top", anchor="w")

        ttk.Label(row1, text="Umgebung:").pack(side="left", padx=(0, 6))
        self._env_var = tk.StringVar(value="INT")
        for env in ("INT", "PROD"):
            ttk.Radiobutton(row1, text=env, variable=self._env_var, value=env,
                            command=self._on_env_change).pack(side="left", padx=4)

        self._url_lbl = ttk.Label(row1, text=ENVIRONMENTS["INT"],
                                   font=("Segoe UI", 8), style="Dim.TLabel")
        self._url_lbl.pack(side="left", padx=12)

        ttk.Button(row1, text="STAC Browser öffnen",
                   command=self._open_stac_browser).pack(side="left")

        row2 = ttk.Frame(sec)
        row2.pack(side="top", anchor="w", pady=(6, 0))

        self._cred_btn = ttk.Button(row2, text="Credentials laden",
                                     command=self._load_credentials, style="Amber.TButton")
        self._cred_btn.pack(side="left", padx=(0, 6))

        self._cred_lbl = ttk.Label(row2, text="nicht geladen",
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
        ttk.Label(ext_frame, text="Freitext im Dateinamen [optional]:").pack(side="left", padx=(6, 4))
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

    def _build_stac_functions(self, parent):
        sec = ttk.LabelFrame(parent, text="STAC-Funktionen",
                             padding=8, style="Section.TLabelframe")
        sec.pack(fill="x", pady=(0, 4))

        row1 = ttk.Frame(sec)
        row1.pack(side="top", anchor="w")

        self._check_btn = ttk.Button(
            row1, text="Assets prüfen  (HEAD)", command=self._check_assets, state="disabled")
        self._check_btn.pack(side="left")

        row2 = ttk.Frame(sec)
        row2.pack(side="top", anchor="w", pady=(6, 0))

        self._export_json_btn = ttk.Button(
            row2, text="Export JSON",
            command=self._export_json, state="disabled")
        self._export_json_btn.pack(side="left", padx=(0, 4))

        self._export_csv_btn = ttk.Button(
            row2, text="Export CSV", command=self._export_csv, state="disabled")
        self._export_csv_btn.pack(side="left", padx=(0, 4))

        self._export_links_btn = ttk.Button(
            row2, text="Export STAC Browser Links",
            command=self._export_stac_browser_links, state="disabled")
        self._export_links_btn.pack(side="left", padx=(0, 4))

        row3 = ttk.Frame(sec)
        row3.pack(side="top", anchor="w", pady=(6, 0))

        self._map_viewer_btn = ttk.Button(
            row3, text="Link auf Kartenviewer",
            command=self._open_map_viewer, state="disabled")
        self._map_viewer_btn.pack(side="left", padx=(0, 4))

        self._viewer_win_btn = ttk.Button(
            row3, text="GUI Viewer öffnen",
            command=self._open_viewer_window, state="disabled")
        self._viewer_win_btn.pack(side="left")

    def _build_tree(self, parent):
        frame = ttk.LabelFrame(parent, text="3   Items & Assets",
                               padding=4, style="Section.TLabelframe")
        frame.pack(fill="both", expand=True, pady=(0, 4))
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(frame)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        select_row = ttk.Frame(toolbar)
        select_row.pack(side="top", anchor="w")

        _btn_w = 16

        self._select_all_btn = ttk.Button(
            select_row, text="Alle auswählen", command=self._select_all,
            state="disabled", width=_btn_w)
        self._select_all_btn.pack(side="left", padx=(0, 4))

        self._deselect_all_btn = ttk.Button(
            select_row, text="Alles abwählen", command=self._deselect_all,
            state="disabled", width=_btn_w)
        self._deselect_all_btn.pack(side="left")

        expand_row = ttk.Frame(toolbar)
        expand_row.pack(side="top", anchor="w", pady=(4, 0))

        self._expand_btn = ttk.Button(
            expand_row, text="Alle aufklappen", command=self._expand_all,
            state="disabled", width=_btn_w)
        self._expand_btn.pack(side="left", padx=(0, 4))

        self._collapse_btn = ttk.Button(
            expand_row, text="Alle einklappen", command=self._collapse_all,
            state="disabled", width=_btn_w)
        self._collapse_btn.pack(side="left")

        error_row = ttk.Frame(toolbar)
        error_row.pack(side="top", anchor="w", pady=(4, 0))

        self._error_filter_var = tk.BooleanVar(value=False)
        self._error_filter_btn = ttk.Checkbutton(
            error_row, text="Assets mit ERRORs anzeigen",
            variable=self._error_filter_var, command=self._on_error_filter_toggle,
            state="disabled")
        self._error_filter_btn.pack(side="left")

        self._tree = ttk.Treeview(
            frame, columns=self._COLS, show="tree headings", selectmode="browse")

        self._tree.column("#0", width=620, minwidth=500, stretch=False)
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
        self._assets_checked_once = False
        self._error_filter_var.set(False)
        self._populate_tree([], [], [], False)  # Bestehende Liste sofort leeren, bevor neu geladen wird
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
        return result

    def _active_terms(self) -> List[str]:
        return [p.lower() for p in self._ext_custom_var.get().replace(",", " ").split()]

    @staticmethod
    def _asset_matches(href: str, key: str, exts: List[str], terms: List[str]) -> bool:
        href_l = href.lower()
        key_l  = key.lower()
        if exts and not any(href_l.endswith(e) or key_l.endswith(e) for e in exts):
            return False
        if terms and not any(t in href_l or t in key_l for t in terms):
            return False
        return True

    def _asset_is_error(self, item_id: str, asset_key: str) -> bool:
        """True, wenn das Asset bereits per HEAD geprüft wurde und dabei
        keinen Status 200 hatte. Noch ungeprüfte Assets gelten nicht als
        Fehler (unbekannt statt fehlerhaft)."""
        info = self._asset_info.get(item_id, {}).get(asset_key)
        return bool(info) and info.get("status") != 200

    def _on_error_filter_toggle(self):
        # Ansicht wechselt (alle Assets <-> nur Fehler) -> bisherige Auswahl
        # verwerfen, damit keine inzwischen unsichtbaren Assets exportiert/
        # geprüft werden.
        self._checked.clear()
        self._apply_filters()

    def _apply_filters(self):
        if not self._all_items:
            return
        self._items_loaded_once = True
        year        = self._year_var.get().strip()
        search      = self._search_var.get().strip()
        exts        = self._active_extensions()
        terms       = self._active_terms()
        errors_only = self._error_filter_var.get()

        items = self._all_items
        if search:
            items = filter_items(items, search)
        if year:
            items = [it for it in items if stac_item_year(it) == year]
        if exts or terms or errors_only:
            def _has_match(it):
                iid = it["id"]
                for k, v in it.get("assets", {}).items():
                    if not self._asset_matches(v.get("href", ""), k, exts, terms):
                        continue
                    if errors_only and not self._asset_is_error(iid, k):
                        continue
                    return True
                return False
            items = [it for it in items if _has_match(it)]

        self._visible_items = items
        self._populate_tree(items, exts, terms, errors_only)

    def _populate_tree(self, items: List[Dict], exts: List[str], terms: List[str],
                        errors_only: bool = False):
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
            if exts or terms or errors_only:
                asset_keys = [
                    k for k, v in assets.items()
                    if self._asset_matches(v.get("href", ""), k, exts, terms)
                    and (not errors_only or self._asset_is_error(iid, k))
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
        self._viewer_win_btn.config(state=state)
        self._expand_btn.config(state=state)
        self._collapse_btn.config(state=state)
        self._select_all_btn.config(state=state)
        self._deselect_all_btn.config(state=state)
        self._error_filter_btn.config(
            state="normal" if (on and self._assets_checked_once) else "disabled")

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
        checked = [
            (d["item_id"], d["asset_key"], d["href"])
            for nid, d in self._nodes.items()
            if d["kind"] == "asset" and d.get("href") and self._is_checked(nid)
        ]
        if not checked:
            self._log_write("[Prüfung] Keine ausgewählten Assets mit URL.\n")
            messagebox.showinfo("Auswahl erforderlich", "Bitte Assets auswählen.")
            return

        # Thumbnails werden übersprungen: viele kleine Zusatzdateien, die die
        # Prüfung stark verlangsamen und für die Datenkontrolle irrelevant sind.
        tasks = [t for t in checked if not is_thumbnail_asset(t[1]) and not is_thumbnail_asset(t[2])]
        n_thumbs = len(checked) - len(tasks)
        if not tasks:
            self._log_write("[Prüfung] Ausgewählte Assets sind ausschliesslich Thumbnails – übersprungen.\n")
            messagebox.showinfo(
                "Nur Thumbnails ausgewählt",
                "Alle ausgewählten Assets sind Thumbnails und werden bei der "
                "Prüfung übersprungen. Bitte andere Assets auswählen.")
            return

        self._check_btn.config(state="disabled")
        hinweis = f"  ({n_thumbs} Thumbnail(s) übersprungen)" if n_thumbs else ""
        self._log_write(f"[Prüfung] {len(tasks)} ausgewählte Assets{hinweis} …\n")

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

                if sc != 200:
                    self._log_write(f"  {ak}  →  {stxt}  {_fmt_size(sz)}\n")

        self._log_write(
            f"[Prüfung] Fertig: ✓ {ok_cnt}  ✗ {err_cnt}  "
            f"|  Gesamtgrösse (200 OK): {_fmt_size(tot_sz)}\n")
        self.after(0, lambda: self._check_btn.config(state="normal"))
        self.after(0, self._enable_error_filter_btn)
        self.after(0, lambda: self._refresh_stats(ok_cnt, err_cnt, tot_sz))

    def _enable_error_filter_btn(self):
        self._assets_checked_once = True
        self._error_filter_btn.config(state="normal")

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
        terms      = self._active_terms()
        items_out  = []
        asset_count = 0
        for item in self._visible_items:
            iid    = item["id"]
            assets = item.get("assets", {})
            assets_out: Dict = {}
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if not self._asset_matches(href, ak, exts, terms):
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

        exts  = self._active_extensions()
        terms = self._active_terms()
        rows = []
        for item in self._visible_items:
            iid    = item["id"]
            year   = stac_item_year(item)
            area   = stac_item_area(item)
            acq    = stac_item_acq_date(item)
            assets = item.get("assets", {})
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if not self._asset_matches(href, ak, exts, terms):
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
        terms  = self._active_terms()
        _pfx   = COLLECTION_ID + "_"
        blocks = []
        for item in self._visible_items:
            iid     = item["id"]
            display = iid[len(_pfx):] if iid.startswith(_pfx) else iid
            assets  = item.get("assets", {})
            asset_entries = []
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if not self._asset_matches(href, ak, exts, terms):
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
        terms   = self._active_terms()
        targets = []
        for item in self._visible_items:
            iid    = item["id"]
            assets = item.get("assets", {})
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if not self._asset_matches(href, ak, exts, terms):
                    continue
                if not self._is_checked(f"asset::{iid}::{ak}"):
                    continue
                if is_cog_asset(href) or is_ebo_ebn_asset(ak) or is_ebo_ebn_asset(href):
                    targets.append((item, ak, href))

        if not targets:
            messagebox.showwarning(
                "Keine darstellbaren Assets",
                "Keine ausgewählten Assets sind als GeoTIFF (.tif/.tiff) oder als "
                "EBO/EBN-Foto (zugehöriges Tages-KML) im Kartenviewer darstellbar.")
            return

        self._log_write(f"[Kartenviewer] Löse Layer für {len(targets)} Asset(s) auf …\n")
        threading.Thread(target=self._worker_open_map_viewer,
                          args=(targets,), daemon=True).start()

    def _worker_open_map_viewer(self, targets: List[Tuple[Dict, str, str]]):
        layers   = []
        bboxes   = []
        seen_kml = set()
        skipped  = 0
        for item, _ak, href in targets:
            if is_cog_asset(href):
                layers.append(f"COG|{href}")
                bboxes.append(item.get("bbox"))
                continue
            kml_href = self._find_ebo_ebn_kml_href(item)
            if kml_href and kml_href not in seen_kml:
                seen_kml.add(kml_href)
                layers.append(f"KML|{kml_href}")
            elif not kml_href:
                skipped += 1

        def _finish():
            if not layers:
                messagebox.showwarning(
                    "Kein Layer gefunden",
                    "Für die ausgewählten Assets konnte kein darstellbarer "
                    "Layer (COG/KML) ermittelt werden.")
                return
            # Nur über die COG-Bboxen zoomen (nicht über die der KML-Tages-
            # übersichten – die Bbox des einzelnen Foto-Items wäre für die
            # Tagesübersicht keine sinnvolle Zoomstufe, s. embed_viewer_url).
            url = map_viewer_url(layers, union_bbox(bboxes))
            webbrowser.open(url)
            hinweis = f"  ({skipped} ohne auflösbares Tages-KML übersprungen)" if skipped else ""
            self._log_write(f"[Kartenviewer] {len(layers)} Layer geöffnet{hinweis}\n{url}\n")
        self.after(0, _finish)

    def _find_ebo_ebn_kml_href(self, item: Dict) -> Optional[str]:
        """Löst zu einem EBO/EBN-Foto-Item das zugehörige Tagesübersicht-Item
        (fixe Zeit 23595900) auf und liefert dessen KML-Asset-Href. Sucht
        zuerst in den bereits geladenen Items, sonst per Direct-Lookup."""
        sibling_id = ebo_ebn_kml_item_id(item.get("id", ""))
        if not sibling_id:
            return None
        sibling = next((it for it in self._all_items if it.get("id") == sibling_id), None)
        if sibling is None and self._auth:
            try:
                sibling = get_item_direct(self._base_url, self._auth, sibling_id)
            except Exception as exc:
                self._log_write(f"[Viewer] KML-Lookup {sibling_id} fehlgeschlagen: {exc}\n")
                return None
        if not sibling:
            return None
        for aval in sibling.get("assets", {}).values():
            href = aval.get("href", "")
            if href.lower().endswith(".kml"):
                return href
        return None

    # ── Angedocktes Viewer-Fenster (Chrome-App-Modus, rechts vom Hauptfenster) ──

    def _checked_asset_nodes(self) -> List[Dict]:
        return [d for nid, d in self._nodes.items()
                if d["kind"] == "asset" and self._is_checked(nid)]

    def _open_viewer_window(self):
        checked = self._checked_asset_nodes()
        if len(checked) != 1:
            messagebox.showinfo(
                "Auswahl erforderlich",
                "Bitte nur ein Asset für die Ansicht im Viewer auswählen. "
                "Bitte treffen Sie eine Auswahl eines Assets.")
            return

        d = checked[0]
        item, ak, href = d["item"], d["asset_key"], d["href"]
        if not (is_cog_asset(href) or is_ebo_ebn_asset(ak) or is_ebo_ebn_asset(href)):
            messagebox.showwarning(
                "Nicht darstellbar",
                f"Das Asset '{ak}' ist weder ein GeoTIFF (.tif/.tiff) noch ein "
                "EBO/EBN-Foto (.jpg) und kann daher nicht im Viewer dargestellt werden.")
            return

        key = (item["id"], ak)

        if self._viewer_proc and self._viewer_proc.poll() is None:
            if key == self._viewer_shown_key:
                self._reposition_viewer_window()
                return
            self._close_viewer_window()

        if is_cog_asset(href):
            self._launch_viewer_window(item, "COG", href, key, fit_to_bbox=True)
            return

        # EBO/EBN-Foto: zugehöriges Tages-KML auflösen – ggf. Netzwerkzugriff,
        # daher im Hintergrundthread, um die GUI nicht zu blockieren.
        self._log_write(f"[Viewer-Fenster] Suche Tages-KML zu {item['id']} / {ak} …\n")

        def _worker():
            kml_href = self._find_ebo_ebn_kml_href(item)
            self.after(0, lambda: self._on_kml_resolved_for_viewer(item, ak, key, kml_href))
        threading.Thread(target=_worker, daemon=True).start()

    def _on_kml_resolved_for_viewer(self, item: Dict, ak: str, key: Tuple[str, str],
                                     kml_href: Optional[str]):
        if not kml_href:
            messagebox.showwarning(
                "Kein KML gefunden",
                f"Für {item['id']} / {ak} konnte kein zugehöriges Tages-KML "
                "gefunden werden.")
            return
        self._launch_viewer_window(item, "KML", kml_href, key, fit_to_bbox=False)

    def _launch_viewer_window(self, item: Dict, layer_type: str, layer_href: str,
                               key: Tuple[str, str], fit_to_bbox: bool):
        if not _WIN32_AVAILABLE:
            messagebox.showerror(
                "Viewer nicht verfügbar",
                "Das Modul 'pywin32' konnte nicht automatisch installiert werden "
                "(kein pip/Internetzugriff?) – ohne dieses Modul kann das "
                "Viewer-Fenster nicht neben dem Hauptfenster positioniert werden.")
            return
        browser = _find_browser_exe()
        if not browser:
            messagebox.showerror(
                "Viewer nicht verfügbar",
                "Weder Google Chrome noch Microsoft Edge wurden auf diesem Rechner "
                "gefunden. Das Viewer-Fenster benötigt einen der beiden Browser.")
            return

        self.update_idletasks()
        x = self.winfo_x() + self.winfo_width()
        y = self.winfo_y()
        w = max(self.winfo_width(), 480)
        h = max(self.winfo_height(), 480)

        url = embed_viewer_url(item, f"{layer_type}|{layer_href}", w, h, fit_to_bbox=fit_to_bbox)

        if self._viewer_profile_dir is None:
            self._viewer_profile_dir = tempfile.mkdtemp(prefix="stac_monitor_viewer_")

        try:
            proc = subprocess.Popen([
                browser, f"--app={url}",
                f"--user-data-dir={self._viewer_profile_dir}",
                f"--window-size={w},{h}", f"--window-position={x},{y}",
                "--no-first-run", "--no-default-browser-check",
            ])
        except OSError as e:
            messagebox.showerror("Viewer nicht verfügbar", f"Browser konnte nicht gestartet werden:\n{e}")
            return

        self._viewer_proc        = proc
        self._viewer_hwnd        = None
        self._viewer_shown_key   = key
        self._log_write(f"[Viewer-Fenster] {item['id']} ({layer_type}) / {layer_href}\n{url}\n")

        threading.Thread(target=self._wait_and_style_viewer_window,
                          args=(proc.pid,), daemon=True).start()

    def _wait_and_style_viewer_window(self, pid: int):
        """Läuft in einem Hintergrundthread: sucht das neu geöffnete Browser-
        Fenster per PID, entfernt Titelleiste/Rahmen für den angedockten Look.
        Reine win32-Aufrufe auf ein fremdes Fensterhandle – keine Tkinter-
        Widget-Zugriffe, daher unkritisch ausserhalb des Hauptthreads.

        Chrome und Edge (beide Chromium-basiert) nutzen normalerweise die
        Fensterklasse "Chrome_WidgetWin_1"; nach 4s ohne Treffer wird die
        Suche zur Sicherheit auf jedes hinreichend grosse PID-Fenster
        erweitert, falls eine Browser-Variante davon abweicht."""
        hwnd_found = None
        deadline    = time.time() + 8
        broaden_at  = time.time() + 4
        while time.time() < deadline and not hwnd_found:
            strict = time.time() < broaden_at

            def _cb(hwnd, _):
                nonlocal hwnd_found
                if hwnd_found or not win32gui.IsWindowVisible(hwnd):
                    return
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                if wpid != pid:
                    return
                if strict:
                    if win32gui.GetClassName(hwnd) == "Chrome_WidgetWin_1":
                        hwnd_found = hwnd
                    return
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                if (right - left) >= 200 and (bottom - top) >= 200:
                    hwnd_found = hwnd
            win32gui.EnumWindows(_cb, None)
            if not hwnd_found:
                time.sleep(0.25)

        if not hwnd_found:
            return

        style = win32gui.GetWindowLong(hwnd_found, win32con.GWL_STYLE)
        style &= ~win32con.WS_CAPTION & ~win32con.WS_THICKFRAME
        win32gui.SetWindowLong(hwnd_found, win32con.GWL_STYLE, style)

        self._viewer_hwnd = hwnd_found
        self.after(0, self._reposition_viewer_window)

    def _reposition_viewer_window(self):
        if not _WIN32_AVAILABLE or not self._viewer_hwnd:
            return
        if not win32gui.IsWindow(self._viewer_hwnd):
            self._viewer_hwnd      = None
            self._viewer_proc      = None
            self._viewer_shown_key = None
            return
        x = self.winfo_x() + self.winfo_width()
        y = self.winfo_y()
        w = max(self.winfo_width(), 480)
        h = max(self.winfo_height(), 480)
        win32gui.SetWindowPos(
            self._viewer_hwnd, None, x, y, w, h,
            win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_FRAMECHANGED)

    def _on_main_window_configure(self, event):
        if event.widget is not self:
            return
        if self._reposition_job is not None:
            self.after_cancel(self._reposition_job)
        self._reposition_job = self.after(80, self._reposition_viewer_window)

    def _close_viewer_window(self):
        if self._viewer_proc and self._viewer_proc.poll() is None:
            try:
                self._viewer_proc.terminate()
            except OSError:
                pass
        self._viewer_proc      = None
        self._viewer_hwnd      = None
        self._viewer_shown_key = None

    def _on_app_close(self):
        self._close_viewer_window()
        self.destroy()

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
