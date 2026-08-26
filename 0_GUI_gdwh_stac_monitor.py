"""
0_GUI_gdwh_stac_monitor.py  –  STAC Monitoring-Tool (read-only)

Zeigt Items und Assets der Collection "ch.swisstopo.spezialbefliegungen"
in einer Baumansicht. Funktionen:
  - Asset-Status-Prüfung via HEAD (HTTP-Code, Dateigrösse, Last-Modified)
  - Statistik: OK / Fehler / Gesamtgrösse
  - Item-JSON Detailansicht (Doppelklick oder Rechtsklick)
  - URL in Zwischenablage kopieren, im Browser öffnen

Credentials: secrets/stac_credentials.json
Format:      {"INT": {"username": "...", "password": "..."}, "PROD": {...}}
"""

import ctypes
import importlib
import json
import os
import re
import site
import subprocess
import sys
import sysconfig
import tempfile
import threading
import time
import concurrent.futures
import webbrowser
from datetime import datetime, timedelta
from email.utils import parsedate
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

from stac_api import (
    COLLECTION_ID, ENVIRONMENTS, AUFTRAGSTYPEN, EXT_PRESETS,
    LARGE_ASSET_THRESHOLD_BYTES,
    get_item_direct, get_collection_items, filter_items,
    check_asset_info, download_asset, browser_url, asset_area,
    stac_item_year, stac_item_area, stac_item_acq_date,
    is_cog_asset, is_ebo_ebn_asset, ebo_ebn_kml_item_id,
    is_thumbnail_asset, map_viewer_url, embed_viewer_url, union_bbox,
)
from gdwh_api import (
    GDWH_GDS_KEYS, GDWH_ENVIRONMENTS,
    gdwh_get_imports, gdwh_import_id, gdwh_import_date,
    gdwh_search_file_metadata, gdwh_index_file_metadata_by_import,
)

# GDWH indexiert den fileMetadata-Suchindex (Area/Jahr/Auftragstyp) zeitversetzt
# zum eigentlichen Import – live verifiziert am 2026-08-20: ein Import von
# vor >1h hatte noch keinen Match, alle älteren Imports desselben Tages
# (ab ca. 4h zuvor) bereits. Innerhalb dieses Fensters ist "kein Match" also
# normal und KEINE Anomalie. Erst danach gilt ein fehlender Match als
# vermutlich verwaister Alt-Eintrag (siehe _gdwh_apply_filter).
_GDWH_PENDING_HOURS = 24


def _gdwh_is_pending(imp: Dict) -> bool:
    """True, wenn ein Import jünger als _GDWH_PENDING_HOURS ist. Für so junge
    Imports ist ein fehlender FileMetadata-Match durch die Indexierungs-
    Verzögerung von GDWH erklärt, nicht durch eine echte Anomalie."""
    try:
        imp_dt = datetime.strptime(gdwh_import_date(imp), "%Y-%m-%d %H:%M")
    except ValueError:
        return False
    return datetime.now() - imp_dt < timedelta(hours=_GDWH_PENDING_HOURS)

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
    if sc == -4:
        return "✓  >50GB", "asset_ok"
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

    _SHOW_FAULTY_BTN_LABEL   = "Fehlerhafte anzeigen"
    _SHOW_NO_THUMB_BTN_LABEL = "ITEMs ohne Thumbnail"
    _SHOW_ALL_BTN_LABEL      = "Alle Assets wieder anzeigen"

    _SELECT_ALL_BTN_LABEL   = "Alles auswählen"
    _DESELECT_ALL_BTN_LABEL = "Alles abwählen"
    _COLLAPSE_ALL_BTN_LABEL = "Alle einklappen"
    _EXPAND_ALL_BTN_LABEL   = "Alle aufklappen"
    # Items mit dieser Zeichenfolge im Namen (Tagesübersicht-Items mit
    # KML-Platzhalter, feste Zeit 23:59:59 – siehe stac_api._KML_DAILY_SUFFIX)
    # haben planmässig nie ein Thumbnail und sind im "ITEMs ohne Thumbnail"-
    # Filter deshalb keine echten Kandidaten.
    _NO_THUMB_EXCLUDE_SUBSTR = "23595900"

    def __init__(self):
        super().__init__()
        self.title("STAC Monitor  —  ch.swisstopo.spezialbefliegungen")
        self.minsize(1040, 720)
        # Standardgrösse bewusst höher als minsize: die Item/Asset-Liste (Kernstück
        # der App) soll ohne manuelles Vergrössern schon genug Zeilen zeigen. Breite
        # bleibt knapp, damit rechts weiterhin Platz für das angedockte Viewer-Fenster
        # bleibt (siehe _reposition_viewer_window).
        try:
            screen_h = self.winfo_screenheight()
            win_h = max(760, min(screen_h - 80, 1000))
            self.geometry(f"1040x{win_h}")
        except tk.TclError:
            pass

        self._dark: bool = True
        self._auth: Optional[Tuple] = None
        self._base_url: str = ""

        self._all_items: List[Dict] = []
        self._visible_items: List[Dict] = []
        self._items_loaded_once: bool = False
        # Items werden beim (Neu-)Aufbau des Baums immer aufgeklappt eingefügt
        # (siehe _populate_tree) – der Toggle-Button startet daher synchron
        # dazu im Zustand "expanded".
        self._all_expanded: bool = True

        # Baum-Metadaten: tree_iid → dict mit kind/item_id/asset_key/href/item
        self._nodes: Dict[str, Dict] = {}
        # Prüfergebnisse: {item_id: {asset_key: {status, size_bytes, last_modified}}}
        self._asset_info: Dict[str, Dict[str, Dict]] = {}
        # Export-Auswahl je Asset-Knoten (tree_iid → bool). Fehlender Eintrag = gewählt.
        self._checked: Dict[str, bool] = {}
        # Wird nach dem ersten "Assets prüfen (HEAD)"-Lauf True -> schaltet den
        # "Nur Fehler-Assets anzeigen"-Filter frei.
        self._assets_checked_once: bool = False
        # Toggle für "ITEMs ohne Thumbnail" (nur bei Auftragstyp RAM sichtbar,
        # analog zum STAC/GDWH Deleting-Tool)
        self._show_no_thumb_only: bool = False

        # Lade-Spinner im "ITEM-Liste laden"-Button
        self._spinner_job: Optional[str] = None
        self._spinner_idx: int = 0

        # Generische Spinner für weitere Buttons (Key: Button-Widget, Value:
        # dict mit idx/after_id/label) – unabhängig vom Lade-Spinner oben,
        # da mehrere Buttons gleichzeitig "busy" sein können.
        self._btn_spinners: Dict[ttk.Button, Dict] = {}

        # Angedocktes Viewer-Fenster (Chrome im App-Modus, rechts neben dem
        # Hauptfenster) – Prozess/Fensterhandle des aktuell offenen Viewers.
        self._viewer_proc: Optional[subprocess.Popen] = None
        self._viewer_hwnd: Optional[int] = None
        self._viewer_shown_key: Optional[Tuple[str, str]] = None
        self._viewer_profile_dir: Optional[str] = None
        self._reposition_job: Optional[str] = None

        # GDWH-Tab: rohe Imports + angereicherte (Import, FileMetadata-Match, GDS-Key)
        # Tripel der aktuell geladenen Umgebung/GDS-Key-Auswahl (ein oder alle Keys).
        self._gdwh_enriched: List[Tuple[Dict, Optional[Dict], str]] = []
        self._gdwh_total_leichen: int = 0
        self._gdwh_current_gds_key: str = ""
        self._gdwh_base_url: str = ""

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

        # Log-Frame VOR dem Notebook packen (side="bottom"): so reserviert das
        # Pack-Layout dessen feste Höhe zuerst, statt dass das Notebook
        # (expand=True) bei fixer Fensterhöhe den gesamten Rest beansprucht
        # und den Log-Bereich verdrängt.
        log_parent = ttk.Frame(self)
        log_parent.pack(side="bottom", fill="x", padx=12, pady=(0, 8))
        self._build_log(log_parent)

        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        stac_tab = ttk.Frame(self._nb)
        self._nb.add(stac_tab, text="STAC")
        self._build_credentials(stac_tab)
        self._build_filters(stac_tab)
        # STAC-Funktionen + Stats-Zeile unterhalb der Liste, aber (wie der
        # Log-Frame oben) VOR der expandierenden Liste gepackt (side="bottom"):
        # so reservieren sie zuerst ihre feste Höhe von unten, statt dass die
        # Liste (expand=True) bei fixer Fensterhöhe den gesamten Rest
        # beansprucht und sie verdrängt. Pack-Reihenfolge unter mehreren
        # side="bottom"-Widgets bestimmt die visuelle Reihenfolge von unten
        # nach oben – zuerst gepackt landet ganz unten.
        self._build_stac_functions(stac_tab)
        self._build_stats(stac_tab)
        self._build_tree(stac_tab)

        gdwh_tab = ttk.Frame(self._nb)
        self._nb.add(gdwh_tab, text="GDWH")
        self._build_gdwh_tab(gdwh_tab)

    def _optional_label(self, parent, text):
        """Label mit einem '[...]'-Zusatz (z.B. '[optional]') kursiv und
        etwas kleiner als der Haupttext – ttk.Label kann keinen gemischten
        Font in einem Widget darstellen, daher zwei Labels in einem Frame."""
        base, _, bracket = text.partition("[")
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=base).pack(side="left")
        if bracket:
            ttk.Label(frame, text="[" + bracket,
                      font=("Segoe UI", 8, "italic")).pack(side="left")
        return frame

    def _build_credentials(self, parent):
        sec = ttk.LabelFrame(parent, text="1   Umgebung & Credentials",
                             padding=8, style="Section.TLabelframe")
        sec.pack(fill="x", pady=(0, 4))

        # Eine einzige Zeile statt zwei: die URL muss nicht separat ausge-
        # schrieben stehen (der "STAC Browser öffnen"-Button verlinkt ohnehin
        # dorthin) – der frei gewordene Platz nimmt stattdessen direkt den
        # "Credentials laden"-Button samt Status auf.
        row1 = ttk.Frame(sec)
        row1.pack(side="top", anchor="w")

        ttk.Label(row1, text="Umgebung:").pack(side="left", padx=(0, 6))
        self._env_var = tk.StringVar(value="INT")
        for env in ("INT", "PROD"):
            ttk.Radiobutton(row1, text=env, variable=self._env_var, value=env,
                            command=self._on_env_change).pack(side="left", padx=4)

        self._cred_btn = ttk.Button(row1, text="Credentials laden",
                                     command=self._load_credentials, style="Amber.TButton")
        self._cred_btn.pack(side="left", padx=(12, 6))

        self._cred_lbl = ttk.Label(row1, text="nicht geladen",
                                    font=("Segoe UI", 9, "italic"), style="Dim.TLabel")
        self._cred_lbl.pack(side="left", padx=(0, 16))

        ttk.Button(row1, text="STAC Browser öffnen",
                   command=self._open_stac_browser).pack(side="left")

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

        # Jahr + Area
        self._optional_label(sec, "Jahr [optional]:").grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        self._year_var = tk.StringVar()
        self._year_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._year_var, width=8).grid(
            row=1, column=1, sticky="w", pady=(6, 0))

        self._optional_label(sec, "Area [optional]:").grid(
            row=1, column=2, sticky="w", padx=(16, 6), pady=(6, 0))
        self._area_var = tk.StringVar()
        self._area_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._area_var, width=20).grid(
            row=1, column=3, sticky="w", pady=(6, 0))

        # Item-ID / Suche (eigene Zeile)
        self._optional_label(sec, "Item-ID / Suche [optional]:").grid(
            row=2, column=0, sticky="w", pady=(6, 0))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._search_var, width=34).grid(
            row=2, column=1, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(sec, text="Teilstring genügt  (für direkten Abruf: vollständige ID)",
                  font=("Segoe UI", 8), style="Dim.TLabel").grid(
            row=2, column=3, sticky="w", padx=(8, 0), pady=(6, 0))

        # Suchfeld gleich mit dem Default-Auftragstyp vorbefüllen (Radiobutton-Command
        # feuert sonst erst bei einem tatsächlichen Klick, nicht bei der Vorauswahl).
        self._on_auftragstyp_change()

        # Dateiendung
        ttk.Label(sec, text="Dateiendung:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ext_frame = ttk.Frame(sec)
        ext_frame.grid(row=3, column=1, columnspan=4, sticky="w", pady=(6, 0))
        self._ext_vars: List[Tuple[tk.BooleanVar, List[str]]] = []
        for label, exts in EXT_PRESETS:
            var = tk.BooleanVar(value=False)
            var.trace_add("write", lambda *_: self._apply_filters())
            self._ext_vars.append((var, exts))
            ttk.Checkbutton(ext_frame, text=label, variable=var).pack(side="left", padx=(0, 10))
        self._optional_label(ext_frame, "Freitext im Dateinamen [optional]:").pack(
            side="left", padx=(6, 4))
        self._ext_custom_var = tk.StringVar()
        self._ext_custom_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(ext_frame, textvariable=self._ext_custom_var, width=14).pack(side="left")

    def _build_stac_functions(self, parent):
        # Kompaktes 2-zeiliges Toolbar-Layout statt einer eigenen Zeile pro
        # Funktionsgruppe: gibt der Item/Asset-Liste darunter deutlich mehr
        # vertikalen Platz. Gruppen bleiben durch eine kleine, dezente
        # Beschriftung + Trennlinie erkennbar.
        sec = ttk.LabelFrame(parent, text="STAC-Funktionen",
                             padding=(8, 4), style="Section.TLabelframe")
        sec.pack(side="bottom", fill="x", pady=(4, 0))

        _hdr_font = ("Segoe UI", 7, "bold")

        bar_top = ttk.Frame(sec)
        bar_top.pack(side="top", anchor="w", fill="x")
        bar_bottom = ttk.Frame(sec)
        bar_bottom.pack(side="top", anchor="w", fill="x", pady=(6, 0))

        def _group(bar, caption):
            grp = ttk.Frame(bar)
            grp.pack(side="left", anchor="n", padx=(0, 10))
            ttk.Label(grp, text=caption, font=_hdr_font,
                      style="Dim.TLabel").pack(side="top", anchor="w")
            row = ttk.Frame(grp)
            row.pack(side="top", anchor="w", pady=(1, 0))
            return row

        def _sep(bar):
            ttk.Separator(bar, orient="vertical").pack(
                side="left", fill="y", padx=(0, 10))

        # ── Zeile 1: Quality Check | Export ─────────────────────────────────
        row1 = _group(bar_top, "Quality Check")

        self._check_btn = ttk.Button(
            row1, text="Assets prüfen  (HEAD)", command=self._check_assets,
            state="disabled")
        self._check_btn.pack(side="left")

        self._show_faulty_btn = ttk.Button(
            row1, text=self._SHOW_FAULTY_BTN_LABEL,
            command=self._toggle_faulty_filter, state="disabled")
        self._show_faulty_btn.pack(side="left", padx=(4, 0))
        # Text-Sync mit self._error_filter_var (existiert erst nach
        # _build_tree()) wird dort verdrahtet, siehe _build_tree().

        # Nur bei Auftragstyp RAM relevant (Thumbnail-Pflicht) – wird erst
        # sichtbar gepackt, wenn AUFTRAGSTYPEN[...] == "ram" ist, siehe
        # _update_no_thumb_btn_visibility().
        self._show_no_thumb_btn = ttk.Button(
            row1, text=self._SHOW_NO_THUMB_BTN_LABEL,
            command=self._toggle_no_thumb_filter, state="disabled")
        self._update_no_thumb_btn_visibility()

        _sep(bar_top)
        row2 = _group(bar_top, "Export")

        self._export_links_btn = ttk.Button(
            row2, text="Export STAC Browser Links",
            command=self._export_stac_browser_links, state="disabled")
        self._export_links_btn.pack(side="left", padx=(0, 4))

        self._create_links_btn = ttk.Button(
            row2, text="create Download-Links",
            command=self._create_download_links, state="disabled")
        self._create_links_btn.pack(side="left")

        # ── Zeile 2: Download | ASSET Viewer ────────────────────────────────
        row_dl = _group(bar_bottom, "Download")

        self._download_btn = ttk.Button(
            row_dl, text="Download ausgewählte ITEMs/ASSETs",
            command=self._download_assets, state="disabled")
        self._download_btn.pack(side="left")

        _sep(bar_bottom)
        row3 = _group(bar_bottom, "ASSET Viewer")

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

        # "ITEM-Liste laden" steht direkt links neben den Auswahl-/Klapp-
        # Toggle-Buttons über der Liste statt in einer eigenen Zeile darüber
        # – spart eine ganze Zeile Höhe zugunsten des Listenfensters.
        #
        # "Alle auswählen"/"Alles abwählen" und "Alle aufklappen"/"Alle
        # einklappen" sind je ein einzelner Toggle-Button (Text + Amber-Stil
        # wechseln je nach aktuellem Zustand) statt zwei Buttons pro Aktion –
        # spart zusätzlich Platz.
        toggle_row = ttk.Frame(toolbar)
        toggle_row.pack(side="top", anchor="w")

        self._load_btn = ttk.Button(
            toggle_row, text=self._LOAD_BTN_LABEL, command=self._load, state="disabled",
            style="AmberBold.TButton")
        self._load_btn.pack(side="left", padx=(0, 16))

        _btn_w = 18

        self._select_toggle_btn = ttk.Button(
            toggle_row, text=self._SELECT_ALL_BTN_LABEL,
            command=self._toggle_select_all, state="disabled", width=_btn_w)
        self._select_toggle_btn.pack(side="left", padx=(0, 4))

        self._expand_toggle_btn = ttk.Button(
            toggle_row, text=self._COLLAPSE_ALL_BTN_LABEL,
            command=self._toggle_expand_all, state="disabled", width=_btn_w)
        self._expand_toggle_btn.pack(side="left")

        # Eigene Checkbox "Assets mit ERRORs anzeigen" entfernt – redundant
        # zum Button "Fehlerhafte anzeigen" (STAC-Funktionen), der dieselbe
        # Variable steuert.
        self._error_filter_var = tk.BooleanVar(value=False)
        # Hält Text und Stil von self._show_faulty_btn (STAC-Funktionen,
        # "Assets prüfen") synchron mit dem Filterzustand. Amber-Stil
        # signalisiert, dass die gefilterte Ansicht aktiv ist.
        self._error_filter_var.trace_add("write", lambda *_: self._show_faulty_btn.config(
            text=self._SHOW_ALL_BTN_LABEL if self._error_filter_var.get()
                 else self._SHOW_FAULTY_BTN_LABEL,
            style="Amber.TButton" if self._error_filter_var.get() else "TButton"))

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
        self._stats_outer.pack(side="bottom", fill="x", pady=(0, 2))
        self._stats_lbl = tk.Label(
            self._stats_outer, text="Keine Daten geladen.",
            font=("Segoe UI", 9), anchor="w")
        self._stats_lbl.pack(side="left", padx=4)

    def _build_log(self, parent):
        # Log ist standardmässig eingeklappt (zeigt nur die letzte Meldung) –
        # der Inhalt ist meist nur "verbunden" / "N Assets geladen" und
        # braucht nicht dauerhaft Platz. Bei Bedarf per Klick aufklappbar.
        frm = ttk.LabelFrame(parent, text="Log",
                              padding=4, style="Section.TLabelframe")
        frm.pack(fill="x")

        header = ttk.Frame(frm)
        header.pack(fill="x")

        self._log_expanded = False
        self._log_toggle_btn = ttk.Button(
            header, text="▸ Details einblenden", width=22,
            command=self._toggle_log)
        self._log_toggle_btn.pack(side="left")

        self._log_preview_lbl = ttk.Label(
            header, text="Bereit.", style="Dim.TLabel",
            font=("Segoe UI", 8, "italic"))
        self._log_preview_lbl.pack(side="left", padx=(10, 0))

        self._log_body = ttk.Frame(frm)
        self._log = scrolledtext.ScrolledText(
            self._log_body, height=6, state="disabled",
            font=("Cascadia Mono", 8), wrap="word")
        self._log.pack(fill="both")

    def _toggle_log(self):
        self._log_expanded = not self._log_expanded
        if self._log_expanded:
            self._log_body.pack(fill="both", pady=(4, 0))
            self._log_toggle_btn.config(text="▾ Details ausblenden")
        else:
            self._log_body.pack_forget()
            self._log_toggle_btn.config(text="▸ Details einblenden")

    # ── GDWH-Tab (read-only) ─────────────────────────────────────────────────

    _GDWH_COLS      = ("year", "auftragstyp", "area", "stac_dt", "status", "gds_key")
    _GDWH_COL_HEADS = {"year": "Jahr", "auftragstyp": "Auftragstyp",
                        "area": "Area", "stac_dt": "StacItemDatetime", "status": "Status",
                        "gds_key": "GDS-Key"}
    _GDWH_COL_W     = {"year": 60, "auftragstyp": 110, "area": 200,
                        "stac_dt": 220, "status": 260, "gds_key": 130}

    _GDWH_LOAD_BTN_LABEL   = "Imports laden"
    _GDWH_RELOAD_BTN_LABEL = "Imports aktualisieren"
    _GDWH_ALL_KEYS_LABEL   = "Alle GDS-Keys"
    _GDWH_LEICHEN_BTN_LABEL      = f"GDWH-Anomalien anzeigen (>{_GDWH_PENDING_HOURS}h ohne Daten)"
    _GDWH_LEICHEN_BTN_LABEL_BACK = "Zurück zur normalen Ansicht"
    _GDWH_AUFTRAGSTYP_OPTIONS   = ["Alle", "RAM", "KRY"]

    def _build_gdwh_tab(self, parent):
        sec1 = ttk.LabelFrame(parent, text="1   Umgebung",
                              padding=8, style="Section.TLabelframe")
        sec1.pack(fill="x", pady=(0, 4))

        row1 = ttk.Frame(sec1)
        row1.pack(side="top", anchor="w")
        ttk.Label(row1, text="Umgebung:").pack(side="left", padx=(0, 6))
        self._gdwh_env_var = tk.StringVar(value="INT")
        for env in ("INT", "PROD"):
            ttk.Radiobutton(row1, text=env, variable=self._gdwh_env_var, value=env,
                            command=self._gdwh_on_env_change).pack(side="left", padx=4)
        self._gdwh_url_lbl = ttk.Label(row1, text=GDWH_ENVIRONMENTS["INT"],
                                       font=("Segoe UI", 8), style="Dim.TLabel")
        self._gdwh_url_lbl.pack(side="left", padx=12)

        ttk.Label(
            sec1,
            text="Authentifizierung: Windows-Session (aktuell eingeloggter User, wie im Browser) "
                 "– read-only, es wird nichts verändert/gelöscht.",
            font=("Segoe UI", 8, "italic"), style="Dim.TLabel",
        ).pack(side="top", anchor="w", pady=(6, 0))

        sec2 = ttk.LabelFrame(parent, text="2   GDS-Key & Filter",
                              padding=8, style="Section.TLabelframe")
        sec2.pack(fill="x", pady=(0, 4))

        ttk.Label(sec2, text="GDS-Key:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._gdwh_gds_key_var = tk.StringVar(value=GDWH_GDS_KEYS[0])
        gdwh_gds_combo = ttk.Combobox(
            sec2, textvariable=self._gdwh_gds_key_var,
            values=[self._GDWH_ALL_KEYS_LABEL] + GDWH_GDS_KEYS, state="readonly", width=24,
        )
        gdwh_gds_combo.grid(row=0, column=1, sticky="w", padx=(0, 16))

        self._optional_label(sec2, "Jahr [optional]:").grid(
            row=0, column=2, sticky="w", padx=(0, 6))
        self._gdwh_year_var = tk.StringVar()
        self._gdwh_year_var.trace_add("write", lambda *_: self._gdwh_apply_filter())
        ttk.Entry(sec2, textvariable=self._gdwh_year_var, width=8).grid(
            row=0, column=3, sticky="w")
        ttk.Label(
            sec2, text="z.B. 2023  —  Leer = alle Jahre",
            font=("Segoe UI", 8, "italic"), style="Dim.TLabel",
        ).grid(row=0, column=4, sticky="w", padx=(8, 0))

        # Normalansicht vs. GDWH-Anomalien: ein Import ohne FileMetadata-Match
        # (match is None) ist entweder frisch importiert und noch nicht
        # indexiert ("pending", jünger als _GDWH_PENDING_HOURS – bleibt in der
        # Normalansicht sichtbar) oder er existiert vermutlich nicht mehr
        # wirklich in GDWH (älter als _GDWH_PENDING_HOURS – nur der
        # Historieneintrag in GET /data/imports bleibt). Diese eigene,
        # separat einsehbare Kategorie sorgt dafür, dass die Standardansicht
        # nur noch wirklich in GDWH vorhandene Daten zeigt.
        self._gdwh_show_leichen_var = tk.BooleanVar(value=False)
        self._gdwh_leichen_btn = ttk.Button(
            sec2, text=self._GDWH_LEICHEN_BTN_LABEL,
            command=self._gdwh_toggle_leichen,
        )
        self._gdwh_leichen_btn.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(sec2, text="Auftragstyp:").grid(
            row=1, column=2, sticky="w", padx=(0, 6), pady=(6, 0))
        self._gdwh_auftragstyp_var = tk.StringVar(value=self._GDWH_AUFTRAGSTYP_OPTIONS[0])
        self._gdwh_auftragstyp_var.trace_add("write", lambda *_: self._gdwh_apply_filter())
        ttk.Combobox(
            sec2, textvariable=self._gdwh_auftragstyp_var,
            values=self._GDWH_AUFTRAGSTYP_OPTIONS, state="readonly", width=10,
        ).grid(row=1, column=3, sticky="w", pady=(6, 0))

        self._optional_label(sec2, "Area [optional]:").grid(
            row=1, column=4, sticky="w", padx=(16, 6), pady=(6, 0))
        self._gdwh_area_var = tk.StringVar()
        self._gdwh_area_var.trace_add("write", lambda *_: self._gdwh_apply_filter())
        ttk.Entry(sec2, textvariable=self._gdwh_area_var, width=16).grid(
            row=1, column=5, sticky="w", pady=(6, 0))

        self._gdwh_load_btn = ttk.Button(
            sec2, text=self._GDWH_LOAD_BTN_LABEL,
            command=self._gdwh_load,
        )
        self._gdwh_load_btn.grid(
            row=3, column=0, columnspan=6, sticky="w", pady=(10, 0))

        sec3 = ttk.LabelFrame(parent, text="3   DataPackages",
                              padding=4, style="Section.TLabelframe")
        sec3.pack(fill="both", expand=True, pady=(0, 4))
        sec3.rowconfigure(0, weight=1)
        sec3.columnconfigure(0, weight=1)

        self._gdwh_tree = ttk.Treeview(
            sec3, columns=self._GDWH_COLS, show="headings", selectmode="browse")
        for col in self._GDWH_COLS:
            self._gdwh_tree.column(col, width=self._GDWH_COL_W[col],
                                   minwidth=55, stretch=(col == "status"), anchor="w")
            self._gdwh_tree.heading(col, text=self._GDWH_COL_HEADS[col])

        vsb = ttk.Scrollbar(sec3, orient="vertical", command=self._gdwh_tree.yview)
        self._gdwh_tree.configure(yscrollcommand=vsb.set)
        self._gdwh_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._gdwh_stats_outer = tk.Frame(parent)
        self._gdwh_stats_outer.pack(fill="x", pady=(2, 0))
        self._gdwh_stats_lbl = tk.Label(
            self._gdwh_stats_outer, text="Keine Daten geladen.",
            font=("Segoe UI", 9), anchor="w")
        self._gdwh_stats_lbl.pack(side="left", padx=4)

    def _gdwh_on_env_change(self):
        self._gdwh_url_lbl.configure(text=GDWH_ENVIRONMENTS[self._gdwh_env_var.get()])
        self._gdwh_enriched = []
        self._gdwh_total_leichen = 0
        self._gdwh_total_pending = 0
        self._gdwh_current_gds_key = ""
        self._gdwh_tree.delete(*self._gdwh_tree.get_children())
        self._gdwh_stats_lbl.configure(text="Keine Daten geladen.")
        self._gdwh_load_btn.config(text=self._GDWH_LOAD_BTN_LABEL, style="TButton")
        self._gdwh_show_leichen_var.set(False)
        self._gdwh_leichen_btn.config(text=self._GDWH_LEICHEN_BTN_LABEL, style="TButton")

    def _gdwh_load(self):
        self._gdwh_load_btn.config(state="disabled")
        self._gdwh_base_url = GDWH_ENVIRONMENTS[self._gdwh_env_var.get()]
        gds_key = self._gdwh_gds_key_var.get()
        env     = self._gdwh_env_var.get()
        self._gdwh_current_gds_key = gds_key
        if gds_key == self._GDWH_ALL_KEYS_LABEL:
            threading.Thread(
                target=self._gdwh_worker_load_all, args=(env,), daemon=True).start()
        else:
            threading.Thread(
                target=self._gdwh_worker_load, args=(env, [gds_key]), daemon=True).start()

    def _gdwh_worker_load_all(self, env: str):
        self._gdwh_worker_load(env, GDWH_GDS_KEYS)

    def _gdwh_worker_load(self, env: str, gds_keys: List[str]):
        try:
            enriched: List[Tuple[Dict, Optional[Dict], str]] = []
            total_metadata = 0
            for gds_key in gds_keys:
                self._log_write(f"[GDWH] Lade Imports für {env}/{gds_key} …\n")
                imports = gdwh_get_imports(self._gdwh_base_url, gds_key)
                self._log_write(f"[GDWH] {len(imports)} Import(s) gefunden. "
                                f"Lade FileMetadata …\n")
                file_metadata = gdwh_search_file_metadata(self._gdwh_base_url, gds_key)
                meta_index = gdwh_index_file_metadata_by_import(file_metadata)
                total_metadata += len(file_metadata)

                for imp in imports:
                    match = meta_index.get(gdwh_import_id(imp))
                    enriched.append((imp, match, gds_key))

            self._gdwh_enriched = enriched
            self._log_write(f"[GDWH] {total_metadata} FileMetadata-Eintrag/Einträge "
                            f"gefunden, {len(enriched)} Import(s) angereichert.\n")
            self.after(0, self._gdwh_apply_filter)
            # Text/Style erst NACH erfolgreichem Laden auf "aktualisieren"/Amber
            # umstellen – bei einem Fehler unten bleibt der Ausgangszustand
            # bestehen, sonst würde ein fehlgeschlagener Ladeversuch fälschlich
            # als "bereits geladen" angezeigt.
            self.after(0, lambda: self._gdwh_load_btn.config(
                text=self._GDWH_RELOAD_BTN_LABEL, style="Amber.TButton"))

            pending_count = sum(1 for imp, match, _gds_key in enriched
                                 if match is None and _gdwh_is_pending(imp))
            if pending_count:
                self.after(50, lambda: self._gdwh_show_pending_notice(pending_count))
        except Exception as exc:
            self._log_write(f"[GDWH FEHLER] {exc}\n")
            self.after(0, lambda: messagebox.showerror("GDWH Fehler", str(exc)))
        finally:
            self.after(0, lambda: self._gdwh_load_btn.config(state="normal"))

    def _gdwh_show_pending_notice(self, count: int):
        """Info-Popup nach dem Laden: weist auf frisch importierte, noch nicht
        indexierte DataPackages hin (siehe _gdwh_is_pending/_GDWH_PENDING_HOURS).
        Läuft nur einmal pro Ladevorgang, nicht bei jeder Filteränderung."""
        messagebox.showinfo(
            "Frisch importierte DataPackages",
            f"{count} DataPackage(s) wurden erst kürzlich importiert. GDWH zeigt "
            f"sie im Portal zwar bereits an, aktualisiert seinen FileMetadata-Index "
            f"(Area/Jahr/Auftragstyp) aber zeitversetzt — das kann einige Stunden "
            f"dauern.\n\n"
            f"Diese Packages sind in der Liste mit „⏳ Frisch importiert“ markiert.\n\n"
            f"Bitte in ein paar Stunden nochmals „{self._GDWH_RELOAD_BTN_LABEL}“ "
            f"klicken, um den aktuellen Stand zu prüfen."
        )

    def _gdwh_toggle_leichen(self):
        active = not self._gdwh_show_leichen_var.get()
        self._gdwh_show_leichen_var.set(active)
        self._gdwh_leichen_btn.config(
            text=self._GDWH_LEICHEN_BTN_LABEL_BACK if active
                 else self._GDWH_LEICHEN_BTN_LABEL,
            style="Amber.TButton" if active else "TButton")
        self._gdwh_apply_filter()

    def _gdwh_apply_filter(self):
        year         = self._gdwh_year_var.get().strip()
        leichen_mode = self._gdwh_show_leichen_var.get()
        auftragstyp  = self._gdwh_auftragstyp_var.get()
        area_query   = self._gdwh_area_var.get().strip().lower()
        # Drei Kategorien pro Import (siehe _gdwh_populate_tree für die Status-Anzeige):
        #   - "aktiv"/"unvollständig": FileMetadata-Match vorhanden.
        #   - "pending": kein Match, aber Import jünger als _GDWH_PENDING_HOURS
        #     → GDWH hat die Attribute nur noch nicht nachindexiert, keine Anomalie.
        #   - "anomalie": kein Match und älter als _GDWH_PENDING_HOURS → GDWH hat
        #     dafür (noch) keinen FileMetadata-Eintrag geliefert.
        # WICHTIG: alle drei Kategorien werden in der Normalansicht IMMER
        # angezeigt (jede reale /data/imports-Zeile bleibt sichtbar, auch ohne
        # FileMetadata-Match – ein fehlender Match beweist nicht, dass die
        # Daten nicht mehr in GDWH existieren, siehe Portal-Abgleich). Der
        # Anomalien-Toggle filtert nur zusätzlich auf "anomalie", er blendet
        # in der Normalansicht nichts aus.
        def _is_leiche(item):
            imp, match, _gds_key = item
            return match is None and not _gdwh_is_pending(imp)

        if leichen_mode:
            data = [item for item in self._gdwh_enriched if _is_leiche(item)]
        else:
            data = list(self._gdwh_enriched)
        if year:
            def _year_matches(item):
                imp, match, _gds_key = item
                if match:
                    for src in (match.get("stac_datetime", ""), match.get("year", "")):
                        m = re.search(r"(?<!\d)(20\d{2})(?!\d)", src)
                        if m:
                            return m.group(1) == year
                m = re.search(r"(?<!\d)(20\d{2})(?!\d)", gdwh_import_date(imp))
                return m.group(1) == year if m else True
            data = [item for item in data if _year_matches(item)]
        if auftragstyp != "Alle":
            data = [item for item in data
                    if (item[1] or {}).get("auftragstyp", "").strip().lower() == auftragstyp.lower()]
        if area_query:
            data = [item for item in data
                    if area_query in (item[1] or {}).get("area", "").lower()]
        filtered = bool(year) or auftragstyp != "Alle" or bool(area_query)
        self._gdwh_total_leichen = sum(
            1 for item in self._gdwh_enriched if _is_leiche(item))
        self._gdwh_total_pending = sum(
            1 for imp, match, _gds_key in self._gdwh_enriched
            if match is None and _gdwh_is_pending(imp))
        self._gdwh_populate_tree(data, filtered=filtered, leichen_mode=leichen_mode)

    def _gdwh_populate_tree(self, enriched: List[Tuple], filtered: bool = False,
                             leichen_mode: bool = False):
        self._gdwh_tree.delete(*self._gdwh_tree.get_children())
        if not enriched:
            if leichen_mode:
                msg = "0 GDWH-Anomalien gefunden."
            elif filtered:
                msg = "0 DataPackages nach Filter."
            else:
                msg = "0 DataPackages gefunden."
            self._gdwh_stats_lbl.configure(text=msg)
            return

        def _year_key(item):
            imp, match, _gds_key = item
            if match:
                for src in (match.get("stac_datetime", ""), match.get("year", "")):
                    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", src)
                    if m:
                        return int(m.group(1))
            m = re.search(r"(?<!\d)(20\d{2})(?!\d)", gdwh_import_date(imp))
            return int(m.group(1)) if m else 0

        incomplete_count = 0
        for imp, match, gds_key in sorted(enriched, key=_year_key, reverse=True):
            area          = match.get("area", "")          if match else ""
            auftragstyp   = match.get("auftragstyp", "")   if match else ""
            stac_datetime = match.get("stac_datetime", "") if match else ""

            year = ""
            if match:
                for src in (stac_datetime, match.get("year", "")):
                    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", src)
                    if m:
                        year = m.group(1)
                        break
            pending = match is None and _gdwh_is_pending(imp)
            if not year:
                m = re.search(r"(?<!\d)(20\d{2})(?!\d)", gdwh_import_date(imp))
                year = m.group(1) if m else ("NEU" if pending else "????")

            if match is None and pending:
                # Frisch importiert, GDWH-Index zieht noch nach – siehe
                # _gdwh_is_pending. Keine Anomalie.
                status, tag = "⏳ Frisch importiert — Attribute folgen", "asset_warn"
            elif match is None:
                # Seit über _GDWH_PENDING_HOURS ohne FileMetadata-Match. Bleibt
                # trotzdem sichtbar (siehe _gdwh_apply_filter) – die Zeile
                # stammt aus /data/imports und ist damit ein realer Import,
                # auch wenn GDWH (noch) keinen FileMetadata-Treffer liefert.
                status, tag = f"⚠ Kein FileMetadata-Match seit über {_GDWH_PENDING_HOURS}h", "asset_err"
            elif not area or not stac_datetime:
                status, tag = "⚠ unvollständig", "asset_warn"
                incomplete_count += 1
            else:
                status, tag = "✓ OK", "asset_ok"

            self._gdwh_tree.insert(
                "", "end",
                values=(year, auftragstyp or "–", area or "–", stac_datetime or "–",
                        status, gds_key),
                tags=(tag,),
            )

        total = len(enriched)
        if leichen_mode:
            self._gdwh_stats_lbl.configure(
                text=f"{total} GDWH-Anomalie(n) — seit über {_GDWH_PENDING_HOURS}h "
                     f"ohne FileMetadata-Match (Filter auf diese Kategorie).")
            return

        anomaly_note = f"  |  ⚠ {incomplete_count} unvollständig" if incomplete_count else ""
        total_pending = getattr(self, "_gdwh_total_pending", 0)
        pending_note = (f"  |  {total_pending} frisch importiert, noch nicht indexiert"
                        ) if total_pending else ""
        total_leichen = getattr(self, "_gdwh_total_leichen", 0)
        leichen_note = (f"  |  {total_leichen} ohne FileMetadata-Match "
                        f"(Button „{self._GDWH_LEICHEN_BTN_LABEL}“ zum Herausfiltern)"
                        ) if total_leichen else ""
        self._gdwh_stats_lbl.configure(
            text=f"{total} DataPackage(s) geladen{anomaly_note}{pending_note}{leichen_note}")

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
        s.configure("TNotebook",
            background=T["panel"], bordercolor=T["sep"], tabmargins=(2, 4, 2, 0))
        s.configure("TNotebook.Tab",
            background=T["btn"], foreground=T["fg_dim"],
            bordercolor=T["sep"], padding=(12, 5), font=("Segoe UI", 9))
        s.map("TNotebook.Tab",
            background=[("selected", T["panel"]), ("active", T["btn_hover"])],
            foreground=[("selected", T["accent"]), ("active", T["fg"])],
            padding=[("selected", (16, 8)), ("!selected", (12, 5))],
            font=[("selected", ("Segoe UI", 9, "bold"))],
            expand=[("selected", (1, 1, 1, 0))])
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
        s.configure("Green.TButton",
            background=T["btn"], foreground=T["ok"],
            bordercolor=T["sep"], relief="flat", padding=(8, 4), focuscolor=T["panel"])
        s.map("Green.TButton",
            background=[("disabled", T["btn"]), ("active", T["btn_hover"]), ("pressed", T["sep"])],
            foreground=[("disabled", T["fg_dim"]), ("active", T["ok"])],
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
        s.configure("TCombobox",
            fieldbackground=T["input"], background=T["btn"], foreground=T["fg"],
            arrowcolor=T["fg"], bordercolor=T["sep"], insertcolor=T["fg"],
            selectbackground=T["sel_bg"], selectforeground=T["sel_fg"])
        s.map("TCombobox",
            fieldbackground=[("readonly", T["input"]), ("disabled", T["panel"])],
            foreground=[("readonly", T["fg"]), ("disabled", T["fg_dim"])],
            background=[("readonly", T["btn"]), ("active", T["btn_hover"])],
            arrowcolor=[("disabled", T["fg_dim"])])
        # Popdown-Listbox der Combobox ist ein natives Tk-Widget, nicht per
        # ttk.Style themebar – daher über die Option-Datenbank einfärben.
        self.option_add("*TCombobox*Listbox.background", T["input"])
        self.option_add("*TCombobox*Listbox.foreground", T["fg"])
        self.option_add("*TCombobox*Listbox.selectBackground", T["sel_bg"])
        self.option_add("*TCombobox*Listbox.selectForeground", T["sel_fg"])
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

        self._gdwh_tree.tag_configure("asset_ok",   foreground=T["tree_ok"])
        self._gdwh_tree.tag_configure("asset_err",  foreground=T["tree_err"])
        self._gdwh_tree.tag_configure("asset_warn", foreground=T["tree_warn"])

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
        self._gdwh_stats_outer.configure(bg=T["panel"])
        self._gdwh_stats_lbl.configure(bg=T["panel"], fg=T["fg_dim"])
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
        # Erster Aufruf kommt aus _build_filters(), BEVOR _build_stac_functions()
        # den Button überhaupt erzeugt hat – dort wird die Sichtbarkeit separat
        # gesetzt, hier nur bei echten (späteren) Auftragstyp-Wechseln nötig.
        if hasattr(self, "_show_no_thumb_btn"):
            self._update_no_thumb_btn_visibility()

    def _update_no_thumb_btn_visibility(self):
        """'ITEMs ohne Thumbnail' ergibt nur bei RAM (Thumbnail-Pflicht) Sinn –
        Button nur dort einblenden. Beim Wegschalten von RAM wird ein aktiver
        Filter automatisch zurückgesetzt (sonst bliebe die Ansicht unsichtbar
        gefiltert hängen)."""
        # winfo_manager() statt winfo_ismapped(): Letzteres hängt zusätzlich
        # davon ab, ob das Fenster gerade tatsächlich auf dem Bildschirm
        # sichtbar ist (z.B. False direkt nach dem Bauen, vor dem ersten
        # Map-Event) – winfo_manager() spiegelt zuverlässig nur den reinen
        # Pack-Zustand des Widgets.
        is_ram = AUFTRAGSTYPEN.get(self._auftragstyp_var.get(), "") == "ram"
        if is_ram:
            if self._show_no_thumb_btn.winfo_manager() != "pack":
                self._show_no_thumb_btn.pack(side="left", padx=(8, 0))
        else:
            if self._show_no_thumb_btn.winfo_manager() == "pack":
                self._show_no_thumb_btn.pack_forget()
            if self._show_no_thumb_only:
                self._show_no_thumb_only = False
                self._show_no_thumb_btn.config(
                    text=self._SHOW_NO_THUMB_BTN_LABEL, style="TButton")

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
        self._show_no_thumb_only = False
        self._show_no_thumb_btn.config(
            text=self._SHOW_NO_THUMB_BTN_LABEL, style="TButton")
        self._check_btn.config(style="TButton")
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
            self._export_links_btn.config(state="disabled")
            self._download_btn.config(state="disabled")
            self._create_links_btn.config(state="disabled")
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

    # ── Generischer Busy-Spinner für weitere Buttons ─────────────────────────
    # (eigenständig vom Lade-Spinner oben, damit mehrere Buttons parallel
    # "busy" sein können, z.B. Kartenviewer während einer HEAD-Prüfung.)

    def _start_btn_spinner(self, btn: ttk.Button, label: str) -> str:
        """Startet eine Spinner-Animation auf `btn` (Text wechselt zyklisch zu
        '<Frame>  <label>'), bis _stop_btn_spinner(btn, ...) sie beendet – für
        Hintergrund-Aktionen (Threads), die währenddessen echt animiert
        werden können. Gibt den ursprünglichen Button-Text zurück."""
        orig_text = btn.cget("text")
        self._btn_spinners[btn] = {"idx": 0, "after_id": None, "label": label}
        self._animate_btn_spinner(btn)
        return orig_text

    def _animate_btn_spinner(self, btn: ttk.Button):
        entry = self._btn_spinners.get(btn)
        if entry is None:
            return
        frame = self._SPINNER_FRAMES[entry["idx"] % len(self._SPINNER_FRAMES)]
        btn.config(text=f"{frame}  {entry['label']}")
        entry["idx"] += 1
        entry["after_id"] = self.after(120, lambda: self._animate_btn_spinner(btn))

    def _stop_btn_spinner(self, btn: ttk.Button, restore_text: str):
        entry = self._btn_spinners.pop(btn, None)
        if entry and entry["after_id"] is not None:
            self.after_cancel(entry["after_id"])
        btn.config(text=restore_text)

    def _run_blocking_with_spinner(self, btn: ttk.Button, label: str, fn):
        """Zeigt auf `btn` kurz einen Busy-Zustand, solange `fn()` synchron
        läuft. Da der Aufruf blockierend ist (kein Hintergrund-Thread), gibt
        es keine Bild-für-Bild-Animation wie bei _start_btn_spinner – aber
        Text/Deaktivierung werden per erzwungenem Redraw vor der (potenziell
        spürbaren) Operation sichtbar gemacht und danach zuverlässig wieder-
        hergestellt (auch bei einer Exception in fn)."""
        orig_text = btn.cget("text")
        btn.config(text=f"{self._SPINNER_FRAMES[0]}  {label}", state="disabled")
        self.update_idletasks()
        try:
            fn()
        finally:
            btn.config(text=orig_text, state="normal")

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
        Fehler (unbekannt statt fehlerhaft). Status -4 (Asset > 50 GB, von
        CloudFront korrekterweise mit 400 beantwortet) gilt ebenfalls nicht
        als Fehler."""
        info = self._asset_info.get(item_id, {}).get(asset_key)
        return bool(info) and info.get("status") not in (200, -4)

    def _on_error_filter_toggle(self):
        # Ansicht wechselt (alle Assets <-> nur Fehler) -> bisherige Auswahl
        # verwerfen, damit keine inzwischen unsichtbaren Assets exportiert/
        # geprüft werden.
        self._checked.clear()
        self._refresh_select_toggle_btn()
        self._apply_filters()

    def _toggle_faulty_filter(self):
        """Schaltet den Fehler-Filter um (self._error_filter_var, Text-/Stil-
        Sync von self._show_faulty_btn per trace_add in _build_tree())."""
        self._error_filter_var.set(not self._error_filter_var.get())
        self._on_error_filter_toggle()

    def _item_has_thumbnail(self, item: Dict) -> bool:
        return any(is_thumbnail_asset(k) or is_thumbnail_asset(v.get("href", ""))
                   for k, v in item.get("assets", {}).items())

    def _no_thumb_excluded(self, iid: str) -> bool:
        """True für Items, die planmässig nie ein Thumbnail haben (siehe
        _NO_THUMB_EXCLUDE_SUBSTR) – im 'ITEMs ohne Thumbnail'-Filter keine
        echten Kandidaten."""
        return self._NO_THUMB_EXCLUDE_SUBSTR in iid.lower()

    def _toggle_no_thumb_filter(self):
        """Blendet die Baumansicht auf Items OHNE Thumbnail-Asset ein/aus
        (nur bei Auftragstyp RAM verfügbar). Kombiniert sich mit den übrigen
        Filtern inkl. 'Fehlerhafte anzeigen'."""
        self._show_no_thumb_only = not self._show_no_thumb_only
        self._show_no_thumb_btn.config(
            text=self._SHOW_ALL_BTN_LABEL if self._show_no_thumb_only
                 else self._SHOW_NO_THUMB_BTN_LABEL,
            style="Amber.TButton" if self._show_no_thumb_only else "TButton")
        self._checked.clear()
        self._refresh_select_toggle_btn()
        self._apply_filters()

    def _apply_filters(self):
        if not self._all_items:
            return
        self._items_loaded_once = True
        year        = self._year_var.get().strip()
        area        = self._area_var.get().strip()
        search      = self._search_var.get().strip()
        exts        = self._active_extensions()
        terms       = self._active_terms()
        errors_only = self._error_filter_var.get()

        items = self._all_items
        if search:
            items = filter_items(items, search)
        if year:
            items = [it for it in items if stac_item_year(it) == year]
        if area:
            items = [it for it in items if area.lower() in stac_item_area(it).lower()]
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
        if self._show_no_thumb_only:
            items = [it for it in items
                     if not self._item_has_thumbnail(it)
                     and not self._no_thumb_excluded(it["id"])]

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
        # _populate_tree fügt Items immer aufgeklappt ein (open=True oben) –
        # Toggle-Button-Zustand nach jedem (Neu-)Aufbau daran angleichen.
        self._all_expanded = True
        self._refresh_expand_toggle_btn()
        self._toggle_tree_buttons(True)

    def _toggle_tree_buttons(self, on: bool):
        state = "normal" if on else "disabled"
        self._check_btn.config(state=state)
        self._export_links_btn.config(state=state)
        self._download_btn.config(state=state)
        self._create_links_btn.config(state=state)
        self._map_viewer_btn.config(state=state)
        self._viewer_win_btn.config(state=state)
        self._expand_toggle_btn.config(state=state)
        self._select_toggle_btn.config(state=state)
        # Unabhängig von der aktuellen Filter-Trefferzahl klickbar halten,
        # sonst könnten sich diese Toggle-Buttons selbst aussperren, falls der
        # gefilterte Blick (z.B. "Fehlerhafte anzeigen") gerade leer ist –
        # ein erneuter Klick müsste dann trotzdem wieder alle Assets zeigen
        # können.
        has_data = bool(self._all_items)
        self._show_faulty_btn.config(
            state="normal" if (has_data and self._assets_checked_once) else "disabled")
        # "ITEMs ohne Thumbnail" ist reine Metadaten-Prüfung – anders als der
        # Fehler-Filter unabhängig von einer HEAD-Prüfung nutzbar.
        self._show_no_thumb_btn.config(state="normal" if has_data else "disabled")

    def _expand_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=True)

    def _collapse_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=False)

    def _toggle_expand_all(self):
        """Toggle-Button: Text zeigt stets die Aktion, die ein Klick auslöst
        ('Alle einklappen' solange aufgeklappt, danach 'Alle aufklappen')."""
        if self._all_expanded:
            self._collapse_all()
        else:
            self._expand_all()
        self._all_expanded = not self._all_expanded
        self._refresh_expand_toggle_btn()

    def _refresh_expand_toggle_btn(self):
        self._expand_toggle_btn.config(
            text=self._COLLAPSE_ALL_BTN_LABEL if self._all_expanded
                 else self._EXPAND_ALL_BTN_LABEL,
            style="TButton" if self._all_expanded else "Amber.TButton")

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

    def _refresh_select_toggle_btn(self):
        """Text + Farbe des kombinierten Auswahl-Toggle-Buttons: 'Alles
        abwählen' + amber, sobald mindestens ein Asset (bzw. Item, da dessen
        Auswahl über seine Assets läuft) ausgewählt ist – sonst 'Alles
        auswählen' in neutraler Standardfarbe."""
        any_checked = any(self._checked.values())
        self._select_toggle_btn.config(
            text=self._DESELECT_ALL_BTN_LABEL if any_checked
                 else self._SELECT_ALL_BTN_LABEL,
            style="Amber.TButton" if any_checked else "TButton")

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
        self._refresh_select_toggle_btn()
        return "break"

    def _toggle_select_all(self):
        """Toggle-Button: solange irgendetwas ausgewählt ist, wählt ein Klick
        alles ab; sonst wählt er alles aus (Text/Farbe siehe
        _refresh_select_toggle_btn)."""
        if any(self._checked.values()):
            self._deselect_all()
        else:
            self._select_all()

    def _select_all(self):
        self._run_blocking_with_spinner(
            self._select_toggle_btn, "Wähle aus …", lambda: self._set_all_checked(True))
        # Erst NACH dem Spinner aktualisieren: _run_blocking_with_spinner
        # stellt sonst wieder den Button-Text von vor dem Klick her.
        self._refresh_select_toggle_btn()

    def _deselect_all(self):
        self._run_blocking_with_spinner(
            self._select_toggle_btn, "Wähle ab …", lambda: self._set_all_checked(False))
        self._refresh_select_toggle_btn()

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
        """Prüft ALLE unter dem aktuellen Filter aufgelisteten Assets (nicht nur
        ausgewählte) – analog zum 'Assets prüfen (HEAD)'-Button in
        topo-deleteDATAfromSTAC. Items ganz ohne Assets ('leere' Items) werden
        dabei ebenfalls erfasst und im Log ausgewiesen, auch wenn für sie kein
        HEAD-Request möglich ist."""
        # Spinner + Deaktivierung SOFORT sichtbar machen (per erzwungenem
        # Redraw), bevor die möglicherweise etwas dauernde Vorarbeit unten
        # (Item-/Asset-Liste zusammenstellen) läuft – sonst wirkt die GUI bis
        # zum Threadstart eingefroren, weil Tk erst beim nächsten Idle-
        # Durchlauf neu zeichnet.
        self._check_btn.config(state="disabled", style="TButton")
        orig_check_text = self._start_btn_spinner(self._check_btn, "Prüfe Assets …")
        self.update_idletasks()

        all_assets = [
            (d["item_id"], d["asset_key"], d["href"])
            for nid, d in self._nodes.items()
            if d["kind"] == "asset" and d.get("href")
        ]
        # Einmaliges Set statt pro Item über alle Knoten zu iterieren
        # (_item_asset_nids wäre hier O(n²) bei vielen Items/Assets).
        items_with_assets = {d["item_id"] for nid, d in self._nodes.items() if d["kind"] == "asset"}
        empty_items = sorted(
            d["item_id"] for nid, d in self._nodes.items()
            if d["kind"] == "item" and d["item_id"] not in items_with_assets)

        if not all_assets and not empty_items:
            self._log_write("[Prüfung] Keine Items/Assets im aktuellen Filter.\n")
            self._stop_btn_spinner(self._check_btn, orig_check_text)
            self._check_btn.config(state="normal")
            return

        # Wie im Delete-Tool: Thumbnails werden mitgeprüft, keine Sonderrolle.
        tasks = all_assets

        self._log_write(f"[Prüfung] {len(tasks)} Assets aus allen aufgelisteten Items …\n")
        if empty_items:
            self._log_write(
                f"[Prüfung] {len(empty_items)} Item(s) ohne Assets (leer): "
                + ", ".join(empty_items) + "\n")

        if not tasks:
            self._assets_checked_once = True
            self._stop_btn_spinner(self._check_btn, orig_check_text)
            self._check_btn.config(state="normal", style="Green.TButton")
            self._enable_error_filter_btn()
            return

        # Spinner setzen
        for iid, ak, _ in tasks:
            nid = f"asset::{iid}::{ak}"
            if self._tree.exists(nid):
                cur = self._tree.item(nid, "values")
                self._tree.item(nid, values=(cur[0], cur[1], "⟳", cur[3], "–", "–"),
                                tags=("asset_dim",))

        threading.Thread(target=self._worker_check, args=(tasks, orig_check_text), daemon=True).start()

    def _worker_check(self, tasks: List[Tuple[str, str, str]], orig_check_text: str):
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

                if sc in (200, -4):
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

                if sc not in (200, -4):
                    self._log_write(f"  {ak}  →  {stxt}  {_fmt_size(sz)}\n")

        self._log_write(
            f"[Prüfung] Fertig: ✓ {ok_cnt}  ✗ {err_cnt}  "
            f"|  Gesamtgrösse (OK): {_fmt_size(tot_sz)}\n")
        self.after(0, lambda: self._stop_btn_spinner(self._check_btn, orig_check_text))
        # Grün = alle Assets/Items unter den aktuellen Filtereinstellungen geprüft.
        self.after(0, lambda: self._check_btn.config(state="normal", style="Green.TButton"))
        self.after(0, self._enable_error_filter_btn)
        self.after(0, lambda: self._refresh_stats(ok_cnt, err_cnt, tot_sz))

    def _enable_error_filter_btn(self):
        self._assets_checked_once = True
        self._show_faulty_btn.config(state="normal")

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

    _LARGE_ASSET_HINT = (
        '  >50GB - Download nur per HTTP Range-Requests möglich (CloudFront-'
        'Limit), siehe swisstopo-Anleitung "Downloading Large Assets (> 50 GB)": '
        'https://docs.geo.admin.ch/download-data/stac-api/large-assets.html'
    )

    def _asset_is_large(self, item_id: str, asset_key: str) -> bool:
        """True, wenn eine vorherige HEAD-Prüfung das Asset als > 50 GB
        identifiziert hat (Status -4, siehe check_asset_info)."""
        info = self._asset_info.get(item_id, {}).get(asset_key)
        if not info:
            return False
        if info.get("status") == -4:
            return True
        sz = info.get("size_bytes")
        return bool(sz and sz > LARGE_ASSET_THRESHOLD_BYTES)

    def _create_download_links(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return

        exts  = self._active_extensions()
        terms = self._active_terms()
        _pfx  = COLLECTION_ID + "_"
        blocks      = []
        asset_count = 0
        large_count = 0
        for item in self._visible_items:
            iid     = item["id"]
            display = iid[len(_pfx):] if iid.startswith(_pfx) else iid
            year    = stac_item_year(item)
            area    = stac_item_area(item)
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
            info = " ".join(v for v in (year, area) if v)
            lines = [f"info: {info}", f"item: {display};", "asset: "]
            for ak, href in asset_entries:
                lines.append(ak)
                lines.append(f"- {href}")
                asset_count += 1
                if self._asset_is_large(iid, ak):
                    lines.append(self._LARGE_ASSET_HINT)
                    large_count += 1
            blocks.append("\n".join(lines))

        if not blocks:
            messagebox.showwarning("Keine Auswahl", "Keine ausgewählten Assets nach Filter.")
            return

        content = "\n\n\n".join(blocks) + "\n"

        def _on_saved(path):
            self._log_write(f"[Export] Download-Links: {path}\n")
            extra = f"  (davon {large_count} > 50GB)" if large_count else ""
            messagebox.showinfo("Export erfolgreich",
                                f"{asset_count} Asset-Link(s) exportiert{extra}.\n{path}")

        ExportPreviewDialog(
            self, self._dark, "Download-Links exportieren", content,
            initialfile=f"download_links_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.txt",
            filetypes=[("Textdatei", "*.txt"), ("Alle Dateien", "*.*")],
            defaultextension=".txt", on_saved=_on_saved,
        )

    # ── ASSET-Download ────────────────────────────────────────────────────────

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', "_", name)

    def _download_assets(self):
        tasks = [
            (d["item_id"], d["asset_key"], d["href"])
            for nid, d in self._nodes.items()
            if d["kind"] == "asset" and d.get("href") and self._is_checked(nid)
        ]
        if not tasks:
            messagebox.showwarning("Keine Auswahl", "Keine Assets ausgewählt.")
            return

        dest_dir = filedialog.askdirectory(title="Zielordner für Download wählen")
        if not dest_dir:
            return

        if not messagebox.askyesno(
                "Download starten",
                f"{len(tasks)} Asset(s) werden nach\n{dest_dir}\n"
                f"heruntergeladen (ein Unterordner pro Item). Assets > 50GB "
                f"werden automatisch per Range-Requests geladen. Fortfahren?"):
            return

        self._download_btn.config(state="disabled", style="TButton")
        orig_text = self._start_btn_spinner(self._download_btn, "Download läuft …")
        threading.Thread(target=self._worker_download,
                          args=(tasks, dest_dir, orig_text), daemon=True).start()

    def _download_progress_cb(self, ak: str):
        """Baut einen throttled progress_cb (max. alle 5s bzw. bei Prozent-
        wechsel eine Logzeile) für download_asset()."""
        state = {"t": 0.0, "pct": -1}

        def cb(downloaded: int, total: Optional[int]):
            now = time.monotonic()
            pct = (downloaded * 100 // total) if total else None
            if pct == state["pct"] and now - state["t"] < 5.0:
                return
            state["t"], state["pct"] = now, pct
            if total:
                msg = f"    {ak}: {pct}%  ({_fmt_size(downloaded)} / {_fmt_size(total)})\n"
            else:
                msg = f"    {ak}: {_fmt_size(downloaded)}\n"
            self.after(0, lambda: self._log_write(msg))
        return cb

    def _worker_download(self, tasks: List[Tuple[str, str, str]], dest_dir: str,
                          orig_text: str):
        ok_cnt = err_cnt = skip_cnt = 0
        tot_bytes = 0

        for iid, ak, href in tasks:
            item_dir = Path(dest_dir) / self._sanitize_filename(iid)
            item_dir.mkdir(parents=True, exist_ok=True)
            fname = Path(urlparse(href).path).name or self._sanitize_filename(ak)
            dest_path = item_dir / fname

            if dest_path.exists():
                skip_cnt += 1
                self._log_write(
                    f"[Download] {iid} / {ak}: bereits vorhanden – übersprungen "
                    f"({dest_path})\n")
                continue

            self._log_write(f"[Download] {iid} / {ak}  →  {dest_path}\n")
            tmp_path = dest_path.with_name(dest_path.name + ".part")
            result = download_asset(href, str(tmp_path), self._auth,
                                    progress_cb=self._download_progress_cb(ak))
            if result["ok"]:
                tmp_path.rename(dest_path)
                ok_cnt    += 1
                tot_bytes += result["bytes"]
                self._log_write(
                    f"[Download] {iid} / {ak}: OK ({_fmt_size(result['bytes'])})\n")
            else:
                err_cnt += 1
                if tmp_path.exists():
                    tmp_path.unlink()
                self._log_write(
                    f"[Download] {iid} / {ak}: FEHLER – {result['error']}\n")

        self._log_write(
            f"[Download] Fertig: ✓ {ok_cnt}  ✗ {err_cnt}  ⏭ {skip_cnt} übersprungen  "
            f"|  {_fmt_size(tot_bytes)} geladen\n")
        self.after(0, lambda: self._stop_btn_spinner(self._download_btn, orig_text))
        self.after(0, lambda: self._download_btn.config(
            state="normal", style="Green.TButton" if err_cnt == 0 else "TButton"))

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
        orig_map_text = self._start_btn_spinner(self._map_viewer_btn, "Löse Layer auf …")
        self._map_viewer_btn.config(state="disabled")
        threading.Thread(target=self._worker_open_map_viewer,
                          args=(targets, orig_map_text), daemon=True).start()

    def _worker_open_map_viewer(self, targets: List[Tuple[Dict, str, str]], orig_map_text: str):
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
            self._stop_btn_spinner(self._map_viewer_btn, orig_map_text)
            self._map_viewer_btn.config(state="normal")
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
            # Letzte nichtleere Zeile auch eingeklappt sichtbar machen – das ist
            # der Teil des Logs, den man im Alltag tatsächlich braucht.
            lines = [ln for ln in text.strip().splitlines() if ln.strip()]
            if lines:
                self._log_preview_lbl.config(text=lines[-1])
        self.after(0, _do)


# ─── Einstiegspunkt ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = StacMonitorApp()
    app.mainloop()
