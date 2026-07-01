"""
1_GUI_stac_monitor.py  –  STAC Monitoring-Tool (read-only)

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
    check_asset_info,
    stac_item_year, stac_item_area, stac_item_acq_date,
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


# ─── Haupt-Applikation ────────────────────────────────────────────────────────

class StacMonitorApp(tk.Tk):

    _COLS      = ("status", "typ", "groesse", "geaendert")
    _COL_HEADS = {"status": "Status", "typ": "Typ / Ext.",
                  "groesse": "Grösse", "geaendert": "Geändert"}
    _COL_W     = {"status": 100, "typ": 90, "groesse": 90, "geaendert": 105}

    def __init__(self):
        super().__init__()
        self.title("STAC Monitor  —  ch.swisstopo.spezialbefliegungen")
        self.minsize(1040, 720)

        self._dark: bool = True
        self._auth: Optional[Tuple] = None
        self._base_url: str = ""

        self._all_items: List[Dict] = []
        self._visible_items: List[Dict] = []

        # Baum-Metadaten: tree_iid → dict mit kind/item_id/asset_key/href/item
        self._nodes: Dict[str, Dict] = {}
        # Prüfergebnisse: {item_id: {asset_key: {status, size_bytes, last_modified}}}
        self._asset_info: Dict[str, Dict[str, Dict]] = {}

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

        self._cred_btn = ttk.Button(sec, text="Credentials laden",
                                     command=self._load_credentials)
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
        ttk.Label(sec, text="Jahr:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self._year_var = tk.StringVar()
        self._year_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._year_var, width=8).grid(
            row=1, column=1, sticky="w", pady=(6, 0))

        ttk.Label(sec, text="Item-ID / Suche:").grid(
            row=1, column=2, sticky="w", padx=(16, 6), pady=(6, 0))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(sec, textvariable=self._search_var, width=34).grid(
            row=1, column=3, sticky="w", pady=(6, 0))
        ttk.Label(sec, text="Teilstring genügt  (für direkten Abruf: vollständige ID)",
                  font=("Segoe UI", 8), style="Dim.TLabel").grid(
            row=1, column=4, sticky="w", padx=(8, 0), pady=(6, 0))

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
        ttk.Label(ext_frame, text="Frei:").pack(side="left", padx=(6, 4))
        self._ext_custom_var = tk.StringVar()
        self._ext_custom_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(ext_frame, textvariable=self._ext_custom_var, width=14).pack(side="left")

    def _build_actions(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 4))

        self._load_all_btn = ttk.Button(
            row, text="Alle Items laden", command=self._load_all, state="disabled")
        self._load_all_btn.pack(side="left", padx=(0, 4))

        self._load_one_btn = ttk.Button(
            row, text="Item direkt (exakte ID)", command=self._load_one, state="disabled")
        self._load_one_btn.pack(side="left", padx=(0, 16))

        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=(0, 16))

        self._check_btn = ttk.Button(
            row, text="Assets prüfen  (HEAD)", command=self._check_assets, state="disabled")
        self._check_btn.pack(side="left", padx=(0, 16))

        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=(0, 16))

        self._export_json_btn = ttk.Button(
            row, text="Export JSON (Kunden-Links)",
            command=self._export_json, state="disabled")
        self._export_json_btn.pack(side="left", padx=(0, 4))

        self._export_csv_btn = ttk.Button(
            row, text="Export CSV", command=self._export_csv, state="disabled")
        self._export_csv_btn.pack(side="left", padx=(0, 16))

        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=(0, 16))

        self._expand_btn = ttk.Button(
            row, text="Alle aufklappen", command=self._expand_all, state="disabled")
        self._expand_btn.pack(side="left", padx=(0, 4))

        self._collapse_btn = ttk.Button(
            row, text="Alle einklappen", command=self._collapse_all, state="disabled")
        self._collapse_btn.pack(side="left")

    def _build_tree(self, parent):
        frame = ttk.LabelFrame(parent, text="3   Items & Assets",
                               padding=4, style="Section.TLabelframe")
        frame.pack(fill="both", expand=True, pady=(0, 4))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            frame, columns=self._COLS, show="tree headings", selectmode="browse")

        self._tree.column("#0", width=430, minwidth=200, stretch=True)
        self._tree.heading("#0", text="Item / Asset")
        for col in self._COLS:
            self._tree.column(col, width=self._COL_W[col],
                              minwidth=55, stretch=False, anchor="center")
            self._tree.heading(col, text=self._COL_HEADS[col])

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self._ctx = tk.Menu(self, tearoff=0)
        self._tree.bind("<Button-3>",  self._on_right_click)
        self._tree.bind("<Double-1>",  self._on_double_click)

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
        self._load_all_btn.config(state="disabled")
        self._load_one_btn.config(state="disabled")

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
            self._load_all_btn.config(state="normal")
            self._load_one_btn.config(state="normal")
            self._log_write(f"[Credentials] {env} – {creds['username']}\n")
        except Exception as exc:
            T = DARK if self._dark else LIGHT
            self._cred_lbl.configure(text="Fehler!", foreground=T["err"])
            messagebox.showerror("Credentials-Fehler", str(exc))

    # ── Laden ─────────────────────────────────────────────────────────────────

    def _load_all(self):
        if not self._auth:
            return
        self._set_busy(True)
        self._all_items.clear()
        self._asset_info.clear()
        threading.Thread(target=self._worker_load_all, daemon=True).start()

    def _load_one(self):
        item_id = self._search_var.get().strip()
        if not item_id:
            messagebox.showwarning("Eingabe fehlt",
                                   "Bitte vollständige Item-ID im Suchfeld eingeben.")
            return
        self._set_busy(True)
        self._all_items.clear()
        self._asset_info.clear()
        threading.Thread(target=self._worker_load_one,
                         args=(item_id,), daemon=True).start()

    def _worker_load_all(self):
        try:
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

    def _worker_load_one(self, item_id: str):
        try:
            self._log_write(f"[Laden] Item direkt: {item_id} …\n")
            item = get_item_direct(self._base_url, self._auth, item_id)
            if item is None:
                self._log_write(f"[Info] Nicht gefunden: {item_id}\n")
                self.after(0, lambda: messagebox.showinfo(
                    "Nicht gefunden", f"Item nicht gefunden:\n{item_id}"))
            else:
                self._all_items = [item]
                self._log_write(f"[OK] {item['id']} geladen.\n")
                self.after(0, self._apply_filters)
        except Exception as exc:
            self._log_write(f"[FEHLER] {exc}\n")
            self.after(0, lambda: messagebox.showerror("Fehler", str(exc)))
        finally:
            self.after(0, lambda: self._set_busy(False))

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self._load_all_btn.config(state=state if self._auth else "disabled")
        self._load_one_btn.config(state=state if self._auth else "disabled")
        if busy:
            self._check_btn.config(state="disabled")
            self._export_json_btn.config(state="disabled")
            self._export_csv_btn.config(state="disabled")

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
            year    = stac_item_year(item)
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

            meta = "  ".join(p for p in [year, area, acq] if p)
            label = display + (f"   [{meta}]" if meta else "")

            node_id = f"item::{iid}"
            self._tree.insert("", "end", iid=node_id,
                              text=f"  {label}",
                              values=("", "", f"{len(asset_keys)} Assets", ""),
                              tags=("item",), open=True)
            self._nodes[node_id] = {"kind": "item", "item_id": iid, "item": item}

            item_info = self._asset_info.get(iid, {})
            for ak in sorted(asset_keys):
                aval     = assets.get(ak, {})
                href     = aval.get("href", "")
                atype    = aval.get("type", "")
                ext      = Path(href).suffix if href else ""
                info     = item_info.get(ak)
                sc       = info.get("status")   if info else None
                sz       = info.get("size_bytes") if info else None
                lm       = info.get("last_modified") if info else None
                stxt, tg = _status_label(sc)

                anid = f"asset::{iid}::{ak}"
                self._tree.insert(node_id, "end", iid=anid,
                                  text=f"        {ak}",
                                  values=(stxt, ext or atype[:22],
                                          _fmt_size(sz), _fmt_date(lm)),
                                  tags=(tg,))
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
        self._expand_btn.config(state=state)
        self._collapse_btn.config(state=state)

    def _expand_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=True)

    def _collapse_all(self):
        for node in self._tree.get_children():
            self._tree.item(node, open=False)

    # ── HEAD-Prüfung ──────────────────────────────────────────────────────────

    def _check_assets(self):
        tasks = [
            (d["item_id"], d["asset_key"], d["href"])
            for nid, d in self._nodes.items()
            if d["kind"] == "asset" and d.get("href")
        ]
        if not tasks:
            self._log_write("[Prüfung] Keine Assets mit URL.\n")
            return

        self._check_btn.config(state="disabled")
        self._log_write(f"[Prüfung] {len(tasks)} Assets …\n")

        # Spinner setzen
        for iid, ak, _ in tasks:
            nid = f"asset::{iid}::{ak}"
            if self._tree.exists(nid):
                cur = self._tree.item(nid, "values")
                self._tree.item(nid, values=("⟳", cur[1], "–", "–"),
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
                cur_typ = ""
                if self._tree.exists(nid):
                    cur_typ = self._tree.item(nid, "values")[1]

                self.after(0, lambda n=nid, s=stxt, t=cur_typ,
                           sz_=_fmt_size(sz), lm_=_fmt_date(lm), tag=tg:
                           self._tree.exists(n) and
                           self._tree.item(n, values=(s, t, sz_, lm_), tags=(tag,)))

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

        self._ctx.tk_popup(event.x_root, event.y_root)

    def _on_double_click(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        d = self._nodes.get(row, {})
        if d.get("kind") == "item":
            ItemJsonDialog(self, d["item"], self._dark)
        elif d.get("kind") == "asset":
            href = d.get("href", "")
            if href:
                webbrowser.open(href)

    def _clip(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self._log_write(f"[Clipboard] {text}\n")

    # ── Export ────────────────────────────────────────────────────────────────

    def _export_json(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("Alle Dateien", "*.*")],
            title="Download-Links exportieren (JSON)",
            initialfile=f"stac_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
        )
        if not path:
            return

        exts       = self._active_extensions()
        items_out  = []
        for item in self._visible_items:
            iid    = item["id"]
            assets = item.get("assets", {})
            asset_list = []
            for ak, aval in assets.items():
                href = aval.get("href", "")
                if exts and not any(href.lower().endswith(e) or ak.lower().endswith(e)
                                    for e in exts):
                    continue
                entry: Dict = {"key": ak, "href": href}
                if aval.get("type"):
                    entry["media_type"] = aval["type"]
                if aval.get("title"):
                    entry["title"] = aval["title"]
                info = self._asset_info.get(iid, {}).get(ak, {})
                if info.get("status") is not None:
                    entry["http_status"] = info["status"]
                if info.get("size_bytes") is not None:
                    entry["size_bytes"] = info["size_bytes"]
                asset_list.append(entry)
            if not asset_list:
                continue
            items_out.append({
                "item_id":  iid,
                "acq_date": stac_item_acq_date(item),
                "area":     stac_item_area(item),
                "assets":   asset_list,
            })

        output = {
            "meta": {
                "export_date":  datetime.now().isoformat(timespec="seconds"),
                "environment":  self._env_var.get(),
                "collection":   COLLECTION_ID,
                "tool":         "STAC Monitor",
                "item_count":   len(items_out),
                "asset_count":  sum(len(it["assets"]) for it in items_out),
            },
            "items": items_out,
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(output, f, indent=2, ensure_ascii=False)
            self._log_write(f"[Export] JSON: {path}\n")
            messagebox.showinfo("Export erfolgreich",
                                f"{len(items_out)} Items  |  "
                                f"{output['meta']['asset_count']} Assets\n{path}")
        except Exception as exc:
            messagebox.showerror("Export-Fehler", str(exc))

    def _export_csv(self):
        if not self._visible_items:
            messagebox.showwarning("Keine Daten", "Keine Items geladen.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Alle Dateien", "*.*")],
            title="Export CSV",
            initialfile=f"stac_monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )
        if not path:
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
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
                writer.writeheader()
                writer.writerows(rows)
            self._log_write(f"[Export] CSV: {path}\n")
            messagebox.showinfo("Export erfolgreich",
                                f"{len(rows)} Zeilen exportiert.\n{path}")
        except Exception as exc:
            messagebox.showerror("Export-Fehler", str(exc))

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
