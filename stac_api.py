"""
stac_api.py  –  STAC API Hilfsfunktionen für ch.swisstopo.spezialbefliegungen
(Monitoring-Version: read-only, keine Delete-Funktionen)
"""

import re
import requests
import urllib3
from urllib.parse import urljoin
from typing import Dict, List, Optional, Tuple

# verify=False ist im Bundesnetz nötig, da proxy-bvcol.admin.ch HTTPS mit
# eigenem Zertifikat terminiert (TLS-Interception). Die dadurch bei jedem
# Request ausgelöste InsecureRequestWarning wird deshalb bewusst unterdrückt.
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# Firmenproxy für externe Verbindungen (data.geo.admin.ch / sys-data.int.bgdi.ch)
_PROXY = {
    "http":  "http://proxy-bvcol.admin.ch:8080",
    "https": "http://proxy-bvcol.admin.ch:8080",
}

COLLECTION_ID = "ch.swisstopo.spezialbefliegungen"

ENVIRONMENTS = {
    "INT":  "https://sys-data.int.bgdi.ch/api/stac/v0.9/",
    "PROD": "https://data.geo.admin.ch/api/stac/v0.9/",
}

# Hash-Routing-Basis des STAC-Browsers je Umgebung (für Kunden-Weitergabe).
# INT läuft direkt auf der Domain-Root, PROD unter /browser/index.html.
_BROWSER_BASE = {
    "INT":  "https://sys-data.int.bgdi.ch/#/collections/{cid}",
    "PROD": "https://data.geo.admin.ch/browser/index.html#/collections/{cid}",
}


def browser_url(env: str, item_id: Optional[str] = None, include_lang: bool = True) -> str:
    """Liefert den STAC-Browser-Link zur Collection, optional zu einem Item.

    ``include_lang=False`` liefert den Link ohne "?.language=en"-Anhängsel,
    z.B. für eine saubere Dokumentation in Exportdateien.
    """
    url = _BROWSER_BASE[env].format(cid=COLLECTION_ID)
    if item_id:
        url += f"/items/{item_id}"
    return (url + "?.language=en") if include_lang else url


# Assets, die map.geo.admin.ch als Cloud-Optimized-GeoTIFF (COG-Layer) direkt
# darstellen kann. Andere Formate (z.B. .copc.laz) unterstützt der Kartenviewer
# über den "layers=COG|..."-Mechanismus nicht.
_COG_EXTENSIONS = (".tif", ".tiff")

_MAP_VIEWER_BASE = "https://map.geo.admin.ch/#/map"


def is_cog_asset(href: str) -> bool:
    """Prüft, ob ein Asset-Href als COG-Layer im map.geo.admin.ch-Kartenviewer
    darstellbar ist (aktuell nur GeoTIFF/.tif/.tiff)."""
    return href.lower().endswith(_COG_EXTENSIONS)


def map_viewer_url(hrefs: List[str]) -> str:
    """Baut einen map.geo.admin.ch-Link, der die übergebenen COG-Asset-URLs
    (GeoTIFF) direkt als Layer einblendet (layers=COG|url1;COG|url2;...).

    Die Asset-URLs werden unverändert (nicht prozentkodiert) eingesetzt, da der
    Kartenviewer die verschachtelte URL innerhalb des Pipe-getrennten
    Layer-Ausdrucks roh erwartet (analog zu WMS|.../WMTS|...-Syntax).
    """
    layers = ";".join(f"COG|{href}" for href in hrefs)
    return f"{_MAP_VIEWER_BASE}?layers={layers}"

AUFTRAGSTYPEN: Dict[str, str] = {
    "KRY (Kryosphäre)":   "kry",
    "RAM (Rapidmapping)": "ram",
    "Alle":               "",
}

EXT_PRESETS: List[Tuple[str, List[str]]] = [
    ("tif / tiff",      [".tif", ".tiff"]),
    ("copc.laz / laz",  [".copc.laz", ".laz"]),
    ("jpg / jpeg",      [".jpg", ".jpeg"]),
    ("png",             [".png"]),
    ("json",            [".json"]),
]


# ─── Interne Session-Funktionen ───────────────────────────────────────────────

# None = noch nicht getestet, True/False = Ergebnis des ersten Verbindungsversuchs.
# Ausserhalb des Bundesnetz (z.B. privater Rechner) ist proxy-bvcol.admin.ch nicht
# auflösbar -> nach einmaligem ProxyError auf Direktverbindung umschalten.
_USE_PROXY: Optional[bool] = None

# Gemeinsame Session für Connection-Pooling (Keep-Alive statt neuem TCP/TLS-
# Handshake pro Request) – wichtig für Performance bei Pagination und den
# parallelen HEAD-Checks. requests.Session ist für Multithreading geeignet,
# solange kein gemeinsamer Zustand (Header etc.) zur Laufzeit verändert wird.
_SESSION = requests.Session()
_SESSION.mount("https://", requests.adapters.HTTPAdapter(pool_maxsize=16))


def _request(method: str, url: str, **kwargs) -> requests.Response:
    global _USE_PROXY
    call = getattr(_SESSION, method)
    if _USE_PROXY is False:
        return call(url, proxies=None, **kwargs)
    try:
        r = call(url, proxies=_PROXY, **kwargs)
        _USE_PROXY = True
        return r
    except requests.exceptions.ProxyError:
        _USE_PROXY = False
        return call(url, proxies=None, **kwargs)


def _session_get(url: str, auth: Tuple, params: dict = None) -> requests.Response:
    return _request("get", url, auth=auth, params=params,
                    verify=False, timeout=(30, 60))


# ─── Öffentliche API-Funktionen ───────────────────────────────────────────────

def get_item_direct(base_url: str, auth: Tuple, item_id: str) -> Optional[Dict]:
    """Holt ein einzelnes Item per exakter ID. Gibt None bei 404 zurück."""
    url = urljoin(base_url, f"collections/{COLLECTION_ID}/items/{item_id.strip()}")
    r = _session_get(url, auth)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def get_collection_items(base_url: str, auth: Tuple, log_fn=print) -> List[Dict]:
    """Holt alle Items der Collection mit Paginierung."""
    all_items = []
    url    = urljoin(base_url, f"collections/{COLLECTION_ID}/items")
    params = {"limit": 1000}
    while url:
        r = _session_get(url, auth, params)
        r.raise_for_status()
        data = r.json()
        all_items.extend(data.get("features", []))
        nxt = next((lk for lk in data.get("links", []) if lk.get("rel") == "next"), None)
        if nxt:
            url    = nxt["href"]
            params = None
            log_fn(f"  Paginierung … bisher {len(all_items)} Items geladen\n")
        else:
            url = None
    return all_items


def filter_items(items: List[Dict], search_term: str = "") -> List[Dict]:
    """Filtert Items nach Teilstring in der ID (case-insensitive)."""
    if not search_term:
        return items
    term = search_term.lower()
    return [item for item in items if term in item.get("id", "").lower()]


def check_asset_info(href: str, auth: Tuple) -> Dict:
    """HEAD-Request auf Asset-URL.
    Gibt dict mit status, size_bytes und last_modified zurück.
    status: HTTP-Code oder negativ (-1=kein href, -2=timeout, -3=exception)."""
    result: Dict = {"status": -1, "size_bytes": None, "last_modified": None}
    if not href:
        return result
    try:
        r = _request("head", href, verify=False,
                    timeout=(5, 15), allow_redirects=True)
        if r.status_code in (401, 403):
            r = _request("head", href, auth=auth, verify=False,
                        timeout=(5, 15), allow_redirects=True)
        result["status"] = r.status_code
        cl = r.headers.get("Content-Length", "")
        if cl.isdigit():
            result["size_bytes"] = int(cl)
        lm = r.headers.get("Last-Modified")
        if lm:
            result["last_modified"] = lm
    except requests.exceptions.Timeout:
        result["status"] = -2
    except Exception:
        result["status"] = -3
    return result


def stac_item_acq_date(item: Dict) -> str:
    """Gibt das Aufnahmedatum aus der Item-ID zurück (Format: YYYY-MM-DD)."""
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", item.get("id", ""))
    if m:
        return m.group(1)
    return item.get("properties", {}).get("datetime", "")


def stac_item_year(item: Dict) -> str:
    """Extrahiert das Aufnahmejahr aus der Item-ID."""
    m = re.search(r"\b(20\d{2})\b", item.get("id", ""))
    if m:
        return m.group(1)
    m = re.search(r"\b(20\d{2})\b", item.get("properties", {}).get("datetime", ""))
    return m.group(1) if m else ""


# Bekannte Schlüssel im "Key: Value, Key: Value, ..."-Format der Asset-Description
# (z.B. "Area: RANDA, TerrainModel: ..., Acquisition time: t1,t2,t3, LineId: ...").
# Einzelne Werte (Acquisition time, LineId) enthalten selbst Kommas – ein Split
# ausschliesslich anhand dieser bekannten Schlüssel verhindert falsches Zerteilen.
_ASSET_DESC_KEYS = [
    "Area", "TerrainModel", "SourceReferenceSystem", "CameraSystem",
    "Acquisition time", "LineId", "Commentary",
]


def parse_asset_description(description: str) -> Dict[str, str]:
    """Zerlegt die Asset-Description in ein Dict der bekannten Schlüssel."""
    if not description:
        return {}
    pattern = r"(" + "|".join(re.escape(k) for k in _ASSET_DESC_KEYS) + r"):\s*"
    matches = list(re.finditer(pattern, description))
    result: Dict[str, str] = {}
    for i, m in enumerate(matches):
        start = m.end()
        end   = matches[i + 1].start() if i + 1 < len(matches) else len(description)
        result[m.group(1)] = description[start:end].rstrip(", ").strip()
    return result


def asset_area(asset: Dict) -> str:
    """Extrahiert den 'Area'-Wert aus der Asset-Description, falls vorhanden."""
    return parse_asset_description(asset.get("description", "")).get("Area", "")


def stac_item_area(item: Dict) -> str:
    """Gibt den AOI-Namen zurück: zuerst aus Item-Properties (falls vorhanden),
    sonst aus der Description des ersten passenden Assets."""
    props = item.get("properties", {})
    for key in ("area", "aoi", "area_name", "region"):
        val = str(props.get(key, "")).strip()
        if val:
            return val.upper()
    for asset in item.get("assets", {}).values():
        area = asset_area(asset)
        if area:
            return area.upper()
    return ""


# ─── STAC-1.0.0-Item-Export ────────────────────────────────────────────────────
#
# Die Items dieses Tools stammen bereits als vollständige, valide STAC-1.0.0-
# Features direkt von der swisstopo-STAC-API (data.geo.admin.ch / sys-data.int.
# bgdi.ch) – inkl. geometry/bbox bereits in WGS84 (EPSG:4326), wie von der STAC-
# Spec zwingend gefordert. Es gibt im Tool aktuell keine LV95(EPSG:2056)-Daten.
# Die folgende Transformation ist daher eine reine Absicherung für den Fall,
# dass geometry/bbox eines Items (z.B. aus einer künftigen Datenquelle) doch in
# LV95 vorliegen – im Normalfall ist sie ein No-Op (Passthrough).

# Wertebereich LV95 (EPSG:2056): Easting E ~2.48–2.84 Mio, Northing N ~1.07–1.30
# Mio. Damit eindeutig von WGS84 Lon/Lat (-180..180 / -90..90) unterscheidbar.
_LV95_E_RANGE = (2_400_000, 2_900_000)
_LV95_N_RANGE = (1_000_000, 1_400_000)


def _is_lv95_coord(x: float, y: float) -> bool:
    """Prüft, ob ein Koordinatenpaar plausibel im LV95-Wertebereich (EPSG:2056) liegt."""
    return _LV95_E_RANGE[0] <= x <= _LV95_E_RANGE[1] and _LV95_N_RANGE[0] <= y <= _LV95_N_RANGE[1]


def _lv95_to_wgs84_transform():
    """Baut die Koordinatentransformation LV95 (EPSG:2056) -> WGS84 (EPSG:4326)
    via osgeo.osr. Lazy Import: osgeo wird nur geladen, wenn tatsächlich LV95-
    Koordinaten erkannt werden, damit das Tool ohne osgeo4w-Umgebung lauffähig
    bleibt, solange keine LV95-Transformation nötig ist."""
    from osgeo import osr
    quelle = osr.SpatialReference()
    quelle.ImportFromEPSG(2056)
    ziel = osr.SpatialReference()
    ziel.ImportFromEPSG(4326)
    # Zwingend setzen: sonst liefert osr (X=lat, Y=lon) statt der von GeoJSON/STAC
    # verlangten Reihenfolge [lon, lat] -> ohne dies wären geometry/bbox vertauscht.
    ziel.SetAxisMappingStrategy(osr.OAMS_TRADITIONAL_GIS_ORDER)
    return osr.CoordinateTransformation(quelle, ziel)


def _transform_coords(coords, transform):
    """Transformiert rekursiv verschachtelte GeoJSON-Koordinaten (Polygon etc.)
    von LV95 nach WGS84. Koordinaten ausserhalb des LV95-Wertebereichs (bereits
    WGS84) werden unverändert durchgereicht."""
    if isinstance(coords[0], (int, float)):
        x, y = coords[0], coords[1]
        if _is_lv95_coord(x, y):
            lon, lat, _ = transform.TransformPoint(x, y)
            return [lon, lat]
        return [x, y]
    return [_transform_coords(c, transform) for c in coords]


def _flatten_points(coords):
    """Liefert alle [lon, lat]-Punkte einer (verschachtelten) GeoJSON-Koordinatenliste."""
    if isinstance(coords[0], (int, float)):
        yield coords
    else:
        for c in coords:
            yield from _flatten_points(c)


def _ensure_wgs84(geometry: Optional[Dict], bbox: Optional[List[float]]) -> Tuple[Optional[Dict], Optional[List[float]]]:
    """Stellt sicher, dass geometry/bbox in WGS84 (EPSG:4326) vorliegen, wie von
    STAC 1.0.0 zwingend gefordert. Normalfall: Koordinaten sind bereits WGS84 ->
    No-Op. Nur falls die Koordinaten im LV95-Wertebereich liegen, wird via
    osgeo.osr nach WGS84 transformiert und die bbox neu aus der Geometrie berechnet."""
    if not geometry or not geometry.get("coordinates"):
        return geometry, bbox

    coords = geometry["coordinates"]
    erster_punkt = next(_flatten_points(coords), None)
    if not erster_punkt or not _is_lv95_coord(erster_punkt[0], erster_punkt[1]):
        return geometry, bbox

    transform = _lv95_to_wgs84_transform()
    neue_coords = _transform_coords(coords, transform)
    neue_geometry = {**geometry, "coordinates": neue_coords}
    punkte = list(_flatten_points(neue_coords))
    lons = [p[0] for p in punkte]
    lats = [p[1] for p in punkte]
    neue_bbox = [min(lons), min(lats), max(lons), max(lats)]
    return neue_geometry, neue_bbox


def build_stac_item(item: Dict, assets: Dict) -> Dict:
    """Baut ein valides STAC-1.0.0-Item (GeoJSON Feature) für den Export.

    Übernimmt die STAC-Pflichtfelder aus dem Original-Item (das bereits ein
    valides Item der swisstopo-API ist) und ersetzt nur "assets" durch die vom
    Aufrufer gefilterte Auswahl (z.B. Extension-/Checkbox-Filter im GUI).
    Die interne Datenhaltung des Tools bleibt davon unberührt.
    """
    geometry, bbox = _ensure_wgs84(item.get("geometry"), item.get("bbox"))

    properties = dict(item.get("properties", {}))
    if not properties.get("datetime"):
        # Fallback: Aufnahmedatum aus der Item-ID (siehe stac_item_acq_date),
        # als ISO-8601-UTC-Datetime gemäss STAC-Vorgabe.
        acq = stac_item_acq_date(item)
        properties["datetime"] = f"{acq}T00:00:00Z" if acq else None

    stac_item: Dict = {
        # Fest auf "1.0.0", unabhängig von der Quell-API-Version (der swisstopo-
        # API-Endpunkt liefert aktuell "0.9.0" in stac_version, obwohl die
        # Item-Struktur bereits 1.0.0-kompatibel ist) – der Export soll immer
        # ein STAC-1.0.0-Item deklarieren.
        "type":         "Feature",
        "stac_version": "1.0.0",
        "id":           item.get("id"),
        "geometry":     geometry,
        "bbox":         bbox,
        "properties":   properties,
        "links":        item.get("links", []),
        "assets":       assets,
    }
    if item.get("collection"):
        stac_item["collection"] = item["collection"]
    if item.get("stac_extensions"):
        stac_item["stac_extensions"] = item["stac_extensions"]
    return stac_item


if __name__ == "__main__":
    print("stac_api.py – STAC Monitoring Modul (read-only)")
    print(f"  Collection:  {COLLECTION_ID}")
    print(f"  Umgebungen:  {list(ENVIRONMENTS.keys())}")
