"""
gdwh_api.py  –  GDWH API Hilfsfunktionen (read-only)

Authentifizierung: Windows SSPI (HttpNegotiateAuth) – kein Benutzername/Passwort nötig.
Der aktuell eingeloggte Windows-User wird automatisch verwendet (gleich wie Browser).

Nur lesende Endpunkte – dieses Tool löscht/verändert nichts im GDWH:
  GET  /api/geodatasets/{gdsKey}/data/imports              → DataPackages laden
  POST /api/geodatasets/{gdsKey}/fileMetadata/search        → fachliche Attribute
       (Area, Jahr, Auftragstyp, StacItemIdDatetime) je Import

Swagger (INT): https://ltgdwhi.adr.admin.ch/gdwh-api/v2/swagger/index.html
"""

import json
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from typing import Dict, List
import requests
import urllib3


_FALLBACK_PROXY = "http://proxy-bvcol.admin.ch:8080"


def _pip_install(pkg: str) -> bool:
    """Installiert ein Paket via pip. Versucht zuerst Proxies aus proxy_config.json,
    dann den Firmen-Fallback-Proxy, zuletzt ohne Proxy. Gibt True bei Erfolg zurück."""
    config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "secrets", "proxy_config.json")
    proxies = []
    if os.path.exists(config_path):
        try:
            with open(config_path, encoding="utf-8") as f:
                cfg = json.load(f)
            proxies = [p["url"] for p in cfg.get("proxies", []) if p.get("enabled") and p.get("url")]
        except Exception:
            pass
    if _FALLBACK_PROXY not in proxies:
        proxies.append(_FALLBACK_PROXY)

    trusted = ["--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org"]
    attempts = []
    for proxy in proxies:
        attempts.append([sys.executable, "-m", "pip", "install", "--user", pkg,
                         "--proxy", proxy] + trusted)
    attempts.append([sys.executable, "-m", "pip", "install", "--user", pkg] + trusted)

    for cmd in attempts:
        try:
            subprocess.check_call(cmd)
            return True
        except subprocess.CalledProcessError:
            continue
    return False


try:
    from requests_negotiate_sspi import HttpNegotiateAuth
except ImportError:
    print("Installiere requests-negotiate-sspi ...")
    if not _pip_install("requests-negotiate-sspi"):
        raise RuntimeError(
            "Installation von requests-negotiate-sspi fehlgeschlagen.\n"
            "Bitte manuell installieren:\n"
            f"  python -m pip install --user requests-negotiate-sspi "
            f"--proxy http://proxy-bvcol.admin.ch:8080"
        )
    from requests_negotiate_sspi import HttpNegotiateAuth

# Interne Firmen-CA nicht im Python-Truststore → Verifikation deaktivieren.
GDWH_SSL_VERIFY: bool = False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def _gdwh_session() -> requests.Session:
    """Session mit frischer SSPI-Auth, direkte Verbindung ohne Proxy.
    GDWH ist ein internes System (ltgdwhi/ltgdwh.adr.admin.ch) – kein Internet-Proxy.
    requests.Session() hält die TCP-Verbindung für den mehrstufigen SSPI-Handshake aufrecht."""
    s = requests.Session()
    s.auth    = HttpNegotiateAuth()
    s.verify  = GDWH_SSL_VERIFY
    s.proxies = {"http": "", "https": ""}
    return s


GDWH_GDS_KEYS = [
    "SB_DOP",
    "SB_DOP_16",
    "SB_DSM",
    "SB_DSM_PUNKTWOLKE",
]

GDWH_ENVIRONMENTS = {
    "INT":  "https://ltgdwhi.adr.admin.ch/gdwh-api/v2/",
    "PROD": "https://ltgdwh.adr.admin.ch/gdwh-api/v2/",
}


def gdwh_get_imports(base_url: str, gds_key: str) -> List[Dict]:
    """Holt alle DataPackages (Imports) für einen GDS-Key."""
    url = f"{base_url}api/geodatasets/{gds_key}/data/imports"
    with _gdwh_session() as s:
        r = s.get(url, timeout=(30, 300))
    r.raise_for_status()
    data = r.json()
    if isinstance(data, list):
        return data
    for key in ("items", "imports", "datapackages", "results", "data"):
        if key in data and isinstance(data[key], list):
            return data[key]
    return [data] if data else []


def gdwh_import_id(imp: Dict) -> str:
    """Extrahiert die DataPackage-ID (UUID) aus einem Import-Objekt."""
    for key in ("uuid", "id", "datapackageId", "package_id", "importId"):
        if imp.get(key):
            return str(imp[key])
    return "?"


def gdwh_import_date(imp: Dict) -> str:
    """Extrahiert und kürzt das Datum eines Imports."""
    for key in ("importDate", "date", "created_at", "createdAt", "timestamp", "created"):
        val = imp.get(key)
        if val:
            return str(val)[:16].replace("T", " ")
    return "–"


# XML-Feldnamen für Metadaten-Extraktion (Vergleich erfolgt lowercase)
_XML_AREA_TAGS        = ("area",)
_XML_LINEID_TAGS      = ("line_id", "lineid")
_XML_COMMENTARY_TAGS  = ("commentary", "kommentar", "comment", "description")
_XML_AUFTRAGSTYP_TAGS = ("auftragstyp", "auftragstype", "ordertype", "type")
_XML_DATETIME_TAGS    = ("stacitemiddatetime", "stac_item_id_datetime",
                         "stacdatetime", "acquisitiondate", "datetime")


def _find_xml_value(root: ET.Element, tags) -> str:
    """Sucht namespace-agnostisch nach dem ersten passenden Tag (alle Ebenen)."""
    tag_set = set(t.lower() for t in tags)
    for el in root.iter():
        # Namespace-Präfix entfernen: '{http://...}Tag' → 'tag'
        local = re.sub(r"^\{[^}]+\}", "", el.tag).lower()
        if local in tag_set and el.text and el.text.strip():
            return el.text.strip()
    return ""


# ─── FileMetadata-Suche (Metadatenquelle für Area/Jahr/Auftragstyp) ─────────
#
# GET /data/imports liefert nur uuid/gdsKey/importDate/footprint – keine
# fachlichen Attribute. POST /fileMetadata/search trägt die eigentlichen
# Attribute über ein customAttributes-XML-Fragment. Der Join zu einem Import
# läuft über importUuid == uuid.

def gdwh_search_file_metadata(base_url: str, gds_key: str) -> List[Dict]:
    """Holt die aktuellen FileMetadata-Einträge (mostRecent) für einen GDS-Key."""
    url = f"{base_url}api/geodatasets/{gds_key}/fileMetadata/search"
    payload = {"gdsKey": gds_key, "mostRecent": True}
    with _gdwh_session() as s:
        r = s.post(url, json=payload, timeout=(30, 300))
    r.raise_for_status()
    data = r.json()
    return data if isinstance(data, list) else []


def _parse_custom_attributes(xml_fragment: str) -> Dict:
    """Parst das customAttributes-XML-Fragment eines FileMetadata-Eintrags.

    Kein eigenständiges XML-Dokument (mehrere Wurzel-Tags aneinandergereiht),
    daher künstliches Wurzelelement zum Parsen nötig."""
    result = {"area": "", "line_id": "", "commentary": "", "auftragstyp": "", "stac_datetime": ""}
    if not xml_fragment:
        return result
    try:
        root = ET.fromstring(f"<root>{xml_fragment}</root>")
    except ET.ParseError:
        return result
    result["area"]          = _find_xml_value(root, _XML_AREA_TAGS)
    result["line_id"]       = _find_xml_value(root, _XML_LINEID_TAGS)
    result["commentary"]    = _find_xml_value(root, _XML_COMMENTARY_TAGS)
    result["auftragstyp"]   = _find_xml_value(root, _XML_AUFTRAGSTYP_TAGS)
    result["stac_datetime"] = _find_xml_value(root, _XML_DATETIME_TAGS)
    return result


def gdwh_index_file_metadata_by_import(file_metadata: List[Dict]) -> Dict[str, Dict]:
    """Baut ein Lookup importUuid → angereicherte Metadaten (Area, Jahr, LineID, …).

    Ein Import kann mehrere FileMetadata-Einträge haben (z.B. je Kachel) – für
    die Anzeige genügt ein repräsentativer Eintrag pro Import, der erste
    Treffer wird verwendet."""
    index: Dict[str, Dict] = {}
    for fm in file_metadata:
        import_uuid = fm.get("importUuid")
        if not import_uuid or import_uuid in index:
            continue
        attrs = _parse_custom_attributes(fm.get("customAttributes", ""))
        file_format = fm.get("fileFormat") or {}
        index[import_uuid] = {
            "area":          attrs["area"],
            "line_id":       attrs["line_id"],
            "commentary":    attrs["commentary"] or fm.get("commentary", ""),
            "auftragstyp":   attrs["auftragstyp"],
            "stac_datetime": attrs["stac_datetime"],
            "year":          str(fm.get("temporalKey") or ""),
            "tile_key":      fm.get("tileKey", ""),
            "file_format":   file_format.get("name", ""),
            "file_extension": file_format.get("extension", ""),
        }
    return index
