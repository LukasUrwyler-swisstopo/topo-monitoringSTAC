"""
test_functions.py  –  Unit-Tests für die reinen Hilfsfunktionen von
stac_api.py und 0_GUI_stac_monitor.py (kein Netzwerk-/GUI-Zugriff).

Aufruf:  pytest test_functions.py
"""

import importlib.util
from pathlib import Path

import pytest

import stac_api as api

# Modulname beginnt mit einer Ziffer ("0_GUI_..."), daher kein regulärer
# Import möglich – Laden über importlib anhand des Dateipfads.
_gui_path = Path(__file__).parent / "0_GUI_stac_monitor.py"
_spec = importlib.util.spec_from_file_location("gui_stac_monitor", _gui_path)
gui = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gui)


# ─── stac_api.browser_url ──────────────────────────────────────────────────

def test_browser_url_collection_int():
    url = api.browser_url("INT")
    assert url.startswith("https://sys-data.int.bgdi.ch/#/collections/")
    assert api.COLLECTION_ID in url
    assert "/items/" not in url


def test_browser_url_item_prod():
    url = api.browser_url("PROD", "kry-2025-09-19t09470000")
    assert url.startswith("https://data.geo.admin.ch/browser/index.html#/collections/")
    assert "/items/kry-2025-09-19t09470000" in url
    assert url.endswith("?.language=en")


# ─── stac_api.filter_items ─────────────────────────────────────────────────

def test_filter_items_empty_term_returns_all():
    items = [{"id": "kry-a"}, {"id": "ram-b"}]
    assert api.filter_items(items, "") == items


def test_filter_items_case_insensitive_substring():
    items = [{"id": "KRY-2025-09-19T09470000"}, {"id": "ram-2025-01-01t00000000"}]
    result = api.filter_items(items, "kry")
    assert result == [items[0]]


# ─── stac_api.stac_item_acq_date / stac_item_year ──────────────────────────

def test_acq_date_from_item_id():
    item = {"id": "kry-2025-09-19t09470000"}
    assert api.stac_item_acq_date(item) == "2025-09-19"


def test_acq_date_fallback_to_properties_datetime():
    item = {"id": "kry-ohne-datum", "properties": {"datetime": "2025-09-19T09:47:00Z"}}
    assert api.stac_item_acq_date(item) == "2025-09-19T09:47:00Z"


def test_item_year_from_item_id():
    assert api.stac_item_year({"id": "kry-2025-09-19t09470000"}) == "2025"


def test_item_year_fallback_to_properties():
    item = {"id": "kry-ohne-jahr", "properties": {"datetime": "2024-01-01T00:00:00Z"}}
    assert api.stac_item_year(item) == "2024"


def test_item_year_missing_returns_empty_string():
    assert api.stac_item_year({"id": "kry-ohne-jahr"}) == ""


# ─── stac_api.parse_asset_description / asset_area ─────────────────────────

def test_parse_asset_description_typical():
    desc = ("Area: RANDA, TerrainModel: DTM, "
             "Acquisition time: t1,t2,t3, LineId: L01,L02, Commentary: ok")
    result = api.parse_asset_description(desc)
    assert result["Area"] == "RANDA"
    assert result["TerrainModel"] == "DTM"
    # Werte mit eingebetteten Kommas dürfen nicht am Komma zerschnitten werden.
    assert result["Acquisition time"] == "t1,t2,t3"
    assert result["LineId"] == "L01,L02"
    assert result["Commentary"] == "ok"


def test_parse_asset_description_empty():
    assert api.parse_asset_description("") == {}


def test_asset_area_present():
    asset = {"description": "Area: BIRCH BLATTEN, TerrainModel: DSM"}
    assert api.asset_area(asset) == "BIRCH BLATTEN"


def test_asset_area_missing():
    assert api.asset_area({"description": "TerrainModel: DSM"}) == ""
    assert api.asset_area({}) == ""


# ─── stac_api.stac_item_area ───────────────────────────────────────────────

def test_item_area_from_properties():
    item = {"properties": {"area": "randa"}}
    assert api.stac_item_area(item) == "RANDA"


def test_item_area_fallback_to_asset_description():
    item = {
        "properties": {},
        "assets": {"nrgb.tif": {"description": "Area: Birch Blatten, TerrainModel: DSM"}},
    }
    assert api.stac_item_area(item) == "BIRCH BLATTEN"


def test_item_area_none_found():
    item = {"properties": {}, "assets": {}}
    assert api.stac_item_area(item) == ""


# ─── 0_GUI_stac_monitor._fmt_size ──────────────────────────────────────────

def test_fmt_size_none():
    assert gui._fmt_size(None) == "–"


def test_fmt_size_bytes():
    assert gui._fmt_size(512) == "512 B"


def test_fmt_size_kb():
    assert gui._fmt_size(2048) == "2.0 KB"


def test_fmt_size_mb():
    assert gui._fmt_size(5 * 1024 ** 2) == "5.0 MB"


def test_fmt_size_gb():
    assert gui._fmt_size(3 * 1024 ** 3) == "3.00 GB"


# ─── 0_GUI_stac_monitor._fmt_date ──────────────────────────────────────────

def test_fmt_date_none():
    assert gui._fmt_date(None) == "–"


def test_fmt_date_valid_http_header():
    assert gui._fmt_date("Fri, 19 Sep 2025 09:47:00 GMT") == "2025-09-19"


def test_fmt_date_unparsable_fallback():
    assert gui._fmt_date("2025-09-19-irgendwas") == "2025-09-19"


# ─── 0_GUI_stac_monitor._status_label ──────────────────────────────────────

def test_status_label_none():
    text, tag = gui._status_label(None)
    assert tag == "asset_dim"


def test_status_label_ok():
    text, tag = gui._status_label(200)
    assert "200" in text
    assert tag == "asset_ok"


def test_status_label_http_error():
    text, tag = gui._status_label(404)
    assert "404" in text
    assert tag == "asset_err"


def test_status_label_timeout():
    text, tag = gui._status_label(-2)
    assert tag == "asset_warn"


def test_status_label_other_error():
    text, tag = gui._status_label(-3)
    assert tag == "asset_warn"


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-v"]))
