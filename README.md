# STAC Monitor – ch.swisstopo.spezialbefliegungen

Read-only Desktop-Tool (Tkinter) zur Überwachung der STAC-Collection
`ch.swisstopo.spezialbefliegungen` auf INT- und PROD-Umgebung von swisstopo /
BGDI. Keine Schreib- oder Löschfunktionen.

## GUI

```bash
python 0_GUI_stac_monitor.py
```

<img width="440" height="559" alt="image" src="https://github.com/user-attachments/assets/74401204-eb8a-4f45-9bf8-1edf99763541" />



## Funktionen

- Items der Collection laden (alle, gefiltert oder per exakter Item-ID)
- Filter nach Auftragstyp (KRY / RAM), Jahr, Item-ID/Suchbegriff, Dateiendung
- Area-Spalte: AOI-Name aus Item-Properties bzw. aus der Asset-Description
  (`Area: ...`) erkannt und angezeigt, sowohl je Item als auch je Asset
- Checkboxen je Item/Asset ("Auswahl"-Spalte) zur Auswahl, was geprüft/exportiert
  wird; Item-Checkbox (de)selektiert alle zugehörigen Assets, Tri-State bei
  Teilauswahl. Buttons "Alle auswählen" / "Alles abwählen" für Massenauswahl
- Asset-Status-Prüfung via HTTP HEAD (Statuscode, Dateigrösse, Last-Modified),
  nur für ausgewählte Assets
- "Nur Fehler-Assets anzeigen": blendet nach einer Prüfung alle Assets mit
  Status 200 aus (Items ohne Fehler verschwinden ganz); leert beim Umschalten
  die bisherige Auswahl, aktiv erst nach dem ersten Prüflauf
- Statistik: OK / Fehler / Gesamtgrösse geprüfter Assets
- Export der Download-Links als JSON (für Kunden), inkl. Area je Asset, nur
  ausgewählte Assets
- Export der Asset-Tabelle als CSV (für interne Auswertung), nur ausgewählte Assets
- Export "Item - STAC Browser Links" als TXT: STAC-Browser-Link je Item plus
  Liste der ausgewählten Assets (für Kunden-Weitergabe)
- Item-JSON-Detailansicht, URL in Zwischenablage kopieren / im Browser öffnen
- STAC Browser öffnen (Collection- oder Item-Deep-Link, für Kunden-Weitergabe)
- "Viewer-Fenster öffnen" / "Link auf Kartenviewer": zeigen COG-Assets
  (.tif/.tiff) direkt in map.geo.admin.ch; bei EBO/EBN-Fotos (.jpg) stattdessen
  automatisch das zugehörige Tages-KML (Item mit fixer Zeit 23595900), da die
  Einzelfotos selbst nicht darstellbar sind. "Viewer-Fenster öffnen" ist ein
  angedocktes Kartenfenster rechts neben dem Hauptfenster (Chrome, ersatzweise
  Edge, im App-Modus, menülos, nur Pan/Zoom) – aktiv, sobald genau 1
  darstellbares Asset ausgewählt ist
- Hell/Dark-Theme

## Voraussetzungen

- Python 3.11+ (funktioniert auch mit 3.6)
- Paket `requests` (`pip install requests`)
- Tkinter (in der Standard-Windows-Python-Installation bereits enthalten)
- Für "Viewer-Fenster öffnen": Google Chrome oder Microsoft Edge (Chrome wird
  bevorzugt, falls beide vorhanden sind). Das Paket `pywin32` (Fenster-
  Positionierung) wird beim ersten Start automatisch per pip nachinstalliert,
  falls noch nicht vorhanden – kein manueller Schritt nötig. Schlägt das
  fehl (kein pip-/Internetzugriff), bleibt der Button deaktiviert bzw. es
  erscheint ein Hinweis

## Einrichtung

1. Zugangsdaten hinterlegen unter `secrets/stac_credentials.json`:

   ```json
   {
     "INT":  {"username": "...", "password": "..."},
     "PROD": {"username": "...", "password": "..."}
   }
   ```

2. Optional: `secrets/proxy_config.json` anpassen, falls ein anderer
   Firmenproxy als `proxy-bvcol.admin.ch:8080` verwendet wird.

Der Ordner `secrets/` ist in `.gitignore` ausgeschlossen und wird nicht
versioniert.

## Netzwerk / Proxy

Der Zugriff auf `sys-data.int.bgdi.ch` (INT) und `data.geo.admin.ch` (PROD)
erfolgt im Bundesnetz über den Proxy `proxy-bvcol.admin.ch:8080`
([stac_api.py](stac_api.py)). Ist dieser Proxy nicht auflösbar (z.B. auf
einem privaten Rechner ausserhalb des Bundesnetz), fällt das Tool nach dem
ersten fehlgeschlagenen Versuch automatisch auf eine Direktverbindung
zurück.

## Start

```bash
python 0_GUI_stac_monitor.py
```

1. Umgebung wählen (INT/PROD) und Credentials laden
2. Filter setzen (optional)
3. "Laden" – bei vollständiger Item-ID im Suchfeld Direct-Lookup, sonst
   automatischer Fallback auf Laden der gesamten Collection + Filter
4. Optional: Auswahl über die Checkbox-Spalte ("Export") anpassen – per
   Klick auf ein Item oder Asset, oder über "Alle auswählen"/"Alles abwählen"
5. Sektion "STAC-Funktionen": "Assets prüfen (HEAD)" für Status/Grösse/
   Last-Modified, danach Export als JSON, CSV oder "Item - STAC Browser Links"
   (jeweils nur Auswahl)

## Dateien

| Datei | Zweck |
|---|---|
| `0_GUI_stac_monitor.py` | GUI-Anwendung (Tkinter) |
| `stac_api.py` | STAC-API-Hilfsfunktionen (read-only) |
| `secrets/stac_credentials.json` | Zugangsdaten INT/PROD (nicht versioniert) |
| `secrets/proxy_config.json` | Proxy-Konfiguration (nicht versioniert) |
