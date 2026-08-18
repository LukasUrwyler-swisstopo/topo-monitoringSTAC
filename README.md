# STAC Monitor – ch.swisstopo.spezialbefliegungen

Desktop-Tool (Tkinter) zur Überwachung der STAC-Collection
`ch.swisstopo.spezialbefliegungen` auf INT- und PROD-Umgebung von swisstopo /
BGDI. Read-only gegenüber der STAC-API – es werden keine Items/Assets
gelöscht oder verändert.

## Schnellstart

1. **Einmalig:** Zugangsdaten hinterlegen, siehe [Einrichtung](#einrichtung)
2. Tool starten:

   ```bash
   python 0_GUI_stac_monitor.py
   ```

3. Im GUI oben: **Umgebung** (INT/PROD) wählen und **Credentials laden**
4. Optional Filter setzen (Auftragstyp, Jahr, Suchbegriff, Dateiendung),
   dann **Laden** klicken
5. In der Baumansicht die gewünschten Items/Assets per Checkbox auswählen
   (oder **Alle auswählen**)
6. Im Bereich **STAC-Funktionen**: **Assets prüfen (HEAD)** für Status &
   Grösse, danach je nach Bedarf herunterladen oder exportieren

<img width="440" height="559" alt="image" src="https://github.com/user-attachments/assets/74401204-eb8a-4f45-9bf8-1edf99763541" />

## Funktionen

- **Items laden & filtern** – ganze Collection oder gezielt per Item-ID;
  Filter nach Auftragstyp, Jahr, Suchbegriff, Dateiendung
- **Auswahl per Checkbox** – einzelne Assets oder ganze Items, inkl.
  "Alle auswählen" / "Alles abwählen"
- **Assets prüfen (HEAD)** – prüft Status, Dateigrösse und Änderungsdatum
  der ausgewählten Assets. Assets über 50 GB (von CloudFront normalerweise
  mit Fehler 400 gemeldet) werden korrekt als ✓ **>50GB** statt als Fehler
  erkannt
- **Fehlerhafte anzeigen** / **ITEMs ohne Thumbnail** – blenden die
  Baumansicht gezielt auf problematische bzw. unvollständige Items ein
- **ITEM and ASSET download** – lädt die ausgewählten Assets direkt auf die
  eigene Festplatte (ein Unterordner pro Item). Assets über 50 GB werden
  automatisch in Teilstücken heruntergeladen, da sie sonst am
  CloudFront-Limit scheitern würden
- **create Download-Links** – erstellt ein Textfile mit den Download-Links
  der Auswahl zum Weitergeben an Kunden; bei Assets über 50 GB inkl.
  Hinweis auf die nötige Download-Methode
- **Weitere Exporte** – Download-Links als JSON, Asset-Tabelle als CSV,
  STAC-Browser-Links als TXT
- **Kartenviewer** – ausgewählte GeoTIFFs bzw. Tagesübersichten direkt in
  map.geo.admin.ch anzeigen, wahlweise im Browser oder in einem
  angedockten Viewer-Fenster
- Statistik (OK/Fehler/Gesamtgrösse), Item-JSON-Detailansicht,
  Hell/Dark-Theme

## Voraussetzungen

- Python 3.11+ mit Tkinter (in der Standard-Windows-Installation enthalten)
- Paket `requests` (`pip install requests`)
- Für das angedockte Viewer-Fenster: Google Chrome oder Microsoft Edge.
  Das Paket `pywin32` wird beim ersten Start bei Bedarf automatisch
  nachinstalliert

## Einrichtung

Zugangsdaten hinterlegen unter `secrets/stac_credentials.json`:

```json
{
  "INT":  {"username": "...", "password": "..."},
  "PROD": {"username": "...", "password": "..."}
}
```

Optional: `secrets/proxy_config.json` anpassen, falls ein anderer
Firmenproxy als `proxy-bvcol.admin.ch:8080` verwendet wird. Der Ordner
`secrets/` ist in `.gitignore` ausgeschlossen und wird nicht versioniert.

Im Bundesnetz läuft der Zugriff automatisch über diesen Proxy; ausserhalb
(z.B. privater Rechner) schaltet das Tool selbstständig auf eine
Direktverbindung um.

## Dateien

| Datei | Zweck |
|---|---|
| `0_GUI_stac_monitor.py` | GUI-Anwendung (Tkinter) |
| `stac_api.py` | STAC-API-Hilfsfunktionen (inkl. Download) |
| `secrets/stac_credentials.json` | Zugangsdaten INT/PROD (nicht versioniert) |
| `secrets/proxy_config.json` | Proxy-Konfiguration (nicht versioniert) |
