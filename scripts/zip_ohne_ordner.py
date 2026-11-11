"""Entfernt Ordner-Einträge aus einer Office-Datei (ZIP). Solche Einträge stören OneDrive und SharePoint.

Aufruf: python3 scripts/zip_ohne_ordner.py datei.docx
"""
import sys
import zipfile
from pathlib import Path

pfad = Path(sys.argv[1])
with zipfile.ZipFile(pfad) as z:
    inhalt = [(i, z.read(i)) for i in z.infolist() if not i.is_dir()]
with zipfile.ZipFile(pfad, 'w', zipfile.ZIP_DEFLATED) as z:
    for info, daten in inhalt:
        z.writestr(info.filename, daten)
