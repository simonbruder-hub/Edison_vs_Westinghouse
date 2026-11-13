#!/usr/bin/env bash
# Erzeugt alle Dateien in abgabe/ und vorbereitung/ neu. Voraussetzungen: siehe README.
set -euo pipefail
cd "$(dirname "$0")/.."

python3 scripts/quellen.py >/dev/null      # bricht ab, wenn main.tex, quellen.bib und quellen.json nicht zusammenpassen

# Handout: PDF aus LaTeX (in einer Kopie, damit keine Hilfsdateien im Projekt landen) und ZIP für Overleaf
if command -v latexmk >/dev/null; then
  tmp=$(mktemp -d)
  cp -r latex/handout/. "$tmp/"
  (cd "$tmp" && latexmk -pdf -interaction=nonstopmode -halt-on-error main.tex >build.log 2>&1) \
    || { tail -n 30 "$tmp/build.log" >&2; exit 1; }
  cp "$tmp/main.pdf" abgabe/Handout_Stromkrieg_Bruder.pdf
  rm -rf "$tmp"
else
  echo "latexmk nicht gefunden: Handout-PDF nicht neu erzeugt (in Overleaf kompilieren)." >&2
fi
python3 - <<'EOF'
import zipfile
from pathlib import Path
quelle, ziel = Path('latex/handout'), Path('abgabe/Handout_Stromkrieg_Overleaf.zip')
with zipfile.ZipFile(ziel, 'w', zipfile.ZIP_DEFLATED) as z:
    for datei in sorted(quelle.rglob('*')):
        if datei.is_file():
            z.write(datei, datei.relative_to(quelle).as_posix())
EOF

node scripts/chart.js                                           # Grafik für die Folien
[ -f assets/titelvideo.mp4 ] || python3 scripts/titelvideo.py   # Video der Frontfolie, nur falls es fehlt
python3 scripts/praesentation_vorlage.py                        # beide Foliensätze (Nummern wie im Handout)
python3 scripts/playbook.py                                     # Playbook und Vorbereitungsliste
python3 scripts/karteikarten.py                                 # Karteikarten zum Ausdrucken (braucht latexmk)
node scripts/handout.js                                         # lange Ausarbeitung (eigene Nummerierung)

# PDFs der Folien; braucht LibreOffice
if command -v soffice >/dev/null; then
  for f in Praesentation_Stromkrieg_Bruder Praesentation_Stromkrieg_Bruder_mit_Diskussion; do
    soffice --headless --convert-to pdf --outdir "$PWD/abgabe" "$PWD/abgabe/$f.pptx" >/dev/null
  done
else
  echo "LibreOffice (soffice) nicht gefunden: PDFs der Folien nicht neu erzeugt." >&2
fi
