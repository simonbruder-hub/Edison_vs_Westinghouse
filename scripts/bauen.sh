#!/usr/bin/env bash
# Erzeugt alle Dateien in abgabe/ und vorbereitung/ neu. Voraussetzungen: siehe README.
set -euo pipefail
cd "$(dirname "$0")/.."

node scripts/chart.js                     # Grafik (weiss für Word, transparent für die Folien)
python3 scripts/handout_kurz.py           # Handout; legt die gemeinsamen Quellennummern fest
python3 scripts/loesung_handout.py        # Lösung (Nummern wie im Handout)
python3 scripts/playbook.py               # Playbook und Vorbereitungsliste
node scripts/handout.js                   # lange Ausarbeitung (eigene Nummerierung)
python3 scripts/praesentation_vorlage.py  # beide Foliensätze (Nummern wie im Handout)

# PDF der Folien ohne die Intro-Folie (sie dient nur dem Morph-Übergang); braucht LibreOffice
if command -v soffice >/dev/null; then
  for f in Praesentation_Stromkrieg_Bruder Praesentation_Stromkrieg_Bruder_mit_Diskussion; do
    n=$(python3 -c "from pptx import Presentation; import sys; print(len(Presentation(sys.argv[1]).slides))" "abgabe/$f.pptx")
    soffice --headless --convert-to "pdf:impress_pdf_Export:{\"PageRange\":{\"type\":\"string\",\"value\":\"2-$n\"}}" \
      --outdir "$PWD/abgabe" "$PWD/abgabe/$f.pptx" >/dev/null
  done
else
  echo "LibreOffice (soffice) nicht gefunden: PDFs der Folien nicht neu erzeugt." >&2
fi
