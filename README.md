# Der Stromkrieg – Referat «Weltmacht USA» (FHNW)

Abgabe in `abgabe/`:

- `Handout_Stromkrieg_Bruder.docx` – kurzes Handout (3 Seiten) im FHNW-Layout: Einleitung, Fragen, Begriffe,
  Etappen, Bedeutung, «Und heute?», Gruppendiskussion, Quellen (IEEE)
- `Loesung_Handout_Stromkrieg_Bruder.docx` – Musterlösungen zu allen Fragen und Gruppenaufträgen des Handouts (1 Seite)
- `Ausarbeitung_Stromkrieg_lang.docx` – ausführliche Hintergrund-Ausarbeitung (30 Quellen) zur eigenen Vorbereitung
- `Praesentation_Stromkrieg_Bruder.pptx` / `.pdf` – Präsentation ohne Diskussion (14 Folien, 10 Minuten) auf Basis der
  Slidesgo-Vorlage «Reconstruction Era and the Gilded Age», mit Sprechernotizen, Morph-Intro und Animationen
- `Praesentation_Stromkrieg_Bruder_mit_Diskussion.pptx` / `.pdf` – dieselbe Präsentation plus Diskussionsblock
  (21 Folien): Handout-Fragen und Gruppenaufträge, Lösungen erscheinen per Klick
  Schriften der Vorlage: Libre Baskerville und Source Sans 3 (Google Fonts) vor dem Präsentieren installieren.
  Morph und Animationen laufen in PowerPoint 2019 / Microsoft 365. Die PDFs zeigen den Endzustand jeder Folie
  (ohne die Intro-Folie, die nur für den Morph-Übergang da ist).

In `vorbereitung/`:

- `Playbook_und_Vorbereitung_Stromkrieg.docx` – Ablauf Folie für Folie, Checkliste, mögliche Fragen,
  Lese- und Schauliste mit Zeitplan

Gelb markierte Stellen im Literaturverzeichnis sind Platzhalter, die noch ergänzt werden müssen.

## Neu erzeugen

Benötigt Node.js mit den Paketen `docx`, `pptxgenjs`, `react`, `react-dom`, `react-icons`, `sharp`
und `playwright` (nur für die Grafik).

```bash
node scripts/chart.js assets/chart_weltindustrie.png   # Grafik Weltindustrieproduktion
node scripts/handout.js                                # lange Ausarbeitung + Zitierreihenfolge
python3 scripts/handout_kurz.py <eigenes-FHNW-Handout.docx> abgabe/Handout_Stromkrieg_Bruder.docx
python3 scripts/loesung_handout.py <eigenes-FHNW-Handout.docx> abgabe/Loesung_Handout_Stromkrieg_Bruder.docx
python3 scripts/playbook.py <eigenes-FHNW-Handout.docx> vorbereitung/Playbook_und_Vorbereitung_Stromkrieg.docx
node scripts/chart_vorlage.js assets/chart_weltindustrie_vorlage.png
python3 scripts/praesentation_vorlage.py struct_full.pptx abgabe/Praesentation_Stromkrieg_Bruder.pptx \
    abgabe/Praesentation_Stromkrieg_Bruder_mit_Diskussion.pptx
```

Quellen stehen zentral in `scripts/quellen.js`; die Nummern [n] ergeben sich aus der Reihenfolge
der ersten Zitierung im Handout und werden in den Folien übernommen.

`struct_full.pptx` ist die Slidesgo-Vorlage mit diesen Folien in dieser Reihenfolge (alle übrigen entfernt):
1 (Kopie, Intro), 1, 4, 17, 11, 25, 25 (Kopie), 9, 29, 19, 30, 6, 5, 6 (Kopie), 6 (Kopie), 9 (Kopie),
6 (Kopie), 6 (Kopie), 6 (Kopie), 35, 6 (Kopie). Die Folien 13–19 bilden den Diskussionsblock; für die Version
ohne Diskussion entfernt das Skript sie. Animationen und Übergänge erzeugt `scripts/animation.py`.
