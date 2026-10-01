# Der Stromkrieg – Referat «Weltmacht USA» (FHNW)

Abgabe in `abgabe/`:

- `Handout_Stromkrieg_Bruder.docx` – Handout (A4, Arial 11, Zitierweise IEEE)
- `Praesentation_Stromkrieg_Bruder.pptx` – Präsentation auf Basis der Slidesgo-Vorlage «Reconstruction Era and the Gilded Age»,
  13 Folien inkl. Titel, Danke und Quellen-Anhang, mit Sprechernotizen (10 Minuten).
  Schriften der Vorlage: Libre Baskerville und Source Sans 3 (Google Fonts) vor dem Präsentieren installieren.

Gelb markierte Stellen im Literaturverzeichnis sind Platzhalter, die noch ergänzt werden müssen.

## Neu erzeugen

Benötigt Node.js mit den Paketen `docx`, `pptxgenjs`, `react`, `react-dom`, `react-icons`, `sharp`
und `playwright` (nur für die Grafik).

```bash
node scripts/chart.js assets/chart_weltindustrie.png   # Grafik Weltindustrieproduktion
node scripts/handout.js                                # Handout + Zitierreihenfolge
node scripts/chart_vorlage.js assets/chart_weltindustrie_vorlage.png
python3 scripts/praesentation_vorlage.py struct.pptx abgabe/Praesentation_Stromkrieg_Bruder.pptx
```

Quellen stehen zentral in `scripts/quellen.js`; die Nummern [n] ergeben sich aus der Reihenfolge
der ersten Zitierung im Handout und werden in den Folien übernommen.

`struct.pptx` ist die Slidesgo-Vorlage mit den Folien 1, 4, 17, 11, 25, 25 (Kopie), 9, 29, 19, 30, 6, 35, 6 (Kopie)
in dieser Reihenfolge; alle übrigen Vorlagenfolien sind entfernt.
