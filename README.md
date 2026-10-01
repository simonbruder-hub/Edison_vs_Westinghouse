# Der Stromkrieg – Referat «Weltmacht USA» (FHNW)

Abgabe in `abgabe/`:

- `Handout_Stromkrieg_Bruder.docx` – Handout (A4, Arial 11, Zitierweise IEEE)
- `Praesentation_Stromkrieg_Bruder.pptx` – 10 Folien + Anhang mit Sprechernotizen (10 Minuten)

Gelb markierte Stellen im Literaturverzeichnis sind Platzhalter, die noch ergänzt werden müssen.

## Neu erzeugen

Benötigt Node.js mit den Paketen `docx`, `pptxgenjs`, `react`, `react-dom`, `react-icons`, `sharp`
und `playwright` (nur für die Grafik).

```bash
node scripts/chart.js assets/chart_weltindustrie.png   # Grafik Weltindustrieproduktion
node scripts/handout.js                                # Handout + Zitierreihenfolge
node scripts/praesentation.js                          # Folien (nutzt dieselben Quellennummern)
```

Quellen stehen zentral in `scripts/quellen.js`; die Nummern [n] ergeben sich aus der Reihenfolge
der ersten Zitierung im Handout und werden in den Folien übernommen.
