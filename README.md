# Der Stromkrieg – Referat «Weltmacht USA» (FHNW)

Referat von Simon Bruder, Modul «Weltmacht USA» (Dozent: Stephan Schwarz), 10. Dezember 2026.

## Abgabe (`abgabe/`)

- `Handout_Stromkrieg_Bruder.docx` – Handout im FHNW-Layout (3 Seiten Inhalt, 1 Seite Quellen): Einleitung mit Leitfrage
  und These, Begriffe, Etappen, Bedeutung für die USA mit Grafik, Mythen, «Und heute?», Gruppendiskussion
- `Loesung_Handout_Stromkrieg_Bruder.docx` – mögliche Antworten zu allen Fragen und Gruppenaufträgen des Handouts
- `Ausarbeitung_Stromkrieg_lang.docx` – ausführliche Ausarbeitung (34 Quellen) zur eigenen Vorbereitung
- `Praesentation_Stromkrieg_Bruder.pptx` / `.pdf` – Präsentation ohne Diskussion (18 Folien, 10 Minuten) auf der
  Slidesgo-Vorlage «Reconstruction Era and the Gilded Age», mit Sprechernotizen, Morph-Intro und Animationen
- `Praesentation_Stromkrieg_Bruder_mit_Diskussion.pptx` / `.pdf` – dieselbe Präsentation mit Diskussionsblock
  (25 Folien): Fragen und Gruppenaufträge des Handouts; die Lösungen erscheinen per Klick

Handout, Lösung und Folien verwenden dieselben Quellennummern (IEEE, 18 Quellen). Die lange Ausarbeitung hat eine
eigene Nummerierung.

Schriften der Folien: Libre Baskerville und Source Sans 3 (Google Fonts) vor dem Präsentieren installieren. Morph und
Animationen laufen in PowerPoint 2019 oder Microsoft 365. Die PDFs zeigen den Endzustand jeder Folie, ohne die
Intro-Folie, die nur dem Morph-Übergang dient.

## Vorbereitung (`vorbereitung/`)

- `Playbook_und_Vorbereitung_Stromkrieg.docx` – Ablauf Folie für Folie, Checkliste, **offene Punkte vor der Abgabe**
  (Seitenzahlen nachtragen, Online-Quellen prüfen), mögliche Fragen, Lese- und Schauliste mit Zeitplan

Gelb markierte Stellen im Literaturverzeichnis der langen Ausarbeitung sind Platzhalter, die noch ergänzt werden müssen.

## Lokaler Ordner ohne GitHub

`Stromkrieg_aktualisieren.cmd` doppelklicken (Windows). Das Skript lädt die Ordner `abgabe` und `vorbereitung`
nach `C:\Users\simon\OneDrive - FHNW\7EIT\wusa\Stromkrieg_Referat` und legt sich beim ersten Lauf selbst dort ab.
Danach genügt ein Doppelklick im Ordner, um die neueste Version zu holen. Selbst geänderte Dateien werden vorher
nach `_Sicherung\<Datum>` kopiert. Kein Git und kein Login nötig.

## Neu erzeugen

```bash
bash scripts/bauen.sh
```

Voraussetzungen: Node.js mit den Paketen `docx` und `playwright` (Chromium, nur für die Grafik), Python 3 mit
`python-pptx` und `lxml`, die beiden Schriften und für die PDFs LibreOffice.

| Datei | Zweck |
| --- | --- |
| `scripts/quellen.json` | alle Quellen im IEEE-Format (eine Stelle für alle Dokumente) |
| `scripts/quellen.py` | Nummerierung: Das Handout legt die Reihenfolge fest (`assets/zitierreihenfolge.json`), Lösung und Folien übernehmen sie |
| `scripts/chart.js` | Grafik «Anteil an der Weltindustrieproduktion» für Word und Folien |
| `scripts/fhnw_docx.py` | Bausteine für Word-Dokumente im FHNW-Layout (`assets/vorlage_fhnw.docx`) |
| `scripts/handout_kurz.py`, `loesung_handout.py`, `playbook.py` | Handout, Lösung, Playbook |
| `scripts/handout.js` | lange Ausarbeitung |
| `scripts/praesentation_vorlage.py`, `animation.py` | füllt `assets/folien_struktur.pptx` und erzeugt beide Foliensätze samt Animationen |

Im Text steht eine Quelle als `{schluessel}`; daraus wird beim Erzeugen die Nummer `[n]`. Eine Quelle, die im Handout
fehlt, lässt den Lauf der Folien abbrechen. So bleiben die Nummern überall gleich.

`assets/folien_struktur.pptx` enthält die 25 Folien der Vorlage in der Reihenfolge des Referats: Intro, Titel,
Einstieg, Leitfrage, Technik, Firmen, Verlauf 1 und 2, Gründe, Grafik, Kennzahlen, Mythen, «Und heute?», Fazit,
Diskussion (7 Folien), Danke, Quellen (3 Folien). Für die Fassung ohne Diskussion entfernt das Skript die sieben
Diskussionsfolien.
