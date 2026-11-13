# Der Stromkrieg – Referat «Weltmacht USA» (FHNW)

Referat von Simon Bruder, Modul «Weltmacht USA» (Dozent: Stephan Schwarz), 10. Dezember 2026.

## Abgabe (`abgabe/`)

- `Handout_Stromkrieg_Bruder.pdf` – Handout im FHNW-Layout (2 Seiten Inhalt, 1 Seite Glossar, Tipps und Quellen) mit
  dem Schwerpunkt Edison gegen Tesla: Überblick mit Leitfrage und These, Edison und Tesla im Vergleich, Etappen, Gründe,
  Bedeutung für die USA mit Grafik, Mythen, «Und heute?», Fazit.
  Rote «S. ?» markieren Seitenzahlen, die noch nachzutragen sind.
- `Handout_Stromkrieg_Overleaf.zip` – dasselbe Handout als LaTeX-Projekt. In Overleaf: «Neues Projekt → Projekt
  hochladen». Der Text steht in `main.tex`, die Quellen in `quellen.bib`.
- `Praesentation_Stromkrieg_Bruder.pptx` / `.pdf` – Präsentation ohne Diskussion (21 Folien, 10 Minuten) auf der
  Slidesgo-Vorlage «Reconstruction Era and the Gilded Age», mit Sprechernotizen, Animationen und einem kurzen Video
  auf der Frontfolie
- `Praesentation_Stromkrieg_Bruder_mit_Diskussion.pptx` / `.pdf` – dieselbe Präsentation mit drei Fragen ans Plenum
  (25 Folien); die möglichen Antworten erscheinen per Klick
- `Ausarbeitung_Stromkrieg_lang.docx` – ausführliche Ausarbeitung (34 Quellen) zur eigenen Vorbereitung

Handout und Folien verwenden dieselben Quellennummern (IEEE nach dem FHNW-Leitfaden, 21 Quellen, darunter drei
Primärquellen von Edison und Tesla). Die lange
Ausarbeitung hat eine eigene Nummerierung.

Schriften der Folien: Libre Baskerville und Source Sans 3 (Google Fonts) vor dem Präsentieren installieren. Morph,
Animationen und Video laufen in PowerPoint 2019 oder Microsoft 365. Die PDFs zeigen den Endzustand jeder Folie.

## Vorbereitung (`vorbereitung/`)

- `Playbook_und_Vorbereitung_Stromkrieg.docx` – Ablauf Folie für Folie, Hinweise zu Video und Bildern, Checkliste,
  **Stand der Quellen** (was geprüft ist und was noch zu prüfen ist), Prüfprotokoll, mögliche Fragen, Lese- und
  Schauliste mit Zeitplan

Gelb markierte Stellen im Literaturverzeichnis der langen Ausarbeitung sind Platzhalter, die noch ergänzt werden müssen.

## Lokaler Ordner ohne GitHub

`Stromkrieg_aktualisieren.cmd` doppelklicken (Windows). Das Skript lädt die Ordner `abgabe` und `vorbereitung`
nach `C:\Users\simon\OneDrive - FHNW\7EIT\wusa\Stromkrieg_Referat` und legt sich beim ersten Lauf selbst dort ab.
Danach genügt ein Doppelklick im Ordner, um die neueste Version zu holen. Selbst geänderte Dateien werden vorher
nach `_Sicherung\<Datum>` kopiert. Kein Git nötig. Weil das Repository privat ist, öffnet das Skript den Download im
Browser (dort bei GitHub angemeldet sein); nach dem Download Enter drücken, dann übernimmt es die ZIP-Datei aus dem
Ordner «Downloads».

## Neu erzeugen

```bash
bash scripts/bauen.sh
```

Voraussetzungen: TeX Live mit `latexmk` und `biber` (oder Overleaf), Node.js mit den Paketen `docx` und `playwright`
(Chromium, nur für die Grafik der Folien), Python 3 mit `python-pptx`, `lxml`, `numpy` und `Pillow`, `ffmpeg` (nur für
das Video), die beiden Schriften und für die PDFs der Folien LibreOffice.

| Datei | Zweck |
| --- | --- |
| `latex/handout/main.tex`, `einstellungen.tex`, `quellen.bib`, `bilder/` | Handout; die Reihenfolge der Quellen im Text legt die Nummern fest |
| `scripts/quellen.json` | alle Quellen im IEEE-Format für Folien, Playbook und Ausarbeitung (dieselben Schlüssel wie in `quellen.bib`) |
| `scripts/quellen.py` | liest die Reihenfolge aus `main.tex`; `python3 scripts/quellen.py` prüft, ob alle Schlüssel übereinstimmen |
| `scripts/titelvideo.py` | erzeugt `assets/titelvideo.mp4` und das Standbild aus dem Stich von 1889 |
| `scripts/praesentation_vorlage.py`, `animation.py` | füllt `assets/folien_struktur.pptx` und erzeugt beide Foliensätze samt Animationen und Video |
| `scripts/chart.js` | Grafik «Anteil an der Weltindustrieproduktion» für die Folien |
| `scripts/playbook.py`, `fhnw_docx.py` | Playbook im FHNW-Layout (`assets/vorlage_fhnw.docx`) |
| `scripts/handout.js` | lange Ausarbeitung |

In den Skripten steht eine Quelle als `{schluessel}`; daraus wird beim Erzeugen die Nummer `[n]`. Eine Quelle, die im
Handout fehlt, lässt den Lauf der Folien abbrechen. So bleiben die Nummern überall gleich.

`assets/folien_struktur.pptx` enthält 25 Folien der Vorlage. Das Skript füllt sie, ergänzt die Folie mit den zwei
Zitaten, die Folie mit der Karikatur und eine vierte Quellenfolie, lässt drei nicht mehr gebrauchte Folien weg und legt
die Reihenfolge fest: Frontfolie, Titel, Einstieg, Leitfrage, Zwei Erfinder, Zwei Blicke (Zitate), Technik, Verlauf 1
und 2, Gründe, Grafik, Kennzahlen, Karikatur, Mythen, «Und heute?», Fazit, Diskussion (4 Folien), Danke, Quellen und
Bildnachweis (4 Folien). Für die Fassung ohne Diskussion
entfernt es die vier Diskussionsfolien.

Die drei historischen Bilder stammen aus National Geographic History; die Bildnachweise stehen am Bild. Sie sind nur
für Vortrag und Handout im Unterricht gedacht.
