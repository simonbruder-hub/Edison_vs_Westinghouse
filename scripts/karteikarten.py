"""Druckbare Karteikarten (A6 quer, vier pro A4-Seite) und Aussprache-Liste für den freien Vortrag.

Inhalt: scripts/karten.py (dieselben Stichpunkte wie in den Sprechernotizen der Folien).
Aufruf:  python3 scripts/karteikarten.py   → vorbereitung/Karteikarten_Stromkrieg.pdf (braucht latexmk)
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

from karten import AUSSPRACHE, DISKUSSION, HAUPTTEIL, KARTEN, TIPPS_AUSSPRACHE
from quellen import ROOT, gemeinsame_zitate

T = gemeinsame_zitate().text
ZIEL = ROOT / 'vorbereitung' / 'Karteikarten_Stromkrieg.pdf'


def tex(s):
    """Text für LaTeX: Sonderzeichen schützen, Pfeil als Formelzeichen."""
    for alt, neu in (('\\', r'\textbackslash{}'), ('&', r'\&'), ('%', r'\%'), ('$', r'\$'), ('#', r'\#'),
                     ('_', r'\_'), ('{', r'\{'), ('}', r'\}'), ('~', r'\textasciitilde{}'), ('^', r'\^{}')):
        s = s.replace(alt, neu)
    return s.replace('→', r'$\rightarrow$').replace('·', r'\textperiodcentered{}')


def karte(nr, k):
    punkte = '\n'.join(rf'\item {tex(T(p))}' for p in k['punkte'])
    weiter = rf'\par\vspace{{3mm}}{{\large$\rightarrow$ \emph{{{tex(k["weiter"])}}}}}' if k['weiter'] else ''
    return (rf'{{\footnotesize\color{{black!55}}Karte {nr} \textperiodcentered{{}} {tex(k["zeit"])}\par}}'
            rf'\vspace{{1mm}}{{\LARGE\bfseries {tex(k["folie"])}\par}}\vspace{{3mm}}'
            rf'{{\large\begin{{itemize}}[leftmargin=6mm,label=--,itemsep=2.2mm,topsep=0pt,parsep=0pt]{punkte}'
            rf'\end{{itemize}}}}'
            + weiter)


def aussprache(teil, nr, von):
    zeilen = '\n'.join(rf'{tex(a)} & {tex(b)} \\' for a, b in teil)
    tipps = ''
    if nr == von:
        tipps = (r'\par\vspace{2mm}{\small ' + ' \\\\ '.join(tex(x) for x in TIPPS_AUSSPRACHE) + r'\par}')
    return (rf'{{\footnotesize\color{{black!55}}Aussprache {nr}/{von} \textperiodcentered{{}} Grossbuchstaben = betonte '
            rf'Silbe\par}}\vspace{{1mm}}{{\Large\bfseries Aussprache\par}}\vspace{{2mm}}'
            rf'{{\small\renewcommand{{\arraystretch}}{{1.12}}\begin{{tabular}}{{@{{}}p{{58mm}}p{{70mm}}@{{}}}}{zeilen}'
            rf'\end{{tabular}}}}' + tipps)


def main():
    inhalte = [karte(i + 1, KARTEN[key]) for i, key in enumerate(HAUPTTEIL + DISKUSSION)]
    n = 3  # drei Karten; die Tipps stehen auf der letzten
    groesse = -(-len(AUSSPRACHE) // n)
    inhalte += [aussprache(AUSSPRACHE[i * groesse:(i + 1) * groesse], i + 1, n) for i in range(n)]

    ecken = [(0, 0), (148.5, 0), (0, 105), (148.5, 105)]  # links oben, rechts oben, links unten, rechts unten (mm)
    seiten = []
    for start in range(0, len(inhalte), 4):
        knoten = []
        for (x, y), inhalt in zip(ecken, inhalte[start:start + 4]):
            knoten.append(rf'\node[anchor=north west,inner sep=0,text width=132mm,align=left] at '
                          rf'($(current page.north west)+({x + 8}mm,-{y + 7}mm)$) {{{inhalt}}};')
        seiten.append(r'\null\begin{tikzpicture}[remember picture,overlay]'
                      r'\draw[black!35,dashed] (current page.north) -- (current page.south);'
                      r'\draw[black!35,dashed] (current page.west) -- (current page.east);'
                      + '\n'.join(knoten) + r'\end{tikzpicture}')

    dokument = (r'''\documentclass[12pt]{article}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[nswissgerman]{babel}
\usepackage{tgtermes}
\usepackage[a4paper,landscape,margin=0mm]{geometry}
\usepackage{xcolor,tikz,enumitem}
\usetikzlibrary{calc}
\pagestyle{empty}
\setlength{\parindent}{0pt}
\begin{document}
''' + '\n\\newpage\n'.join(seiten) + '\n\\end{document}\n')

    tmp = Path(tempfile.mkdtemp(prefix='karten_'))
    (tmp / 'karten.tex').write_text(dokument, encoding='utf-8')
    erg = subprocess.run(['latexmk', '-pdf', '-interaction=nonstopmode', '-halt-on-error', 'karten.tex'],
                         cwd=tmp, capture_output=True, text=True)
    if erg.returncode != 0:
        print((tmp / 'karten.log').read_text(encoding='utf-8', errors='replace')[-3000:])
        raise SystemExit('LaTeX-Fehler, siehe oben')
    ZIEL.parent.mkdir(exist_ok=True)
    shutil.copy(tmp / 'karten.pdf', ZIEL)
    shutil.rmtree(tmp)
    print('OK', ZIEL, len(inhalte), 'Karten')


if __name__ == '__main__':
    main()
