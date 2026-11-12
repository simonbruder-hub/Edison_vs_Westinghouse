"""Gemeinsames Quellenverzeichnis und IEEE-Nummerierung für Lösung, Playbook und Folien.

Das LaTeX-Handout (latex/handout/main.tex) legt die Nummern fest: Reihenfolge der ersten Nennung.
Die fertig formatierten Einträge stehen in scripts/quellen.json (deutsches IEEE-Format wie im Handout),
die Daten für LaTeX in latex/handout/quellen.bib. Beide verwenden dieselben Schlüssel.
Markup im Text: {schluessel} oder {a,b} wird zu [n] bzw. [n], [m].

Aufruf zur Kontrolle: python3 scripts/quellen.py
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REFS = json.loads((ROOT / 'scripts' / 'quellen.json').read_text(encoding='utf-8'))
HANDOUT = ROOT / 'latex' / 'handout' / 'main.tex'
BIB = ROOT / 'latex' / 'handout' / 'quellen.bib'


def reihenfolge_im_handout():
    """Schlüssel in der Reihenfolge der ersten Nennung im LaTeX-Handout (\\cite und \\citeoffen)."""
    text = '\n'.join(re.sub(r'(?<!\\)%.*', '', zeile) for zeile in HANDOUT.read_text(encoding='utf-8').splitlines())
    order = []
    for m in re.finditer(r'\\cite(?:offen)?(?:\[[^\]]*\])*\{([^}]*)\}', text):
        for key in (k.strip() for k in m.group(1).split(',')):
            if key and key not in order:
                order.append(key)
    return order


class Zitate:
    def __init__(self, order=None):
        self.fixed = order is not None
        self.order = list(order or [])
        self.used = []

    def n(self, *keys):
        nums = []
        for k in keys:
            if k not in REFS:
                raise KeyError(f'Unbekannte Quelle: {k}')
            if k not in self.order:
                if self.fixed:
                    raise KeyError(f'Quelle {k} wird im Handout nicht zitiert; zuerst dort aufnehmen.')
                self.order.append(k)
            if k not in self.used:
                self.used.append(k)
            nums.append(self.order.index(k) + 1)
        nums = sorted(set(nums))
        parts, i = [], 0
        while i < len(nums):
            j = i
            while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
                j += 1
            parts += [f'[{nums[i]}]–[{nums[j]}]'] if j - i >= 2 else [f'[{x}]' for x in nums[i:j + 1]]
            i = j + 1
        return ', '.join(parts)

    def text(self, s):
        return re.sub(r'\{([A-Za-z0-9,]+)\}', lambda m: self.n(*m.group(1).split(',')), s)

    def eintraege(self, nur_zitierte=False):
        """(Nummer, IEEE-Text) in Reihenfolge der Nummern."""
        keys = [k for k in self.order if not nur_zitierte or k in self.used]
        return [(self.order.index(k) + 1, REFS[k]) for k in keys]


def gemeinsame_zitate():
    """Nummern wie im Handout."""
    order = reihenfolge_im_handout()
    fehlend = [k for k in order if k not in REFS]
    if fehlend:
        raise KeyError(f'In scripts/quellen.json fehlen: {", ".join(fehlend)}')
    return Zitate(order)


if __name__ == '__main__':
    order = reihenfolge_im_handout()
    bib = set(re.findall(r'^@\w+\{([^,\s]+),', BIB.read_text(encoding='utf-8'), re.M))
    print(len(order), 'Quellen im Handout:')
    for nr, key in enumerate(order, 1):
        print(f'  [{nr}] {key}')
    probleme = [f'{k}: fehlt in quellen.bib' for k in order if k not in bib]
    probleme += [f'{k}: fehlt in quellen.json' for k in order if k not in REFS]
    probleme += [f'{k}: steht in quellen.bib, wird aber nicht zitiert' for k in sorted(bib - set(order))]
    print('\n'.join(probleme) if probleme else 'Schlüssel in main.tex, quellen.bib und quellen.json stimmen überein.')
    raise SystemExit(1 if probleme else 0)
