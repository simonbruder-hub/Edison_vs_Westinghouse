"""Gemeinsames Quellenverzeichnis (scripts/quellen.json) und IEEE-Nummerierung.

Das Handout legt die Nummern fest (Reihenfolge der ersten Nennung) und schreibt sie nach
assets/zitierreihenfolge.json. Folien und Lösung lesen diese Datei und verwenden dieselben Nummern.
Markup im Text: {schluessel} oder {a,b} wird zu [n] bzw. [n], [m].
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REFS = json.loads((ROOT / 'scripts' / 'quellen.json').read_text(encoding='utf-8'))
ORDER_FILE = ROOT / 'assets' / 'zitierreihenfolge.json'


def typografisch(text):
    """IEEE-Einträge: gerade Anführungszeichen zu typografischen."""
    return re.sub(r'"([^"]*)"', '“\\1”', text)


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
        return [(self.order.index(k) + 1, typografisch(REFS[k])) for k in keys]

    def speichern(self):
        ORDER_FILE.write_text(json.dumps(self.order, indent=2), encoding='utf-8')


def gemeinsame_zitate():
    return Zitate(json.loads(ORDER_FILE.read_text(encoding='utf-8')))
