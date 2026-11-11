"""Lösungen zum Handout (FHNW-Layout, Quellennummern wie im Handout).

Aufruf: python3 scripts/loesung_handout.py   (nach handout_kurz.py)
"""
from fhnw_docx import ROOT, build, heading, items, para, references
from quellen import gemeinsame_zitate

z = gemeinsame_zitate()
T = z.text

b = []
b.append(para('Der Stromkrieg – Lösungen zum Handout', b=True, sz=36, jc='center', after=60))
b.append(para('Mögliche Antworten. Andere gut begründete Antworten sind ebenfalls richtig.', jc='center', after=240))

b.append(heading('Einstiegsfragen (Seite 1)'))
b.append(items([
    '**Meist beides.** Das Netz liefert Wechselstrom. Elektronik, LED-Lampen und alles mit Akku (Handy, Laptop, E-Auto) '
    'arbeiten intern mit Gleichstrom; das Netzteil wandelt um. Direkt mit Wechselstrom laufen klassische Motoren und '
    'Heizgeräte, zum Beispiel der Wasserkocher. Solarzellen liefern Gleichstrom.',
    T('Ladestecker für E-Autos in den USA (Tesla-Stecker gegen CCS) {nacs}, USB-C gegen den Lightning-Stecker von Apple '
      '{eu}; früher VHS gegen Betamax bei Videokassetten oder Blu-ray gegen HD DVD.'),
]))
b.append(para('', after=120))

b.append(heading('Bedeutung für die Weltmacht USA (Seite 2)'))
b.append(items([
    T('Billiger Strom trieb Elektromotoren in den Fabriken an: 1899 lieferten sie knapp 5 %, 1929 rund 78 % der '
      'Antriebsleistung der US-Industrie {devine}. Fabriken konnten neu und effizienter gebaut werden, die Produktivität '
      'stieg. Mehr Produktion bedeutete mehr wirtschaftliche Macht (Abb. 1 im Handout).'),
    T('Ruhm allein reicht nicht. Entscheidend waren Kosten und Kapital: 1892 fusionierte Edison General Electric mit dem '
      'Rivalen Thomson-Houston zu General Electric. Edisons Name verschwand, die Führung übernahm Thomson-Houston, und GE '
      'setzte auf Wechselstrom {jonnes}. Das ist typisch für das Gilded Age: Banken und Grosskonzerne bestimmen den Markt, '
      'nicht einzelne Erfinder.'),
]))
b.append(para('', after=120))

b.append(heading('Diskussion in Gruppen (Seite 3)'))
for g, t in [
    ('Gruppe 1 – Angst als Werbung',
     'Mögliche Beispiele: Werbung, die ein Konkurrenzprodukt schlechtmacht; Angstkampagnen im Wahlkampf. Kernpunkt: Angst '
     'wirkt kurzfristig, setzt sich aber selten gegen einen klaren Nutzen und tiefere Kosten durch. Die Kampagne gegen den '
     'Wechselstrom ist dafür ein Beispiel: Westinghouse konterte sie und machte die Vorwürfe mit besserer Technik '
     'gegenstandslos {cole}.'),
    ('Gruppe 2 – Der Elefant Topsy',
     'Die Geschichte hat einen klaren Bösewicht und bleibt deshalb hängen. Es gibt sogar einen Film der Tötung von 1903, '
     'gedreht von Edisons Filmfirma. Das wirkt wie ein Beweis, obwohl es keinen Beleg gibt, dass Edison selbst beteiligt war '
     '{topsy}. Websites und sogar Firmen wie ABB übernehmen die Geschichte ungeprüft {abb}. Lehre: Primärquellen prüfen, '
     'zum Beispiel im Archiv der Edison Papers.'),
    ('Gruppe 3 – Standardkriege heute',
     'Beispiele: VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker, iOS gegen Android. Oft gewinnt nicht die technisch '
     'beste Lösung, sondern die mit mehr Partnern, tieferen Kosten und mehr Kapital. Hat sich ein Standard einmal '
     'durchgesetzt, bleibt er lange; der Fachbegriff dafür ist Pfadabhängigkeit {david92}. Beim Stromkrieg sprachen Technik '
     'und Kosten beide für den Wechselstrom.'),
]:
    b.append(para(g, b=True, after=40, keep=True))
    b.append(para(T(t), after=160))
b.append(para('', after=120))

b.append(heading('Quellen (Nummern wie im Handout)'))
b.append(references(z.eintraege(nur_zitierte=True)))

build(ROOT / 'abgabe' / 'Loesung_Handout_Stromkrieg_Bruder.docx', b, 'Der Stromkrieg – Lösungen zum Handout')
