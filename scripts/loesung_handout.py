"""Lösungen zum kurzen Handout (gleiches FHNW-Layout).

Aufruf: python3 scripts/loesung_handout.py <vorlage.docx> abgabe/Loesung_Handout_Stromkrieg_Bruder.docx
"""
import sys

from fhnw_docx import build, heading, items, para, references

src, out = sys.argv[1], sys.argv[2]

b = []
b.append(para('Der Stromkrieg – Lösungen zum Handout', b=True, sz=36, jc='center', after=60))
b.append(para('Mögliche Antworten. Andere gut begründete Antworten sind ebenfalls richtig.', jc='center', after=240))

b.append(heading('Einstiegsfragen (Seite 1)'))
b.append(items([
    '**Gleichstrom:** alles mit Batterie oder Akku, z. B. Handy, Laptop, E-Auto, Taschenlampe, dazu Solarzellen. '
    '**Wechselstrom:** alles, was direkt an der Steckdose hängt, z. B. Kühlschrank, Waschmaschine, Staubsauger. '
    'Gut zu wissen: Das Ladegerät des Handys wandelt den Wechselstrom aus der Steckdose in Gleichstrom um. '
    'Heute arbeiten beide Systeme zusammen.',
    'Ladestecker für E-Autos in den USA (Tesla-Stecker gegen CCS), USB-C gegen Apples Lightning-Stecker '
    '(Apple stellte 2023 um), früher VHS gegen Betamax bei Videokassetten oder Blu-ray gegen HD DVD (2008 entschieden).',
]))
b.append(para('', after=120))

b.append(heading('Bedeutung für die Weltmacht USA (Seite 2)'))
b.append(items([
    'Billiger Strom trieb Elektromotoren in den Fabriken an: 1899 lieferten sie knapp 5 %, 1929 rund 78 % der '
    'Antriebsleistung der US-Industrie [1]. Fabriken konnten neu und effizienter gebaut werden, die Produktivität stieg. '
    'An den Niagarafällen zog billige Wasserkraft neue Industrien an, z. B. Aluminium. Mehr Produktion bedeutete mehr '
    'wirtschaftliche Macht.',
    'Ruhm allein reicht nicht. Entscheidend waren Kosten und Kapital: 1892 fusionierte der Bankier J. P. Morgan Edisons '
    'Firma zu General Electric. Edison verlor die Kontrolle, und GE setzte danach selbst auf Wechselstrom. Das ist typisch '
    'für das Gilded Age: Banken und Grosskonzerne bestimmen den Markt, nicht einzelne Erfinder.',
]))
b.append(para('', after=120))

b.append(heading('Diskussion in Gruppen (Seite 3)'))
for g, t in [
    ('Gruppe 1 – Angst als Werbung',
     'Mögliche Beispiele: Werbung, die ein Konkurrenzprodukt schlechtmacht; Angstkampagnen im Wahlkampf; Debatten über '
     'Handystrahlung (5G). Kernpunkt: Angst wirkt kurzfristig, setzt sich aber selten gegen einen klaren Nutzen und '
     'tiefere Kosten durch. Edisons Kampagne ist dafür ein Beispiel [2].'),
    ('Gruppe 2 – Der Elefant Topsy',
     'Die Geschichte hat einen klaren Bösewicht und bleibt deshalb hängen. Es gibt sogar einen Film der Tötung von 1903, '
     'gedreht von Edisons Filmfirma. Das macht die Verbindung glaubwürdig, obwohl es keinen Beleg gibt, dass Edison selbst '
     'beteiligt war [3]. Websites und sogar Firmen wie ABB übernehmen die Geschichte ungeprüft. '
     'Lehre: Primärquellen prüfen, z. B. im Archiv der Edison Papers.'),
    ('Gruppe 3 – Standardkriege heute',
     'Beispiele: VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker, iOS gegen Android. Oft gewinnt nicht die technisch '
     'beste Lösung, sondern die mit mehr Partnern, tieferen Kosten und mehr Kapital. Hat sich ein Standard einmal '
     'durchgesetzt, bleibt er lange – der Fachbegriff dafür ist Pfadabhängigkeit [4]. Beim Stromkrieg sprachen Technik '
     'und Kosten beide für den Wechselstrom.'),
]:
    b.append(para(g, b=True, after=40, keep=True))
    b.append(para(t, after=160))
b.append(para('', after=120))

b.append(heading('Quellen'))
b.append(references([
    'W. D. Devine, Jr., “From shafts to wires: Historical perspective on electrification,” *The Journal of Economic History*, vol. 43, no. 2, pp. 347–372, 1983.',
    'B. M. Cole and D. Chandler, “A model of competitive impression management: Edison versus Westinghouse in the War of the Currents,” *Administrative Science Quarterly*, vol. 64, no. 4, pp. 1020–1063, 2019, doi: 10.1177/0001839218821439.',
    'Thomas A. Edison Papers, “Myth buster: Topsy the elephant,” Rutgers Univ. [Online]. Available: https://edison.rutgers.edu/life-of-edison/essaying-edison/essay/myth-buster-topsy-the-elephant',
    'P. A. David, “Heroes, herds and hysteresis in technological history: Thomas Edison and ‘the battle of the systems’ reconsidered,” *Industrial and Corporate Change*, vol. 1, no. 1, pp. 129–180, 1992.',
]))
b.append(para('Online-Quellen abgerufen am 1.10.2026.', sz=18, after=0))

build(src, out, b, 'Der Stromkrieg – Lösungen zum Handout')
