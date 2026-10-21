"""Kurzes Handout im Stil von Simons FHNW-Handout (Kopf-/Fusszeile und Logo aus dessen .docx).

Aufruf: python3 scripts/handout_kurz.py <vorlage.docx> abgabe/Handout_Stromkrieg_Bruder.docx
"""
import sys
from pathlib import Path

from fhnw_docx import PB, build, heading, image, items, para, references, table

ROOT = Path(__file__).resolve().parent.parent
src, out = sys.argv[1], sys.argv[2]

# ---------------- Inhalt ----------------
b = []
b.append(para('Der Stromkrieg – Handout', b=True, sz=36, jc='center', after=60))
b.append(para('Edison vs. Westinghouse und die Rolle von Nikola Tesla · Simon Bruder · Modul WUSA', sz=20, jc='center', after=240))
b.append(para('Um 1890 stritten in den USA zwei Firmen darum, wie das Land mit Strom versorgt wird. Thomas Edison setzte auf '
              '**Gleichstrom**, George Westinghouse mit den Erfindungen von Nikola Tesla auf **Wechselstrom**. Edisons Lager '
              'versuchte, Wechselstrom als lebensgefährlich darzustellen: Es tötete öffentlich Tiere und half, den elektrischen '
              'Stuhl einzuführen. Gewonnen hat trotzdem der Wechselstrom, weil er sich über weite Strecken billiger übertragen '
              'liess [1], [2]. Der Streit zeigt, wie Technik, Geld und Werbung den Aufstieg der USA zur Industriemacht prägten.',
              after=200))
b.append(items(['Welche Geräte in deinem Alltag laufen mit Gleichstrom, welche mit Wechselstrom?',
                'Kennst du ein aktuelles Beispiel, bei dem sich Firmen um einen technischen Standard streiten?']))
b.append(para('', after=120))
b.append(heading('Begriffe:'))
b.append(table([
    ['Gleichstrom (DC)', 'Strom, der immer in die gleiche Richtung fliesst, z. B. aus einer Batterie.'],
    ['Wechselstrom (AC)', 'Strom, der ständig die Richtung wechselt; so kommt er heute aus der Steckdose.'],
    ['Spannung (Volt)', 'Der «Druck», mit dem Strom durch eine Leitung fliesst. Hohe Spannung = weniger Verlust auf langen Strecken.'],
    ['Transformator', 'Gerät, das die Spannung von Wechselstrom hoch- oder herunterschaltet. Bei Gleichstrom ging das um 1890 nicht.'],
    ['Patent', 'Recht, eine Erfindung allein zu nutzen oder zu verkaufen.'],
    ['Standard', 'Technische Lösung, auf die sich alle einigen (z. B. Stecker, Spannung). Ein «Standardkrieg» ist der Streit darum.'],
    ['Gilded Age', '«Vergoldetes Zeitalter» (ca. 1870–1900): Zeit des schnellen Wachstums und der Grosskonzerne in den USA.'],
    ['Monopol / Duopol', 'Ein bzw. zwei Unternehmen beherrschen einen Markt.'],
    ['HGÜ', 'Hochspannungs-Gleichstrom-Übertragung: moderne Technik, um Gleichstrom über sehr lange Strecken zu schicken.'],
], [1838, 7414]))
b.append(PB)

b.append(heading('Die wichtigsten Etappen'))
b.append(table([
    ['1882', 'Edison eröffnet in New York das erste Gleichstrom-Kraftwerk (Pearl Street) – Reichweite nur ein Quartier.'],
    ['1886', 'Westinghouse baut das erste Wechselstromnetz mit Transformatoren in den USA.'],
    ['1888', 'Westinghouse kauft Teslas Patente für den Wechselstrommotor. Edisons Lager startet die Angstkampagne.'],
    ['1890', 'Erste Hinrichtung auf dem elektrischen Stuhl – mit Wechselstrom von Westinghouse.'],
    ['1892', 'Bankier J. P. Morgan fusioniert Edisons Firma zu General Electric (GE). Edison verliert die Kontrolle.'],
    ['1893', 'Westinghouse beleuchtet die Weltausstellung in Chicago.'],
    ['1896', 'Strom von den Niagarafällen fliesst über 30 km nach Buffalo. Wechselstrom hat sich durchgesetzt.'],
], [1100, 8152], head=['Jahr', 'Ereignis']))
b.append(para('', after=120))
b.append(heading('Warum gewann der Wechselstrom?'))
b.append(para('Mit dem Transformator liess sich die Spannung erhöhen. So kam Strom mit wenig Verlust über weite Strecken, und '
              'wenige grosse Kraftwerke reichten. Edisons Gleichstrom brauchte dagegen alle paar Kilometer ein Kraftwerk. '
              'Die Angstkampagne half Edison nicht: Die Öffentlichkeit verband Wechselstrom nicht dauerhaft mit dem Tod [3]. '
              'Entschieden haben am Ende die Kosten – und das Kapital von Banken wie J. P. Morgan.', after=200))
b.append(heading('Bedeutung für die Weltmacht USA'))
b.append(para('Während des Stromkriegs überholten die USA Grossbritannien als grösste Industriemacht. Billiger Strom trieb '
              'Fabriken an, und aus dem Streit entstanden Grosskonzerne wie General Electric.', after=60))
b.append(image('rIdChart', 5.6, 2.64, 501, 'Grafik Weltindustrie',
               'Anteil an der Weltindustrieproduktion 1860–1913: USA steigen auf 32 Prozent, Grossbritannien fällt auf 13,6 Prozent.'))
b.append(para('Anteil an der Weltindustrieproduktion in Prozent, 1860–1913. Daten: Bairoch [4].', sz=16, jc='center', after=160))
b.append(items(['Warum war billiger Strom für den Aufstieg der USA so wichtig?',
                'Edison war der berühmtere Erfinder – trotzdem verlor er. Was sagt das über die Rolle von Geld und Firmen aus?']))
b.append(PB)

b.append(heading('Und heute?'))
b.append(items([
    '**Steckdose:** Weltweit kommt Wechselstrom aus der Steckdose – in der Schweiz 230 Volt, in den USA 120 Volt.',
    '**Firmen:** General Electric und Westinghouse gibt es bis heute (z. B. Turbinen, Kernkraftwerke).',
    '**Gleichstrom kommt zurück:** Solarzellen, Batterien, Handys und E-Autos arbeiten mit Gleichstrom. Für sehr lange '
    'Leitungen nutzt man heute HGÜ – das erste kommerzielle System ging 1954 in Schweden in Betrieb, gebaut von ASEA, '
    'heute Teil von ABB [5].',
    '**Neue Standardkriege:** Bei den Ladesteckern für E-Autos in den USA hat sich ab 2023 der Stecker der Firma Tesla '
    'durchgesetzt – benannt nach Nikola Tesla. In der EU ist USB-C für Handys seit 2024 Pflicht.',
], letters=False))
b.append(para('', after=120))
b.append(heading('Diskussion in Gruppen'))
for g, t in [
    ('Gruppe 1:', 'Edison tötete Tiere, um Wechselstrom gefährlich wirken zu lassen. Wo gibt es heute Werbung oder '
                  'Kampagnen, die mit Angst arbeiten?'),
    ('Gruppe 2:', 'Ein Artikel von ABB behauptet, Edison habe sogar einen Elefanten töten lassen. Die Edison Papers der '
                  'Rutgers University widerlegen das [6]. Warum halten sich solche Geschichten so lange?'),
    ('Gruppe 3:', 'Der Stromkrieg war ein Streit um einen Standard. Welche Standardkriege kennt ihr heute (Ladestecker, '
                  'Streaming, Betriebssysteme)? Wer gewinnt meistens – die bessere Technik oder die stärkere Firma?'),
]:
    b.append(para(g, b=True, after=40, keep=True))
    b.append(para(t, after=160))
b.append(para('', after=120))
b.append(heading('Quellen'))
refs = [
    'J. Jonnes, *Empires of Light: Edison, Tesla, Westinghouse, and the Race to Electrify the World*. New York, NY, USA: Random House, 2003.',
    'Thomas A. Edison Papers, “The current wars,” Rutgers Univ. [Online]. Available: https://edison.rutgers.edu/life-of-edison/essaying-edison/essay/the-current-wars',
    'B. M. Cole and D. Chandler, “A model of competitive impression management: Edison versus Westinghouse in the War of the Currents,” *Administrative Science Quarterly*, vol. 64, no. 4, pp. 1020–1063, 2019, doi: 10.1177/0001839218821439.',
    'P. Bairoch, “International industrialization levels from 1750 to 1980,” *Journal of European Economic History*, vol. 11, no. 2, pp. 269–333, 1982.',
    '“Milestones: Gotland High Voltage Direct Current Link, 1954,” *ETHW*, IEEE. [Online]. Available: https://ethw.org/Milestones:Gotland_High_Voltage_Direct_Current_Link,_1954',
    'Thomas A. Edison Papers, “Myth buster: Topsy the elephant,” Rutgers Univ. [Online]. Available: https://edison.rutgers.edu/life-of-edison/essaying-edison/essay/myth-buster-topsy-the-elephant',
]
b.append(references(refs))
b.append(para('Alle Online-Quellen abgerufen am 1.10.2026.', sz=18, after=0))

build(src, out, b, 'Der Stromkrieg – Handout', chart=ROOT / 'assets/chart_weltindustrie.png')
