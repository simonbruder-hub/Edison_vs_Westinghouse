"""Handout im FHNW-Layout. Legt die gemeinsamen Quellennummern für Folien und Lösung fest.

Aufruf: python3 scripts/handout_kurz.py   → abgabe/Handout_Stromkrieg_Bruder.docx, assets/zitierreihenfolge.json
"""
from fhnw_docx import PB, ROOT, build, heading, image, items, para, references, table
from quellen import Zitate

z = Zitate()
T = z.text

b = []
b.append(para('Der Stromkrieg – Handout', b=True, sz=36, jc='center', after=60))
b.append(para('Edison vs. Westinghouse und die Rolle von Nikola Tesla · Simon Bruder · Modul «Weltmacht USA» · 10.12.2026',
              jc='center', after=200))
b.append(para(T(
    'Um 1890 stritten in den USA die grossen Elektrofirmen darum, wie das Land mit Strom versorgt wird. Thomas Edisons '
    'Firmen setzten auf **Gleichstrom**, George Westinghouse auf **Wechselstrom**; auch der dritte grosse Anbieter, '
    'Thomson-Houston, bot Wechselstrom an {jonnes}. Ab 1888 nutzte Westinghouse zudem die Motorpatente von Nikola '
    'Tesla {carlson}. Edisons Lager versuchte, Wechselstrom als lebensgefährlich darzustellen: Harold P. Brown tötete '
    'öffentlich Tiere und half, den elektrischen Stuhl einzuführen {hughes58}. Gewonnen hat trotzdem der Wechselstrom. '
    'Der «Stromkrieg» (englisch «War of the Currents») fällt in die späten 1880er- und frühen 1890er-Jahre {edpCW}; '
    'dieses Handout rechnet von 1886 (erste Wechselstromnetze) bis 1892 (Gründung von General Electric).'), after=160))
b.append(table([
    ['**Leitfrage:** Warum setzte sich in den USA der Wechselstrom gegen Edisons Gleichstrom durch – und was zeigt der '
     'Konflikt über den Aufstieg der USA zur führenden Industriemacht?'],
    ['**These:** Entschieden haben nicht Propaganda oder ein einzelnes Genie, sondern Kosten, Kapital und Technik. Aus dem '
     'Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.'],
], [9252], first_center=False))
b.append(para('', after=100))
b.append(items(['Welche Geräte in deinem Alltag laufen mit Gleichstrom, welche mit Wechselstrom?',
                'Kennst du ein aktuelles Beispiel, bei dem sich Firmen um einen technischen Standard streiten?']))
b.append(para('', after=100))
b.append(heading('Tab. 1: Begriffe'))
b.append(table([
    ['Gleichstrom (DC)', 'Strom, der immer in dieselbe Richtung fliesst, z. B. aus einer Batterie.'],
    ['Wechselstrom (AC)', 'Strom, der periodisch die Richtung wechselt; so kommt er aus der Steckdose (Europa 50 Hz, USA 60 Hz).'],
    ['Transformator', 'Setzt die Spannung von Wechselstrom herauf oder herab. Bei gleicher Leistung bedeutet höhere Spannung '
     'kleineren Strom und damit kleinere Leitungsverluste (*P* = *I*² · *R*). Für Gleichstrom gab es um 1890 kein '
     'vergleichbar einfaches Gerät.'],
    ['Drehstrom', 'Wechselstrom mit drei Phasen; Grundlage der heutigen Netze und der Industriemotoren.'],
    ['Thomson-Houston', 'Dritter grosser US-Elektrokonzern neben Edison und Westinghouse; fusionierte 1892 mit Edison '
     'General Electric zu General Electric (GE).'],
    ['Patent, Patentpool', 'Ein Patent ist ein zeitlich befristetes Schutzrecht auf eine Erfindung. In einem Patentpool teilen '
     'Firmen ihre Patente, statt sich zu verklagen (GE und Westinghouse ab 1896).'],
    ['Standardkrieg', 'Streit darum, welche technische Lösung sich als gemeinsamer Standard durchsetzt (z. B. Spannung, Stecker).'],
    ['Pfadabhängigkeit', T('Hat sich ein Standard einmal durchgesetzt, bleibt er lange bestehen, auch wenn später andere '
                           'Lösungen möglich wären {david92}.')],
    ['Gilded Age', '«Vergoldetes Zeitalter» (ca. 1870–1900), ein Begriff von Mark Twain: an der Oberfläche Wachstum und '
     'Grosskonzerne, darunter Korruption und soziale Ungleichheit.'],
    ['HGÜ', 'Hochspannungs-Gleichstrom-Übertragung: Gleichstrom wird mit Stromrichtern über sehr lange Strecken übertragen.'],
], [1838, 7414]))
b.append(PB)

b.append(heading('Tab. 2: Die wichtigsten Etappen'))
b.append(table([
    ['1882', T('Edison eröffnet in New York sein erstes dauerhaftes kommerzielles Zentralkraftwerk (Pearl Street): Gleichstrom '
               'für rund 85 Kunden und 400 Lampen, Reichweite ein Quartier {pearl}.')],
    ['1886', T('Westinghouse gründet seine Elektrofirma; in Great Barrington läuft das erste Wechselstromnetz der USA mit '
               'Transformatoren {jonnes}.')],
    ['1888', T('Westinghouse erwirbt Teslas Patente für den Wechselstrommotor {carlson}. Edisons Lager startet die '
               'Angstkampagne: Harold P. Brown tötet öffentlich Hunde {hughes58}.')],
    ['1890', T('Erste Hinrichtung auf dem elektrischen Stuhl {essig}. Den Westinghouse-Generator dafür hatte Brown gegen den '
               'Willen von Westinghouse über Dritte beschafft {hughes58}.')],
    ['1891', T('Westinghouse gerät in Finanznot; Tesla verzichtet auf seine Lizenzgebühren {carlson}.')],
    ['1892', T('Edison General Electric und der Rivale Thomson-Houston fusionieren zu General Electric, eingefädelt von Banken '
               'um J. P. Morgan. Edisons Name verschwindet, die Führung übernimmt Thomson-Houston {jonnes}.')],
    ['1893', T('Westinghouse beleuchtet die Weltausstellung in Chicago {jonnes}.')],
    ['1896', T('Strom von den Niagarafällen fliesst rund 40 km nach Buffalo; GE und Westinghouse teilen ihre Patente {jonnes}.')],
], [1100, 8152], head=['Jahr', 'Ereignis']))
b.append(para('', after=100))
b.append(heading('Warum gewann der Wechselstrom?'))
b.append(para(T(
    'Mit dem Transformator liess sich die Spannung erhöhen. So kam Strom mit wenig Verlust über weite Strecken, und wenige '
    'grosse Kraftwerke reichten; Edisons Gleichstrom brauchte dagegen alle paar Kilometer ein Kraftwerk {jonnes}. Die '
    'Angstkampagne half Edison nicht: Westinghouse ignorierte sie zuerst, bestritt dann die Vorwürfe und machte sie '
    'schliesslich mit besserer Technik gegenstandslos {cole}. Entschieden haben am Ende die Kosten und das Kapital: Seit 1892 '
    'setzte mit General Electric auch Edisons frühere Firma auf Wechselstrom {jonnes}. In Europa lief es ähnlich: 1891 '
    'übertrugen die Maschinenfabrik Oerlikon und die AEG Drehstrom über 175 km von Lauffen nach Frankfurt {steig}.'),
    after=160))
b.append(heading('Bedeutung für die Weltmacht USA'))
b.append(para(T(
    'Zwischen 1880 und 1900, in den Jahrzehnten des Stromkriegs, überholten die USA Grossbritannien als grösste '
    'Industriemacht (Abb. 1) {bairoch}. Billiger Strom trieb die Fabriken an: Elektromotoren lieferten 1899 knapp 5 %, 1929 '
    'rund 78 % der Antriebsleistung der US-Industrie. Der Output pro Arbeitsstunde wuchs vor 1919 um 1,3 %, danach um 3,1 % '
    'pro Jahr {devine}. Aus dem Streit gingen Grosskonzerne hervor: General Electric und Westinghouse teilten ab 1896 ihre '
    'Patente und beherrschten den Markt {jonnes}.'), after=40))
b.append(image('rIdChart', 6.0, 2.37, 501, 'Abbildung 1',
               'Liniendiagramm 1860 bis 1913: Der Anteil der USA an der Weltindustrieproduktion steigt von 7,2 auf 32,0 Prozent, '
               'Grossbritannien fällt von 22,9 Prozent (1880) auf 13,6 Prozent, Deutschland steigt auf 14,8 Prozent.'))
b.append(para(T('Abb. 1: Anteil an der Weltindustrieproduktion in Prozent. Die Linien verbinden die Stichjahre 1860, 1880, '
                '1900 und 1913; es sind Schätzwerte. Eigene Darstellung nach {bairoch}.'), sz=16, jc='center', after=120))
b.append(items(['Warum war billiger Strom für den Aufstieg der USA so wichtig?',
                'Edison war der berühmtere Erfinder – trotzdem verlor er. Was sagt das über die Rolle von Geld und Firmen aus?']))
b.append(PB)

b.append(heading('Tab. 3: Drei Mythen'))
b.append(table([
    [T('Edison führte öffentlich «Experimente mit Hunden, Pferden und sogar einem Elefanten» vor {abb}.'),
     T('Die Elefantin Topsy starb 1903, über zehn Jahre nach dem Stromkrieg, und ohne Edisons Beteiligung {topsy}.')],
    ['Der Stromkrieg war ein Duell «Tesla gegen Edison».',
     T('Es stritten Firmen und ihre Geldgeber. Tesla lieferte Westinghouse 1888 die Motorpatente {carlson}.')],
    ['Tesla zerriss aus Grossmut seinen Vertrag mit Westinghouse.',
     T('Belegt sind der Verzicht auf Lizenzgebühren (1891) und der spätere Verkauf der Patente {carlson}.')],
], [4626, 4626], head=['Mythos', 'Befund'], first_center=False))
b.append(para('', after=100))
b.append(heading('Und heute?'))
b.append(items([
    T('**Wechselstrom im Netz:** Weltweit kommt Wechselstrom aus der Steckdose, in der Schweiz mit 230 Volt, in den USA mit '
      '120 Volt {iec}.'),
    T('**Gleichstrom kehrt zurück:** Akkus, Solarzellen und Elektronik arbeiten mit Gleichstrom. Für sehr lange Leitungen '
      'nutzt man HGÜ; die erste kommerzielle Verbindung moderner Bauart ging 1954 in Schweden in Betrieb, gebaut von ASEA '
      '{gotland}.'),
    T('**Schweizer Spur:** Charles E. L. Brown von der Maschinenfabrik Oerlikon gründete 1891 in Baden die BBC {steig}. BBC '
      'und ASEA schlossen sich 1988 zu ABB zusammen {abbHistory}.'),
    T('**Neue Standardkriege:** In den USA stellten ab 2023 fast alle Autohersteller auf den Ladestecker der Firma Tesla um '
      '{nacs}. In der EU ist USB-C für Handys seit Ende 2024 Pflicht {eu}.'),
], letters=False))
b.append(para('', after=100))
b.append(heading('Kernaussagen'))
b.append(items([
    'Der Wechselstrom gewann, weil er billiger war: hohe Spannung, wenige grosse Kraftwerke, weite Strecken.',
    'Entschieden hat das Kapital, nicht die Angstkampagne: 1892 ging Edisons Firma in General Electric auf.',
    'Billiger Strom und Grosskonzerne trugen den Aufstieg der USA zur grössten Industriemacht.',
], letters=False))
b.append(para('', after=100))
b.append(heading('Diskussion in Gruppen'))
for g, t in [
    ('Gruppe 1:', 'Harold P. Brown tötete öffentlich Tiere, um Wechselstrom gefährlich wirken zu lassen. Wo gibt es heute '
                  'Werbung oder Kampagnen, die mit Angst arbeiten?'),
    ('Gruppe 2:', 'Der ABB-Artikel erzählt die Geschichte vom Elefanten, die Edison Papers der Rutgers University widerlegen '
                  'sie (Tab. 3). Warum halten sich solche Geschichten so lange?'),
    ('Gruppe 3:', 'Der Stromkrieg war ein Streit um einen Standard. Welche Standardkriege kennt ihr heute (Ladestecker, '
                  'Streaming, Betriebssysteme)? Wer gewinnt meistens: die bessere Technik oder die stärkere Firma?'),
]:
    b.append(para(g, b=True, after=30, keep=True))
    b.append(para(t, after=110))
b.append(PB)

b.append(heading('Quellen'))
b.append(references(z.eintraege()))

z.speichern()
build(ROOT / 'abgabe' / 'Handout_Stromkrieg_Bruder.docx', b, 'Der Stromkrieg – Handout',
      chart=ROOT / 'assets' / 'chart_weltindustrie.png')
print(len(z.order), 'Quellen:', ', '.join(z.order))
