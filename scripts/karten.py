"""Karteikarten für den freien Vortrag: wenige Stichpunkte pro Folie.

Eine Quelle steht als {schluessel} und wird zur Nummer wie im Handout.
Verwendet von praesentation_vorlage.py (Sprechernotizen) und karteikarten.py (Druckversion, A6).
"""

# Jede Karte: Folie, Zeitfenster, Stichpunkte, Überleitung (weiter). «Klick:» markiert, wann du klickst.
KARTEN = {
    'front': dict(folie='Frontfolie', zeit='0:00–0:20', punkte=[
        'Warten, bis es ruhig ist. Klick: Video läuft (11 s, ohne Ton)',
        'Frage: «Wer kennt Nikola Tesla? Hand hoch.» – «Und wer kennt George Westinghouse?»',
        '«New York 1889: Elektrisches Licht erobert die Stadt. Aber mit welchem Strom?»',
    ], weiter='Klick: Titelfolie'),
    'titel': dict(folie='Titel', zeit='0:20–0:30', punkte=[
        '«Mein Thema: der Stromkrieg, Edison gegen Tesla.»',
        'Animation abwarten: E links, T rechts, dann W',
        '«W steht für Westinghouse, den Mann mit dem Geld.»',
    ], weiter='«Der Streit ging bis zum Äussersten.»'),
    'auburn': dict(folie='Auburn 1890', zeit='0:30–0:55', punkte=[
        '6. August 1890, Auburn, Staat New York',
        'William Kemmler: erste Hinrichtung auf dem elektrischen Stuhl {essig}',
        'Generator von Westinghouse; Harold P. Brown hatte ihn gegen dessen Willen beschafft {hughes58}',
        'Botschaft dahinter: «Wechselstrom tötet»',
    ], weiter='«Daraus ergibt sich meine Leitfrage.»'),
    'leitfrage': dict(folie='Leitfrage und These', zeit='0:55–1:20', punkte=[
        'Leitfrage vorlesen (der einzige Satz zum Ablesen)',
        'Klick: These',
        'In eigenen Worten: Edison war berühmter, Tesla hatte die Technik für weite Strecken; gewonnen hat das System '
        'mit dem Geld dahinter',
    ], weiter='«Wer waren die beiden?»'),
    'erfinder': dict(folie='Zwei Erfinder und ihre Geldgeber', zeit='1:20–2:00', punkte=[
        'Edison: Erfinder und Unternehmer, Glühlampe, Gleichstromnetz {jonnes}',
        'Tesla: aus Smiljan (heute Kroatien), 1884 zu Edison; nach sechs Monaten getrennt: andere Charaktere, Streit '
        'um Wechselstrom {kaick}',
        'Tesla 1931 über Edison: zu viel Probieren, zu wenig Theorie {nyt31}',
        'Geld: Morgan hinter Edison; Westinghouse kauft 1888 Teslas Patente {jonnes,kaick}',
    ], weiter='«Was konnten die beiden Systeme?»'),
    'systeme': dict(folie='Zwei Systeme', zeit='2:00–2:45', punkte=[
        'Gleichstrom: niedrige Spannung, Akkus, erprobte Motoren; aber nur ein Quartier weit {pearl}',
        'Wechselstrom: Transformator, hohe Spannung, weite Strecken {jonnes}',
        'Schwächen: Hochspannung ist gefährlich; Teslas Motor gibt es erst ab 1888 {essig,kaick}',
        'Merksatz: 10-fache Spannung, 1/100 des Verlusts',
    ], weiter='«Beide Seiten hatten gute Argumente.»'),
    'argumente': dict(folie='Sicherheit gegen Kosten', zeit='2:45–3:20', punkte=[
        'Edison 1889: nichts rechtfertige Hochspannung und Wechselstrom {edison89}',
        'Anlass: Oktober 1889 stirbt der Arbeiter John Feeks in den Leitungen über Manhattan {essig}',
        'Klick: Westinghouse und Tesla: billiger Strom über weite Strecken, auch aus Wasserkraft {kaick}',
        'Fair bleiben: Edisons Sorge war echt, und sie passte zu seinem Geschäft',
    ], weiter='«Wie lief der Streit ab?»'),
    'verlauf1': dict(folie='1882–1890', zeit='3:20–4:10', punkte=[
        '1882 Pearl Street: 85 Kunden, 400 Lampen {pearl}',
        '1884 Tesla bei Edison, nach sechs Monaten getrennt {kaick}',
        '1888 Westinghouse kauft Teslas Patente {kaick}; Brown tötet öffentlich Hunde {hughes58}',
        '1890 elektrischer Stuhl {essig}',
    ], weiter='«Entschieden hat die Kampagne aber nichts.»'),
    'verlauf2': dict(folie='1891–1896', zeit='4:10–5:00', punkte=[
        '1891 Westinghouse in Geldnot; Tesla verzichtet auf Lizenzgebühren {carlson}',
        '1892 Edisons Firma und Thomson-Houston werden General Electric, mit Morgan {jonnes}',
        '1893 Weltausstellung Chicago mit Teslas System {kaick}',
        '1896 Niagara–Buffalo, rund 40 km {jonnes}',
    ], weiter='«Warum also Teslas System?»'),
    'warum': dict(folie='Warum Teslas System gewann', zeit='5:00–5:40', punkte=[
        'Kosten: wenige grosse Kraftwerke statt vieler kleiner {jonnes}',
        'Kapital: Westinghouse; 1892 setzt auch GE auf Wechselstrom {kaick,jonnes}',
        'Kampagne verpufft: ignorieren, bestreiten, mit Technik widerlegen {cole}',
    ], weiter='«Typisch amerikanisch? Ein Blick nach Europa.»'),
    'europa': dict(folie='USA und Europa', zeit='5:40–6:20', punkte=[
        'USA: Gilded Age, kaum Regeln, Bankiers wie Morgan, Patentprozesse {jonnes}',
        'Europa: Transformator von Gaulard und Gibbs; Drehfeld zeitgleich bei Ferraris in Turin {jonnes,carlson}',
        'Schweiz: 1891 MFO Oerlikon und AEG, Drehstrom von Lauffen nach Frankfurt, 175 km {steig}',
        'Bis heute: USA 120 V / 60 Hz, Schweiz 230 V / 50 Hz {iec}',
    ], weiter='«Und was bedeutete das für die USA?»'),
    'grafik': dict(folie='Die USA überholen Grossbritannien', zeit='6:20–7:00', punkte=[
        '1880: Grossbritannien 22,9 %, USA 14,7 % {bairoch}',
        '1913: USA 32 %, mehr als Grossbritannien und Deutschland zusammen',
        'Fabriken: Elektromotoren 5 % (1899), 78 % (1929) {devine}',
        'Vorsicht: Zusammenhang, kein Beweis; auch Stahl, Bahn, Binnenmarkt',
    ], weiter='«Edisons erster Rivale war übrigens nicht Tesla.»'),
    'karikatur': dict(folie='Vor Tesla: Edison gegen Brush', zeit='7:00–7:25', punkte=[
        'Bild von 1880: Edison in New Jersey, Charles Brush (Bogenlampe) in New York {met}',
        'Rivalen um das Licht in New York, Jahre vor Tesla {natgeo}',
        'Klick: 1889 kauft Thomson-Houston Brushs Firma, 1892 sind beide in GE {ech}',
        'Klick: Erzählt wird ein Duell, entschieden haben Firmen; heute dasselbe mit Tesla',
    ], weiter='«So entstehen Mythen.»'),
    'mythen': dict(folie='Drei Mythen', zeit='7:25–8:10', punkte=[
        'Elefant: Tierversuche machte Brown; Topsy starb 1903, ohne Edison {hughes58,topsy}',
        '50’000 Dollar: Tesla selbst (1919) nennt «den Manager» {tesla19}',
        'Tesla allein: Transformator aus Europa, Ferraris zeitgleich {jonnes,carlson}',
    ], weiter='«Was bleibt bis heute?»'),
    'heute': dict(folie='Und heute?', zeit='8:10–9:00', punkte=[
        'Wechselstrom aus jeder Steckdose {iec}',
        'Gleichstrom kehrt zurück: Akkus, E-Autos, HGÜ; SuedLink ab 2028 {gotland,bnetza}',
        'Tesla: Einheit, Autohersteller, Ladestecker-Norm {kaick,nacs}',
        'Pointe: Edisons Gleichstrom floss in Manhattan bis 2007 (Con Edison) {lee}',
    ], weiter='«Zurück zur Leitfrage.»'),
    'fazit': dict(folie='Fazit', zeit='9:00–10:00', punkte=[
        'Teslas System gewann: billiger über weite Strecken',
        'Edisons Sorge war echt; entschieden haben Kosten und Kapital',
        'USA: billiger Strom und Grosskonzerne, Aufstieg zur Industriemacht',
        'Schlusssatz: «Edison hatte den Namen, Tesla die Technik. Gewonnen hat das System mit dem Geld dahinter.»',
    ], weiter='«Danke – Fragen?»'),
    'danke': dict(folie='Danke', zeit='ab 10:00', punkte=[
        'Fragen sammeln; Quellen und Bildnachweis nur zeigen, wenn jemand fragt',
        'Weisst du etwas nicht: «Gute Frage, das habe ich nicht untersucht.»',
    ], weiter=''),
    'diskussion': dict(folie='Diskussion', zeit='nach Absprache', punkte=[
        'Nur, wenn Zeit bleibt; Dauer vorher mit dem Dozenten klären',
        'Pro Frage zwei, drei Stimmen sammeln, dann die Antworten per Klick zeigen',
    ], weiter=''),
    'frage1': dict(folie='Frage 1: Angst als Werbung', zeit='Diskussion', punkte=[
        'Beispiele sammeln: Werbung gegen die Konkurrenz, Wahlkampf',
        'Angst wirkt kurz; Nutzen und Kosten entscheiden {cole}',
    ], weiter=''),
    'frage2': dict(folie='Frage 2: Tesla als Kultfigur', zeit='Diskussion', punkte=[
        'Armes Genie gegen reichen Unternehmer: die Geschichte bleibt hängen',
        'Filme, Internet, Autohersteller machen den Namen bekannt {kaick}',
        'Quellen sind nüchterner, z. B. bei den 50’000 Dollar {tesla19}',
    ], weiter=''),
    'frage3': dict(folie='Frage 3: Standardkriege heute', zeit='Diskussion', punkte=[
        'VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker {nacs}',
        'Oft gewinnt, wer Partner, tiefe Kosten und Kapital hat; Standards bleiben lange {david92}',
    ], weiter=''),
}

# Reihenfolge im Vortrag (Hauptteil, danach die Diskussion)
HAUPTTEIL = ['front', 'titel', 'auburn', 'leitfrage', 'erfinder', 'systeme', 'argumente', 'verlauf1', 'verlauf2',
             'warum', 'europa', 'grafik', 'karikatur', 'mythen', 'heute', 'fazit', 'danke']
DISKUSSION = ['diskussion', 'frage1', 'frage2', 'frage3']

# Aussprache der englischen Namen und Begriffe (Grossbuchstaben = betonte Silbe)
AUSSPRACHE = [
    ('Thomas Alva Edison', 'TOM-mes ÄL-we ÄD-di-ssn'),
    ('Nikola Tesla', 'NI-ko-la TES-la (serbisch, Betonung vorne)'),
    ('George Westinghouse', 'DSCHORDSCH WES-ting-haus'),
    ('J. P. Morgan', 'DSCHEI-PII MOR-gn'),
    ('Harold P. Brown', 'HÄ-rold PII BRAUN'),
    ('William Kemmler', 'WIL-jem KEM-ler'),
    ('Auburn', 'OO-börn'),
    ('Pearl Street', 'PÖRL STRIIT'),
    ('Manhattan', 'män-HÄT-tn'),
    ('Thomson-Houston', 'TOM-ssn HJUU-sstn'),
    ('General Electric', 'DSCHE-ne-rel i-LEK-trik'),
    ('Westinghouse Electric', 'WES-ting-haus i-LEK-trik'),
    ('Chicago', 'schi-KAA-go'),
    ('Niagara Falls', 'nai-ÄG-re FOOLS'),
    ('Buffalo', 'BAF-fe-lo'),
    ('Currier & Ives', 'KÖ-ri-er änd AIWS'),
    ('Charles Brush', 'TSCHAARLS BRASCH'),
    ('Harper’s Weekly', 'HAAR-pers WII-kli'),
    ('Gilded Age', 'GIL-did EIDSCH'),
    ('Robber Barons', 'ROB-ber BÄ-rens'),
    ('War of the Currents', 'WOOR ow dhe KA-rents'),
    ('Alternating Current', 'OL-ter-nei-ting KA-rent'),
    ('Direct Current', 'dai-REKT KA-rent'),
    ('Con Edison', 'KON ÄD-di-ssn'),
    ('Topsy', 'TOP-si'),
    ('Gaulard (franz.) und Gibbs', 'go-LAAR und GIBS'),
    ('Galileo Ferraris (ital.)', 'ga-li-LÄ-o fer-RAA-ris'),
    ('Smiljan (kroat.)', 'SMI-ljan'),
    ('Schenectady (falls erwähnt)', 'ske-NEK-te-di'),
]
TIPPS_AUSSPRACHE = [
    'th wie ein gelispeltes s (Zunge an die Zähne): «the», «Thomson»',
    'w mit runden Lippen wie ein kurzes u: «Westinghouse», «War»',
    'Namen in einem deutschen Vortrag ruhig deutsch betonen; wichtig ist, dass sie sicher klingen',
]


def notiz(key, text):
    """Sprechernotiz für PowerPoint; text wandelt {schluessel} in Quellennummern um."""
    k = KARTEN[key]
    nr = (HAUPTTEIL + DISKUSSION).index(key) + 1
    zeilen = [f'KARTE {nr} · {k["zeit"]} · {k["folie"]}']
    zeilen += [f'• {p}' for p in k['punkte']]
    if k['weiter']:
        zeilen.append(f'→ {k["weiter"]}')
    return text('\n'.join(zeilen))
