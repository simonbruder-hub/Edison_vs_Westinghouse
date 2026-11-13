"""Playbook für die Präsentation, Stand der Quellen und Vorbereitungsliste (FHNW-Layout).

Aufruf: python3 scripts/playbook.py   → vorbereitung/Playbook_und_Vorbereitung_Stromkrieg.docx
Quellennummern: {schluessel} wird zu [n], mit denselben Nummern wie in Handout und Folien.
"""
from fhnw_docx import PB, ROOT, build, heading, items, para, table
from karten import AUSSPRACHE, HAUPTTEIL, KARTEN, TIPPS_AUSSPRACHE
from quellen import gemeinsame_zitate

Z = gemeinsame_zitate()
T = Z.text


def zeilen(rows):
    return [[T(zelle) for zelle in row] for row in rows]


def punkte(lst, letters=False):
    return items([T(x) for x in lst], letters=letters)


b = []
b.append(para('Der Stromkrieg – Playbook für die Präsentation', b=True, sz=36, jc='center', after=60))
b.append(para('10 Minuten · 22 Folien (mit Diskussion: 26) · 10. Dezember 2026', jc='center', after=240))

b.append(heading('Deine Kernbotschaft'))
b.append(para('Edison gegen Tesla: Nicht der berühmtere Erfinder gewann, sondern das billigere System mit dem stärkeren '
              'Geldgeber – Teslas Wechselstrom, finanziert von Westinghouse. Edisons Sicherheitsbedenken waren echt, '
              'entschieden haben Kosten und Kapital. Billiger Strom und Grosskonzerne machten die USA zur Industriemacht.',
              after=200))

b.append(heading('Ablauf: Folie für Folie'))
ablauf = []
for nr, key in enumerate(HAUPTTEIL, 1):
    k = KARTEN[key]
    ablauf.append([k['zeit'], f'{nr} {k["folie"]}', ' · '.join(T(x) for x in k['punkte'][:2]), k['weiter']])
ablauf.append(['–', '18–22 Quellen', 'Quellen und Bildnachweis. Nur zeigen, wenn jemand danach fragt.', ''])
b.append(table(ablauf, [1050, 1700, 4452, 2050], head=['Zeit', 'Folie', 'Kern (ausführlich: Karteikarte)', 'Überleitung'],
               sz=18))
b.append(para('Die Karteikarten stehen als Sprechernotizen in der PowerPoint (Ansicht → Notizen) und druckfertig in '
              '«Karteikarten_Stromkrieg.pdf» (A6, vier pro Seite; an den gestrichelten Linien schneiden).', sz=18,
              after=200))

b.append(heading('Video und Bilder'))
b.append(punkte([
    '**Video auf der Frontfolie:** Es steckt in der PowerPoint-Datei (MP4, 11 Sekunden, ohne Ton) und startet mit dem '
    'ersten Klick. Es beginnt und endet mit dem unveränderten Stich von 1889; dazwischen wird das Bild dunkel, dann '
    'gehen die Lampen nacheinander an. Das Abdunkeln ist eine eigene Bearbeitung und steht so im Bildnachweis.',
    '**Vorher testen:** Spiele die Folie einmal am Präsentationsrechner ab. Startet das Video nicht, klickst du weiter; '
    'das Bild bleibt stehen. Im PDF ist nur das Standbild zu sehen.',
    '**Echten Filmclip einbauen:** Von meiner Umgebung aus konnte ich keine Videos herunterladen. So geht es in '
    'PowerPoint: Einfügen → Video → «Dieses Gerät» (eine MP4-Datei, läuft ohne Internet) oder «Onlinevideo» '
    '(YouTube-Adresse, braucht Internet im Raum). Danach im Register «Wiedergabe»: Start «In Klickreihenfolge», '
    '«Video kürzen» auf 10 bis 20 Sekunden, Lautstärke «Stumm».',
    '**Vorschlag für einen Clip:** Edisons Film «Pan-American Exposition by Night» (1901, Library of Congress). Er zeigt '
    'die Weltausstellung in Buffalo bei Nacht, beleuchtet mit Strom von den Niagarafällen. Passende Stelle: nach '
    'Folie 9. Titel, Jahr und Rechte bitte selbst nachsehen, bevor du ihn zeigst.',
    '**Bilder:** Die drei Bilder stammen aus dem Artikel von National Geographic {natgeo}; die Bildnachweise stehen '
    'direkt am Bild und auf der letzten Folie. Im Handout stehen nur Karikatur und Hinrichtung; das Nachtbild von 1889 '
    'ist auf der Frontfolie. Für Vortrag und Handout im Unterricht ist das mit Quellenangabe üblich. '
    'Stelle Folien und Handout aber nicht öffentlich ins Netz: Die Rechte an den Aufnahmen liegen bei den Agenturen.',
    '**Porträts (optional):** In die Ovale der Titelfolie kannst du Porträts einsetzen; die Anleitung steht in den '
    'Notizen der Folie.',
], letters=False))
b.append(para('', after=100))

b.append(heading('Version mit Diskussion'))
b.append(para('Datei «…_mit_Diskussion»: Nach dem Fazit folgen vier Folien (17–20): das Kapitel «Diskussion» und drei '
              'Fragen ans Plenum (Angst als Werbung, Tesla als Kultfigur, Standardkriege heute). Die möglichen Antworten '
              'erscheinen einzeln per Klick: zuerst Stimmen sammeln, dann klicken. Gruppenarbeit gibt es keine mehr. '
              'Danach kommen «Danke» und die Quellen.', after=200))

b.append(heading('Checkliste'))
b.append(punkte([
    'Lerne pro Folie nur den Satz aus «Das sagst du», nicht die ganzen Notizen. Lies die Folien nicht vor.',
    'Übe dreimal laut mit Stoppuhr. Wirst du zu lang, sprich bei Folie 12 nur über 1913 und lass die Fabrikzahlen weg.',
    'Klicks: Video (Folie 1), These (Folie 4), Gegenargument (Folie 7), zweiter Satz (Folie 13), Antworten in der '
    'Diskussion. Alles andere läuft von selbst. Morph, Animationen und Video funktionieren nur in PowerPoint (2019 oder '
    'Microsoft 365).',
    'Diese Zahlen müssen sitzen: 1882, 1884, 1888, 1890, 1892, 1893, 1896 · 32 % · 5 % → 78 % · 2007.',
    'Aussprache der englischen Namen: Liste unten und auf den letzten drei Karteikarten. Einmal laut durchgehen.',
    'Installiere die Schriften Libre Baskerville und Source Sans 3 (kostenlos bei Google Fonts) und nimm zusätzlich '
    'das PDF der Folien mit. Im PDF sehen die Folien auf jedem Computer richtig aus.',
    'Handout: einmal schwarz-weiss ausdrucken und prüfen, ob Bilder und Grafik lesbar sind. Verteile es erst nach dem '
    'Vortrag, sonst lesen alle statt zuzuhören.',
    'Wenn du den Faden verlierst: zurück zur Leitfrage.',
], letters=False))
b.append(PB)

b.append(heading('Quellen: was geprüft ist und was du prüfen musst'))
b.append(para('Von hier aus konnte ich nur den Artikel aus «Der Radiologe» im Volltext lesen (deine Datei). Alle '
              'Internetseiten waren gesperrt; bei einigen habe ich die Aussage über Suchtreffer bestätigt, die Seite '
              'selbst aber nicht gesehen. Alles Übrige stammt aus Fachwissen und muss an der Quelle geprüft werden. '
              'Im Handout sind fehlende Seitenzahlen rot markiert.', after=100))
b.append(table(zeilen([
    ['{edpCW}', 'Edison Papers, «The current wars»', 'nicht geöffnet',
     'Datierung «späte 1880er- und frühe 1890er-Jahre»; Zugriffsdatum eintragen.'],
    ['{pearl}', 'ETHW, Pearl Street Station', 'Zahlen über Suchtreffer',
     'Rund 85 Kunden, 400 Lampen, Versorgungsgebiet eine Viertelquadratmeile (0,65 km²).'],
    ['{jonnes}', 'Jonnes, *Empires of Light*', 'nicht eingesehen',
     'Seitenzahlen: Morgan als Geldgeber Edisons; Gaulard und Gibbs (US-Rechte bei Westinghouse); Patentprozesse; '
     'Fusion 1892; Niagara–Buffalo und Patentabkommen 1896; Kraftwerk alle paar Kilometer.'],
    ['{kaick}', 'van Kaick und Bachert, Der Radiologe', 'Volltext gelesen',
     'Erledigt: S. 993–994 belegen Herkunft und Studium, 1884 bei Edison und die Trennung nach sechs Monaten, den '
     'Patentverkauf im Juli 1888, Chicago und Niagara 1893, die Einheit Tesla (1960), 0,35–3 T, den Autohersteller '
     '(2003) und Teslas Tod 1943 in Geldnot.'],
    ['{hughes58}', 'Hughes, Aufsatz zu Harold P. Brown', 'nicht eingesehen',
     'Seitenzahlen: Tierversuche mit Hilfe aus Edisons Labor; elektrischer Stuhl; Generator über Dritte beschafft.'],
    ['{edison89}', 'Edison, «The dangers of electric lighting» (1889)', 'Titel und Seiten über Suchtreffer',
     'Wortlaut und Seite des Satzes «There is no plea which will justify the use of high-tension and alternating '
     'currents …» (JSTOR oder Bibliothek).'],
    ['{nyt31}', 'New York Times, 19. Okt. 1931', 'Titel und Datum über Suchtreffer; Wortlaut je nach Abdruck verschieden',
     'Teslas Satz mit der Nadel im Heuhaufen und den 90 Prozent im Original nachlesen; Seite eintragen.'],
    ['{natgeo}', 'National Geographic History, «Edison and Tesla»', 'Titel und Heft über Suchtreffer',
     'Autorin oder Autor und Datum eintragen (Heft Juli/August 2016). Die Bildlegenden hast du mir geschickt.'],
    ['{essig}', 'Essig, *Edison and the Electric Chair*', 'nicht eingesehen',
     'Seitenzahlen: erste Hinrichtung, 6. August 1890, William Kemmler, Auburn; Tod des Arbeiters John Feeks im '
     'Oktober 1889 (Kapitel 17).'],
    ['{carlson}', 'Carlson, *Tesla*', 'nicht eingesehen',
     'Seitenzahlen: Verzicht auf Lizenzgebühren 1891; Galileo Ferraris und das Drehfeld (1885–1888).'],
    ['{cole}', 'Cole und Chandler (hast du als PDF)', 'nur Zusammenfassung bekannt',
     'Seitenzahl: Westinghouse ignorierte, bestritt und entkräftete die Vorwürfe.'],
    ['{bairoch}', 'Bairoch, Industrialisierungsniveaus', 'nicht eingesehen',
     'Tabelle und Seite der Anteile für 1860, 1880, 1900 und 1913 (Hinweis ohne Gewähr: Tab. 10, S. 296).'],
    ['{devine}', 'Devine, «From shafts to wires»', 'Zahlen über Suchtreffer',
     'Seitenzahlen für 5 % → 78 % und 1,3 % / 3,1 %. Frei im Netz: gwern.net/doc/economics/automation/1983-devine.pdf.'],
    ['{abb}', 'ABB, «Tesla gegen Edison»', 'Zitat über Suchtreffer',
     'Wortlaut «grausame Experimente mit Hunden, Pferden und sogar einem Elefanten», Autor und Datum (31.10.2023).'],
    ['{topsy}', 'Edison Papers, «Myth buster: Topsy»', 'nicht geöffnet',
     'Topsy starb 1903; keine Beteiligung Edisons; nur die Filmfirma filmte.'],
    ['{steig}', 'HLS, Charles E. L. Brown', 'nicht geöffnet', 'Lauffen–Frankfurt 1891 über 175 km mit der MFO.'],
    ['{iec}', 'IEC, World plugs', 'nicht geöffnet', 'Titel der Seite; Schweiz 230 V / 50 Hz, USA 120 V / 60 Hz.'],
    ['{tesla19}', 'Tesla, «My inventions», Teil IV (1919)', 'Stelle über Suchtreffer',
     'Titel, Heft und Seite von Teil IV; Satz «the Manager had promised me fifty thousand dollars … practical joke» '
     '(frei im Netz).'],
    ['{gotland}', 'ETHW, Gotland-Verbindung', 'nicht geöffnet', '1954, erste kommerzielle HGÜ-Verbindung, gebaut von ASEA.'],
    ['{bnetza}', 'Bundesnetzagentur, SuedLink (2025)', 'Meldung über Suchtreffer',
     'Rund 700 km, Windstrom von Nord nach Süd, Inbetriebnahme Ende 2028; sonst eine Seite von TenneT oder TransnetBW.'],
    ['{nacs}', 'Joint Office, SAE J3400', 'Aussage über Suchtreffer',
     'Titel und Datum der Meldung (19.12.2023); Ladestecker von Tesla wird zur Norm SAE J3400.'],
    ['{lee}', 'New York Times, City Room, 14.11.2007', 'Ereignis über AP-Meldungen bestätigt',
     'Abschaltung des letzten Gleichstromanschlusses durch Con Edison am 14. November 2007 im Original nachlesen.'],
    ['{met}', 'The Met, Katalog: Currier & Ives, «The Electric Light»', 'Angaben über Suchtreffer',
     'Titel, Jahr 1880 und Bildbeschreibung: Edison in New Jersey, Brush in Manhattan und Brooklyn.'],
    ['{ech}', 'Encyclopedia of Cleveland History, Brush Electric Co.', 'Angaben über Suchtreffer',
     'Kauf durch Thomson-Houston 1889 und General Electric 1892 (andere Artikel dort nennen 1891).'],
    ['{srf}', 'SRF Kultur, Romane über Edison', 'Artikel über Suchtreffer',
     'Autorin oder Autor und Datum eintragen; beide Romane werden dort vorgestellt.'],
    ['{david92}', 'David, «Heroes, herds and hysteresis»', 'nicht eingesehen', 'Seitenzahl zur Pfadabhängigkeit.'],
]), [600, 2450, 1700, 4502], head=['Nr.', 'Quelle', 'Stand', 'Das prüfst du'], sz=18))
b.append(para('Hast du eine Seitenzahl gefunden, ersetzt du im Overleaf-Projekt \\citeoffen{…} durch \\cite[Seite]{…}. '
              'Sind alle eingetragen, aktivierst du in main.tex die Zeile \\entwurffalse.', sz=18, after=200))

b.append(heading('Deine zusätzlichen Links'))
b.append(table(zeilen([
    ['Wikipedia, «Stromkrieg»', 'Gut zum Einlesen. Nach dem FHNW-Leitfaden nicht zitierbar; nimm die dort genannten Quellen.'],
    ['ABB, «Tesla gegen Edison»', 'Fehlerhaft, wie du schreibst. Steht im Handout nur als Beispiel für einen Mythos {abb}.'],
    ['Blog von naturenergie netze; Ørsted «energiewinde»', 'Firmenseiten ohne Belege. Zum Einlesen brauchbar, nicht als '
     'Beleg. Achte auf die Erzählung «Tesla gegen Edison»: ein weiteres Beispiel für Quellenkritik.'],
    ['YouTube-Video und BBC-Sendung', 'Konnte ich nicht öffnen. Als Tipp oder Clip verwendbar. Wenn du daraus zitierst, '
     'brauchst du Titel, Kanal oder Sender, Datum und die Minute.'],
    ['SRF Kultur, Romane über Edison', 'Als Tipp am Ende des Handouts eingebaut {srf}. Romane sind keine Quelle.'],
    ['National Geographic History', 'Quelle der drei Bilder {natgeo}; seriöses Magazin, aber populär. Für Fakten lieber '
     'Jonnes {jonnes} oder die Edison Papers {edpCW}.'],
    ['van Kaick und Bachert, Der Radiologe', 'Fachzeitschrift, kurzer Überblick zu Tesla. Als geprüfte Quelle '
     'eingebaut {kaick}.'],
]), [3000, 6252], head=['Link', 'Einschätzung'], first_center=False, sz=18))
b.append(para('', after=160))

b.append(heading('Prüfprotokoll, dritte Runde (Fachgutachten)'))
b.append(para('**1. Was schon gut ist**', after=40))
b.append(punkte([
    'Klare Leitfrage mit These; der Bezug zum Modul «Weltmacht USA» ist mit Zahlen belegt {bairoch,devine}.',
    'Beide Seiten kommen fair zu Wort: Edisons Sicherheitsargument {edison89} gegen das Kostenargument von Westinghouse '
    'und Tesla; Stärken und Schwächen beider Systeme stehen in einer Tabelle.',
    'Kontext USA und Europa (Gilded Age, europäische Pioniere, 50 und 60 Hz) und Brücke in die Gegenwart (HGÜ, '
    'SuedLink, Ladestecker, Con Edison 2007).',
], letters=False))
b.append(para('**2. Was noch fehlt**', after=40))
b.append(punkte([
    'Seitenzahlen fehlen bei elf Quellen, an 29 Stellen (rot im Handout).',
    'Die Primärquellen {edison89}, {tesla19}, {nyt31} und {lee} kenne ich nur aus Abdrucken und Meldungen. Den Wortlaut '
    'im Original prüfen: JSTOR (North American Review), worldradiohistory.com (Electrical Experimenter), TimesMachine '
    'der New York Times.',
    'Der Artikel aus «Der Radiologe» ist eine kurze Biografie in einer Medizin-Fachzeitschrift: korrekt, aber keine '
    'technikhistorische Studie. Wo möglich, dieselben Angaben bei Carlson {carlson} nachschlagen.',
    'Bei {natgeo} und {srf} fehlen Autorin oder Autor und Datum.',
], letters=False))
b.append(para('**3. Fazit**', after=40))
b.append(para('Inhaltlich und didaktisch vortragsreif. Abgabereif wird die Arbeit, wenn die Seitenzahlen eingetragen und die '
              'Primärquellen am Original geprüft sind.', after=200))
b.append(PB)

b.append(heading('Mögliche Fragen – und kurze Antworten'))
for q, a in [
    ('War Wechselstrom nicht doch gefährlicher?',
     'Hohe Spannung ist gefährlich. Darum wird sie vor dem Haus wieder heruntertransformiert. Edisons Sorge war nicht '
     'erfunden, aber er übertrieb sie, weil es seinem Geschäft half.'),
    ('War der Stromkrieg wirklich ein Duell Edison gegen Tesla?',
     'Es war der Kampf ihrer beiden Systeme: Edisons Gleichstrom gegen Teslas Wechselstrom. Geführt haben ihn aber vor '
     'allem die Firmen: Edison General Electric gegen Westinghouse, das mit Teslas Patenten arbeitete. Tesla lieferte '
     'die Erfindung, Westinghouse Fabriken, Netze und Geld.'),
    ('Hat Edison Tesla um 50’000 Dollar betrogen?',
     'Belegt ist das nicht. In Teslas eigenem Bericht von 1919 machte «der Manager» das Versprechen, nicht Edison. '
     'Sicher ist nur: Die beiden trennten sich nach sechs Monaten, und Edison wollte vom Wechselstrom nichts wissen.'),
    ('War Edison der Bösewicht?',
     'Nein. Edison baute 1882 das erste kommerzielle Stromnetz und hielt Hochspannung ehrlich für gefährlich; 1889 '
     'starben tatsächlich Menschen an Leitungen über New York. Er hielt aber auch an Gleichstrom fest, weil sein '
     'Geschäft daran hing, und sein Lager unterstützte Browns Tierversuche.'),
    ('Warum zeigst du ein Bild von Edison und Brush, wenn es um Edison und Tesla geht?',
     'Das Bild stammt von 1880, Jahre bevor Tesla nach New York kam. Es zeigt zweierlei: Edison hatte von Anfang an '
     'Rivalen, und schon damals stellte man den Wettbewerb als Duell zweier Erfinder dar. Am Ende landeten Brushs und '
     'Edisons Firma 1892 beide in General Electric. Entschieden haben also Firmen und Kapital, nicht ein Duell; dasselbe '
     'gilt für «Edison gegen Tesla».'),
    ('Warum haben die USA 60 Hz und Europa 50 Hz?',
     'Nach gängiger Darstellung stellte Westinghouse um 1890 auf 60 Hz um, auch weil Teslas Motoren damit besser '
     'liefen; in Europa setzte die AEG nach Lauffen–Frankfurt 50 Hz durch. Die Einzelheiten sind in der Literatur '
     'nicht ganz sicher; heute sind beide Frequenzen technisch gleichwertig.'),
    ('Wer war der bessere Erfinder?',
     'Das lässt sich nicht messen. Edison war ein Tüftler und Unternehmer mit grossem Labor, Tesla ein Theoretiker. Für '
     'weite Strecken hatte Tesla die bessere Technik; Edison hatte den Namen und das erste Netz.'),
    ('Warum haben die USA 120 Volt und wir 230 Volt?',
     'Die USA blieben bei der tiefen Spannung aus der Anfangszeit, die auf Edisons rund 110 Volt zurückgeht. In Europa '
     'setzte man später höhere Spannungen durch, weil dafür dünnere Kupferleitungen reichen.'),
    ('Gab es den Stromkrieg auch in Europa?',
     'Weniger heftig. 1891 zeigte die Übertragung von Lauffen nach Frankfurt über 175 km, dass Drehstrom funktioniert, '
     'mit der Maschinenfabrik Oerlikon. Im selben Jahr gründete Charles E. L. Brown in Baden die BBC, die 1988 mit ASEA '
     'zu ABB wurde.'),
    ('War das wirklich der erste Standardkrieg?',
     'Nein. Schon in den 1840er-Jahren stritten britische Bahngesellschaften um die Spurweite. Der Stromkrieg ist aber '
     'der prägende Standardkrieg der Elektrifizierung.'),
    ('Warum zeigt die Grafik nur vier Jahre?',
     'Bairoch schätzt die Anteile nur für Stichjahre. Die Linien verbinden diese Punkte. Sicher ist deshalb nur: Die '
     'USA überholten Grossbritannien zwischen 1880 und 1900.'),
    ('Woher weisst du, dass die Geschichte mit dem Elefanten falsch ist?',
     'Aus den Edison Papers der Rutgers University, einem Archiv mit Edisons Briefen und Akten. Es gibt keinen Beleg für '
     'Edisons Beteiligung; nur seine Filmfirma filmte die Tötung 1903.'),
    ('Ist das Bild auf der Frontfolie echt?',
     'Ja, ein Stich von William Allen Rogers für «Harper’s Weekly» von 1889. Nur das Abdunkeln und Aufleuchten im Video '
     'habe ich selbst gemacht.'),
    ('Wem gehört die HGÜ-Technik von ASEA heute?',
     'ABB hat das Stromnetzgeschäft 2020 an Hitachi verkauft. Es heisst heute Hitachi Energy und sitzt in Zürich.'),
]:
    b.append(para(q, b=True, after=20, keep=True))
    b.append(para(a, after=110))
b.append(para('Weisst du etwas nicht: «Gute Frage, das habe ich nicht untersucht – ich schaue nach.» Das ist besser als raten.',
              after=0))
b.append(PB)

b.append(heading('Aussprache der englischen Namen und Begriffe'))
b.append(para('Grossbuchstaben markieren die betonte Silbe.', sz=18, after=60))
b.append(table([[a, b_] for a, b_ in AUSSPRACHE], [3600, 5652], head=['Name oder Begriff', 'So sprichst du es'],
               first_center=False, sz=18))
b.append(punkte(TIPPS_AUSSPRACHE))
b.append(PB)

b.append(para('Vorbereitung: Was lesen, was schauen?', b=True, sz=28, after=200))
b.append(heading('1. Zuerst lesen (zusammen etwa 4 Stunden)'))
b.append(table([
    ['van Kaick und Bachert (2016), Der Radiologe', 'Hast du als PDF. Vier Seiten auf Deutsch: Tesla, Westinghouse, '
     'die Einheit Tesla. Ca. 15 Min.'],
    ['Edison Papers: «The Current Wars»', 'Kurzer Überblick der Rutgers University. Hintergrund für Folien 5–9. '
     'Online, ca. 15 Min.'],
    ['Tesla: «My Inventions», Teil IV (1919)', 'Teslas eigene Erinnerung an seine Zeit bei Edison, mit der Stelle zu '
     'den 50’000 Dollar. Frei im Netz, englisch, ca. 20 Min.'],
    ['Edison Papers, digitale Edition (edisondigital.rutgers.edu)', 'Originaldokumente auf Englisch, z. B. '
     'Zeitungsausschnitte zum Tod von John Feeks (Oktober 1889) und Briefe von Harold P. Brown. Ca. 30 Min.'],
    ['Edison Papers: «Myth Buster: Topsy the Elephant»', 'Für die Mythen-Folie. Online, ca. 10 Min.'],
    ['Cole & Chandler (2019)', 'Hast du als PDF. Nur Einleitung und Schluss lesen (ca. 1 Std.): Warum Edisons '
     'Angstkampagne scheiterte. Für Folie 10.'],
    ['Jonnes: *Empires of Light* (2003)', 'Das Standardbuch zu Edison, Tesla und Westinghouse, gut lesbar. Nicht alles '
     'lesen: im Inhaltsverzeichnis die Kapitel zu 1884–1896 suchen (Tesla bei Edison, elektrischer Stuhl, Chicago, Niagara). '
     'Ca. 2 Std. Dabei gleich die Seitenzahlen notieren.'],
    ['National Geographic History: «Edison and Tesla»', 'Reich bebildert, englisch. Quelle deiner drei Bilder.'],
    ['Deine Artikel aus Spiegel und Tages-Anzeiger', 'Deutschsprachiger Einstieg. Lies sie mit einer Frage im Kopf: '
     'Erzählen sie ein Duell zweier Genies oder die Geschichte vom Elefanten? Das ist dein Beispiel für Quellenkritik.'],
], [3000, 6252], head=['Was', 'Wofür / wie lange'], first_center=False))
b.append(para('', after=160))
b.append(heading('2. Schauen und hören'))
b.append(table([
    ['ZDF Terra X: «Nikola Tesla – Magier der Elektrizität» (2024)', 'Auf Deutsch, in der ZDF-Mediathek. Guter Einstieg. '
     'Achte darauf, wie stark Tesla als einsames Genie gezeigt wird; das passt zur Mythen-Folie.'],
    ['PBS American Experience: «Edison» (2015) und «Tesla» (2016)', 'Edison als Unternehmer; Teslas Zeit bei '
     'Westinghouse. Englisch. Kann ausserhalb der USA gesperrt sein.'],
    ['BBC World Service (dein Link) und dein YouTube-Video', 'Konnte ich nicht öffnen. Notiere Titel, Datum und die '
     'Stelle, falls du etwas daraus verwendest.'],
    ['Spielfilm «The Current War» (2017), deutsch «Edison – Ein Leben voller Licht»',
     'Erzählt den Stromkrieg selbst. Zeigt gut, wie Filme die Geschichte zuspitzen. Nicht als Quelle zitieren.'],
    ['Deine Filmtipps', '«Prestige» (2006): Tesla als geheimnisvoller Erfinder, gespielt von David Bowie; die Handlung '
     'ist erfunden. «Tajna Nikole Tesle» (1980): jugoslawischer Spielfilm über Tesla. «Edison, the Man» (1940): Edison '
     'als Held, ohne die dunklen Seiten. «Murdoch Mysteries»: Krimiserie, in der Tesla als Figur auftritt. Alle vier '
     'sind Unterhaltung und zeigen, wie die Mythen entstehen.'],
], [3000, 6252], head=['Was', 'Wofür'], first_center=False))
b.append(para('', after=160))
b.append(heading('3. Vertiefen, falls Zeit bleibt'))
b.append(table([
    ['Carlson: *Tesla* (2013)', 'Die Kapitel zu Teslas Zeit bei Edison (1884) und bei Westinghouse 1888–1896. Belegt '
     'Lizenzverzicht und Patentverkauf.'],
    ['Hughes (1958): Aufsatz über Harold P. Brown', 'Tierversuche und die Beschaffung des Generators. 23 Seiten.'],
    ['Hughes: *Networks of Power* (1983)', 'Das wissenschaftliche Standardwerk zur Elektrifizierung. Das Kapitel zum '
     '«Battle of the Systems» reicht.'],
    ['David (1992): «Heroes, Herds and Hysteresis»', 'Wenn du mehr zu Standardkriegen sagen willst. Anspruchsvoll; '
     'die Einleitung reicht.'],
    ['Romane: McCarten, «Licht»; Moore, «Die letzten Tage der Nacht»', 'Zwei Blickwinkel auf denselben Streit; '
     'vorgestellt von SRF Kultur. Zur Unterhaltung, keine Quelle.'],
], [3000, 6252], head=['Was', 'Wofür'], first_center=False))
b.append(para('Bücher findest du über swisscovery (FHNW-Bibliothek), Aufsätze meist über den Online-Zugang der '
              'Bibliothek. Alle vollständigen Angaben stehen im Handout und in der langen Ausarbeitung.', sz=18, after=200))

b.append(heading('Zeitplan bis zum Vortrag'))
b.append(table([
    ['bis 22. November', 'Liste 1 lesen, Liste 2 schauen, Quellen öffnen und Seitenzahlen eintragen (Tabelle oben).'],
    ['bis 3. Dezember', 'Sprechernotizen in eigene Worte umschreiben, Video am Präsentationsrechner testen, offene '
     'Fragen klären.'],
    ['4.–9. Dezember', 'Dreimal laut üben, Handout drucken, Folien als PDF sichern, Technik prüfen.'],
], [2000, 7252], first_center=False))

build(ROOT / 'vorbereitung' / 'Playbook_und_Vorbereitung_Stromkrieg.docx', b, 'Der Stromkrieg – Playbook und Vorbereitung')
