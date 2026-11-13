"""Playbook für die Präsentation, Stand der Quellen und Vorbereitungsliste (FHNW-Layout).

Aufruf: python3 scripts/playbook.py   → vorbereitung/Playbook_und_Vorbereitung_Stromkrieg.docx
Quellennummern: {schluessel} wird zu [n], mit denselben Nummern wie in Handout und Folien.
"""
from fhnw_docx import PB, ROOT, build, heading, items, para, table
from quellen import gemeinsame_zitate

Z = gemeinsame_zitate()
T = Z.text


def zeilen(rows):
    return [[T(zelle) for zelle in row] for row in rows]


def punkte(lst, letters=False):
    return items([T(x) for x in lst], letters=letters)


b = []
b.append(para('Der Stromkrieg – Playbook für die Präsentation', b=True, sz=36, jc='center', after=60))
b.append(para('10 Minuten · 20 Folien (mit Diskussion: 24) · 10. Dezember 2026', jc='center', after=240))

b.append(heading('Deine Kernbotschaft'))
b.append(para('Den Stromkrieg gewann nicht die bessere Propaganda und nicht das grössere Genie, sondern das billigere '
              'System mit dem stärkeren Kapital – und genau so wurden die USA zur Industriemacht.', after=200))

b.append(heading('Ablauf: Folie für Folie'))
b.append(table([
    ['0:00–0:15', '1 Frontfolie',
     'Steht, während alle Platz nehmen. Erster Klick: Das Video läuft (11 Sekunden, ohne Ton). Dazu: «New York im Jahr '
     '1889. Elektrisches Licht verändert die Stadt. Aber wer liefert den Strom dafür?»',
     'Zweiter Klick: Titelfolie.'],
    ['0:15–0:25', '2 Titel',
     '«Ich spreche über den Stromkrieg: den Streit zwischen den Firmen von Edison und Westinghouse – und über die Rolle '
     'von Nikola Tesla.» Edison, Westinghouse und Tesla erscheinen von selbst.',
     '«Der Streit ging bis zum Äussersten.»'],
    ['0:25–0:50', '3 Auburn 1890',
     '«6. August 1890: Zum ersten Mal stirbt ein Mensch auf dem elektrischen Stuhl, mit einem Generator von '
     'Westinghouse. Beschafft hatte ihn Harold P. Brown gegen den Willen von Westinghouse.»',
     '«Daraus ergibt sich meine Leitfrage.»'],
    ['0:50–1:20', '4 Leitfrage, These',
     'Leitfrage vorlesen. Klick: Die These erscheint. «Entschieden haben Kosten, Kapital und Technik.»',
     '«Dafür braucht es eine Minute Technik.»'],
    ['1:20–2:05', '5 Spannung',
     'Hohe Spannung heisst wenig Verlust. Nur Wechselstrom liess sich damals mit dem Transformator umspannen.',
     '«Wie haben die Zeitgenossen den Streit gesehen?»'],
    ['2:05–2:25', '6 Karikatur',
     '«Die Presse zeichnete den Wettlauf um das Licht als Zweikampf zweier Männer.» Klick: «Tatsächlich stritten Firmen '
     'und ihre Geldgeber.»',
     '«Wer stritt wirklich?»'],
    ['2:25–3:00', '7 Firmen',
     'Drei Konzerne und ein Bankier: Edison General Electric, Westinghouse, Thomson-Houston, J. P. Morgan. Tesla verkauft '
     'Westinghouse 1888 seine Patente.',
     '«Wie lief der Streit ab?»'],
    ['3:00–3:55', '8 Verlauf I',
     'Erst Wettbewerb, ab 1888 Angstkampagne: Brown tötet Hunde, dann der elektrische Stuhl.',
     '«Entschieden hat die Kampagne aber nichts.»'],
    ['3:55–5:00', '9 Verlauf II',
     'Wendepunkt 1892: Edison General Electric fusioniert mit Thomson-Houston zu General Electric. 1893 Chicago, 1896 Niagara.',
     '«Warum also Wechselstrom?»'],
    ['5:00–5:55', '10 Warum',
     'Kosten und Kapital. Die Angstkampagne scheiterte, weil Westinghouse sie mit besserer Technik entkräftete.',
     '«Was heisst das für die USA?»'],
    ['5:55–6:45', '11 Grafik',
     'Zwischen 1880 und 1900 überholen die USA Grossbritannien. 1913: 32 % der Weltindustrie.',
     '«Was steckt hinter dieser Kurve?»'],
    ['6:45–7:25', '12 Zahlen',
     'Elektromotoren in Fabriken: 5 % → 78 %. Der Output pro Arbeitsstunde wächst danach mehr als doppelt so schnell.',
     '«Viele kennen die Geschichte anders.»'],
    ['7:25–8:15', '13 Mythen',
     'Der Elefant, «Tesla gegen Edison», der zerrissene Vertrag: Alle drei halten der Prüfung nicht stand.',
     '«Was bleibt davon bis heute?»'],
    ['8:15–9:05', '14 Und heute?',
     'Wechselstrom im Netz, Gleichstrom kehrt mit HGÜ zurück, Schweizer Spur über BBC und ABB, neue Standardkriege bei '
     'Ladesteckern.',
     '«Zurück zur Leitfrage.»'],
    ['9:05–10:00', '15–16 Fazit, Danke',
     'Beide Teile der Leitfrage beantworten: Der Wechselstrom gewann über Kosten und Kapital. Billiger Strom und '
     'Grosskonzerne trugen den Aufstieg der USA.',
     '«Danke – Fragen?»'],
    ['–', '17–20 Quellen', 'Quellen und Bildnachweis. Nur zeigen, wenn jemand danach fragt.', ''],
], [1050, 1500, 4652, 2050], head=['Zeit', 'Folie', 'Das sagst du', 'Überleitung'], sz=18))
b.append(para('Die ausführlichen Sprechernotizen stehen in der PowerPoint (Ansicht → Notizen).', sz=18, after=200))

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
    'direkt am Bild und auf der letzten Folie. Für Vortrag und Handout im Unterricht ist das mit Quellenangabe üblich. '
    'Stelle Folien und Handout aber nicht öffentlich ins Netz: Die Rechte an den Aufnahmen liegen bei den Agenturen.',
    '**Porträts (optional):** In die Ovale der Titelfolie kannst du Porträts einsetzen; die Anleitung steht in den '
    'Notizen der Folie.',
], letters=False))
b.append(para('', after=100))

b.append(heading('Version mit Diskussion'))
b.append(para('Datei «…_mit_Diskussion»: Nach dem Fazit folgen vier Folien (16–19): das Kapitel «Diskussion» und drei '
              'Fragen ans Plenum (Angst als Werbung, die Elefantin Topsy, Standardkriege heute). Die möglichen Antworten '
              'erscheinen einzeln per Klick: zuerst Stimmen sammeln, dann klicken. Gruppenarbeit gibt es keine mehr. '
              'Danach kommen «Danke» und die Quellen.', after=200))

b.append(heading('Checkliste'))
b.append(punkte([
    'Lerne pro Folie nur den Satz aus «Das sagst du», nicht die ganzen Notizen. Lies die Folien nicht vor.',
    'Übe dreimal laut mit Stoppuhr. Wirst du zu lang, kürze Folie 12 oder lass sie weg.',
    'Klicks: Video (Folie 1), These (Folie 4), zweiter Satz (Folie 6), Antworten in der Diskussion. Alles andere läuft '
    'von selbst. Morph, Animationen und Video funktionieren nur in PowerPoint (2019 oder Microsoft 365).',
    'Diese Zahlen müssen sitzen: 1882, 1888, 1890, 1892, 1893, 1896 · 32 % · 5 % → 78 %.',
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
     'Seitenzahlen: Westinghouse und Great Barrington 1886; Thomson-Houston mit Wechselstrom; Fusion 1892 (Morgan, '
     'Führung bei Thomson-Houston); Niagara–Buffalo und Patentabkommen 1896; Kraftwerk alle paar Kilometer.'],
    ['{kaick}', 'van Kaick und Bachert, Der Radiologe', 'Volltext gelesen',
     'Erledigt: S. 994 belegt den Patentverkauf im Juli 1888, Chicago 1893, die Einheit Tesla (1960) und 0,35–3 T.'],
    ['{hughes58}', 'Hughes, Aufsatz zu Harold P. Brown', 'nicht eingesehen',
     'Seitenzahlen: Tierversuche mit Hilfe aus Edisons Labor; elektrischer Stuhl; Generator über Dritte beschafft.'],
    ['{natgeo}', 'National Geographic History, «Edison and Tesla»', 'Titel und Heft über Suchtreffer',
     'Autorin oder Autor und Datum eintragen (Heft Juli/August 2016). Die Bildlegenden hast du mir geschickt.'],
    ['{essig}', 'Essig, *Edison and the Electric Chair*', 'nicht eingesehen',
     'Seitenzahl: erste Hinrichtung, 6. August 1890, William Kemmler, Auburn.'],
    ['{carlson}', 'Carlson, *Tesla*', 'nicht eingesehen',
     'Seitenzahlen: Verzicht auf Lizenzgebühren 1891; späterer Verkauf der Patente.'],
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
    ['{iec}', 'IEC, World plugs', 'nicht geöffnet', 'Titel der Seite; Schweiz 230 V / 50 Hz, USA 120 V / 60 Hz.'],
    ['{gotland}', 'ETHW, Gotland-Verbindung', 'nicht geöffnet', '1954, erste kommerzielle HGÜ-Verbindung, gebaut von ASEA.'],
    ['{steig}', 'HLS, Charles E. L. Brown', 'nicht geöffnet',
     'Lauffen–Frankfurt 1891 über 175 km; Gründung der BBC 1891 in Baden.'],
    ['{abbHistory}', 'ABB, History', 'nicht geöffnet', 'Zusammenschluss von ASEA und BBC 1988; Titel der Seite.'],
    ['{nacs}', 'Joint Office, SAE J3400', 'Aussage über Suchtreffer',
     'Titel und Datum der Meldung (19.12.2023); Ladestecker von Tesla wird zur Norm SAE J3400.'],
    ['{eu}', 'Richtlinie (EU) 2022/2380', 'nicht geöffnet',
     'Titel, Amtsblatt L 315, S. 30–43, Datum; USB-C-Pflicht für Handys ab Ende 2024.'],
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

b.append(heading('Prüfprotokoll, zweite Runde (streng)'))
b.append(para('**1. Kritische Mängel**', after=40))
b.append(punkte([
    'Seitenzahlen fehlen bei acht Quellen, an 21 Stellen (rot im Handout). Der FHNW-Leitfaden verlangt sie bei Büchern '
    'und Aufsätzen.',
    'Zwanzig von einundzwanzig Quellen sind nicht an der Quelle geprüft (Tabelle oben). Ohne diese Prüfung sind die '
    'Belege formal richtig, aber nicht gesichert.',
    'Bei zwei Online-Quellen fehlen Autorin oder Autor und das genaue Datum ({natgeo}, {srf}).',
    'Die drei Bilder sind Agenturbilder. Nur im Unterricht verwenden und den Bildnachweis stehen lassen.',
], letters=False))
b.append(para('**2. Optimierungspotenzial**', after=40))
b.append(punkte([
    'Folien 4, 7, 13 und 14 haben deutlich mehr als 20 Wörter. Nicht vorlesen; pro Folie einen Satz frei sprechen.',
    '15 Inhaltsfolien in 10 Minuten sind knapp. Folie 12 ist die Reserve zum Kürzen.',
    'Das Video ist selbst erzeugt. Ein echter Filmclip wäre stärker (Vorschlag oben).',
    'Auf der Titelfolie stehen Buchstaben statt Porträts.',
], letters=False))
b.append(para('**3. Gesamtfazit**', after=40))
b.append(para('Aufbau, Leitfrage, These und Quellenkritik tragen; Handout und Folien stimmen in Inhalt und Nummern '
              'überein. Abgabereif ist die Arbeit erst, wenn die Seitenzahlen eingetragen und die Quellen einmal '
              'geöffnet sind.', after=200))
b.append(PB)

b.append(heading('Mögliche Fragen – und kurze Antworten'))
for q, a in [
    ('War Wechselstrom nicht doch gefährlicher?',
     'Hohe Spannung ist gefährlich. Darum wird sie vor dem Haus wieder heruntertransformiert. Edisons Sorge war nicht '
     'erfunden, aber er übertrieb sie, weil es seinem Geschäft half.'),
    ('Was war Teslas Rolle?',
     'Er lieferte mit dem Wechselstrommotor (Patente 1888) das fehlende Stück, damit Wechselstrom auch Fabriken antreiben '
     'konnte. Firma, Netze und Geld kamen von Westinghouse.'),
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

b.append(para('Vorbereitung: Was lesen, was schauen?', b=True, sz=28, after=200))
b.append(heading('1. Zuerst lesen (zusammen etwa 4 Stunden)'))
b.append(table([
    ['van Kaick und Bachert (2016), Der Radiologe', 'Hast du als PDF. Vier Seiten auf Deutsch: Tesla, Westinghouse, '
     'die Einheit Tesla. Ca. 15 Min.'],
    ['Edison Papers: «The Current Wars»', 'Kurzer Überblick der Rutgers University. Hintergrund für Folien 7–9. '
     'Online, ca. 15 Min.'],
    ['Edison Papers: «Myth Buster: Topsy the Elephant»', 'Für die Mythen-Folie. Online, ca. 10 Min.'],
    ['Cole & Chandler (2019)', 'Hast du als PDF. Nur Einleitung und Schluss lesen (ca. 1 Std.): Warum Edisons '
     'Angstkampagne scheiterte. Für Folie 10.'],
    ['Jonnes: *Empires of Light* (2003)', 'Das Standardbuch zum Stromkrieg, gut lesbar. Nicht alles lesen: im '
     'Inhaltsverzeichnis die Kapitel zu 1886–1896 suchen (Great Barrington, elektrischer Stuhl, Chicago, Niagara). '
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
    ['Carlson: *Tesla* (2013)', 'Die Kapitel zu Westinghouse 1888–1896. Belegt Lizenzverzicht und Patentverkauf.'],
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
