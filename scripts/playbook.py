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
b.append(para('10 Minuten · 21 Folien (mit Diskussion: 25) · 10. Dezember 2026', jc='center', after=240))

b.append(heading('Deine Kernbotschaft'))
b.append(para('Edison gegen Tesla: Nicht der berühmtere Erfinder gewann, sondern das billigere System mit dem stärkeren '
              'Geldgeber – Teslas Wechselstrom, finanziert von Westinghouse. Billiger Strom und Grosskonzerne machten die '
              'USA zur Industriemacht.', after=200))

b.append(heading('Ablauf: Folie für Folie'))
b.append(table([
    ['0:00–0:15', '1 Frontfolie',
     'Steht, während alle Platz nehmen. Erster Klick: Das Video läuft (11 Sekunden, ohne Ton). Dazu: «New York im Jahr '
     '1889. Elektrisches Licht verändert die Stadt. Aber welcher Strom soll fliessen: Edisons Gleichstrom oder Teslas '
     'Wechselstrom?»',
     'Zweiter Klick: Titelfolie.'],
    ['0:15–0:25', '2 Titel',
     '«Ich spreche über den Stromkrieg: den Streit zwischen Thomas Edison und Nikola Tesla – und über die Rolle von '
     'George Westinghouse.» Edison und Tesla fliegen ein, dann erscheint das W für Westinghouse.',
     '«Der Streit ging bis zum Äussersten.»'],
    ['0:25–0:50', '3 Auburn 1890',
     '«6. August 1890: Zum ersten Mal stirbt ein Mensch auf dem elektrischen Stuhl, mit einem Generator von '
     'Westinghouse. Beschafft hatte ihn Harold P. Brown gegen den Willen von Westinghouse.»',
     '«Daraus ergibt sich meine Leitfrage.»'],
    ['0:50–1:15', '4 Leitfrage, These',
     'Leitfrage vorlesen. Klick: Die These erscheint. «Nicht der berühmtere Erfinder gewann, sondern das billigere System '
     'mit dem stärkeren Geldgeber.»',
     '«Wer waren die beiden?»'],
    ['1:15–2:00', '5 Zwei Erfinder',
     'Edison, Erfinder und Unternehmer. Tesla kam 1884 zu Edison; nach sechs Monaten trennten sie sich. Hinter beiden '
     'stand Geld: Morgan bei Edison, Westinghouse bei Tesla.',
     '«Wie dachten die beiden übereinander?»'],
    ['2:00–2:40', '6 Zwei Blicke',
     'Edison 1889: Nichts rechtfertige Hochspannung und Wechselstrom. Klick: Tesla 1931: Edison suche die Nadel im '
     'Heuhaufen Halm für Halm.',
     '«Wer hatte technisch recht?»'],
    ['2:40–3:20', '7 Spannung',
     'Hohe Spannung heisst wenig Verlust. Nur Wechselstrom liess sich damals umspannen, und Teslas Motor trieb Fabriken an.',
     '«Wie lief der Streit ab?»'],
    ['3:20–4:15', '8 Verlauf I',
     '1882 Pearl Street, 1884 Tesla bei Edison, 1888 Teslas Patente an Westinghouse, 1890 der elektrische Stuhl.',
     '«Entschieden hat die Kampagne aber nichts.»'],
    ['4:15–5:10', '9 Verlauf II',
     '1891 Lizenzverzicht, 1892 geht Edisons Firma in General Electric auf, 1893 Chicago, 1896 Niagara: Teslas System '
     'setzt sich durch.',
     '«Warum also Teslas System?»'],
    ['5:10–5:55', '10 Warum',
     'Kosten und Kapital. Die Angstkampagne scheiterte, weil Westinghouse sie mit besserer Technik entkräftete.',
     '«Was heisst das für die USA?»'],
    ['5:55–6:35', '11 Grafik',
     'Zwischen 1880 und 1900 überholen die USA Grossbritannien. 1913: 32 % der Weltindustrie.',
     '«Was steckt hinter dieser Kurve?»'],
    ['6:35–7:05', '12 Zahlen',
     'Elektromotoren in Fabriken: 5 % → 78 %. Reserve: Diese Folie kürzt du, wenn die Zeit knapp wird.',
     '«Wie haben die Zeitgenossen den Streit gesehen?»'],
    ['7:05–7:25', '13 Karikatur',
     'Die Presse liebte Duelle: Edison gegen Brush. Klick: Genauso erzählt man heute Edison gegen Tesla.',
     '«Aus solchen Zuspitzungen entstehen Mythen.»'],
    ['7:25–8:15', '14 Mythen',
     'Der Elefant, die 50’000 Dollar, der zerrissene Vertrag: Die Quellen erzählen es anders.',
     '«Was bleibt von Edison und Tesla bis heute?»'],
    ['8:15–9:05', '15 Und heute?',
     'Teslas Wechselstrom im Netz, Edisons Gleichstrom kommt zurück, Tesla als Name: Einheit, Autohersteller, Ladestecker.',
     '«Zurück zur Leitfrage.»'],
    ['9:05–10:00', '16–17 Fazit, Danke',
     '«Edison hatte den Namen, Tesla die Technik. Gewonnen hat das System mit dem Geld dahinter.» Billiger Strom und '
     'Grosskonzerne trugen den Aufstieg der USA.',
     '«Danke – Fragen?»'],
    ['–', '18–21 Quellen', 'Quellen und Bildnachweis. Nur zeigen, wenn jemand danach fragt.', ''],
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
b.append(para('Datei «…_mit_Diskussion»: Nach dem Fazit folgen vier Folien (17–20): das Kapitel «Diskussion» und drei '
              'Fragen ans Plenum (Angst als Werbung, Tesla als Kultfigur, Standardkriege heute). Die möglichen Antworten '
              'erscheinen einzeln per Klick: zuerst Stimmen sammeln, dann klicken. Gruppenarbeit gibt es keine mehr. '
              'Danach kommen «Danke» und die Quellen.', after=200))

b.append(heading('Checkliste'))
b.append(punkte([
    'Lerne pro Folie nur den Satz aus «Das sagst du», nicht die ganzen Notizen. Lies die Folien nicht vor.',
    'Übe dreimal laut mit Stoppuhr. Wirst du zu lang, kürze Folie 12 oder lass sie weg.',
    'Klicks: Video (Folie 1), These (Folie 4), Teslas Zitat (Folie 6), zweiter Satz (Folie 13), Antworten in der '
    'Diskussion. Alles andere läuft '
    'von selbst. Morph, Animationen und Video funktionieren nur in PowerPoint (2019 oder Microsoft 365).',
    'Diese Zahlen müssen sitzen: 1882, 1884, 1888, 1890, 1892, 1893, 1896 · 32 % · 5 % → 78 %.',
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
     'Seitenzahlen: Morgan als Geldgeber Edisons; Thomson-Houston mit Wechselstrom; Fusion 1892 (Führung bei '
     'Thomson-Houston); Niagara–Buffalo und Patentabkommen 1896; Kraftwerk alle paar Kilometer.'],
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
    ['{tesla19}', 'Tesla, «My inventions», Teil IV (1919)', 'Stelle über Suchtreffer',
     'Titel, Heft und Seite von Teil IV; Satz «the Manager had promised me fifty thousand dollars … practical joke» '
     '(frei im Netz).'],
    ['{gotland}', 'ETHW, Gotland-Verbindung', 'nicht geöffnet', '1954, erste kommerzielle HGÜ-Verbindung, gebaut von ASEA.'],
    ['{nacs}', 'Joint Office, SAE J3400', 'Aussage über Suchtreffer',
     'Titel und Datum der Meldung (19.12.2023); Ladestecker von Tesla wird zur Norm SAE J3400.'],
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
    'Seitenzahlen fehlen bei elf Quellen, an 24 Stellen (rot im Handout). Der FHNW-Leitfaden verlangt sie bei Büchern, '
    'Aufsätzen und Zeitungsartikeln.',
    'Zwanzig von einundzwanzig Quellen sind nicht an der Quelle geprüft (Tabelle oben). Ohne diese Prüfung sind die '
    'Belege formal richtig, aber nicht gesichert.',
    'Bei zwei Online-Quellen fehlen Autorin oder Autor und das genaue Datum ({natgeo}, {srf}).',
    'Die drei Bilder sind Agenturbilder. Nur im Unterricht verwenden und den Bildnachweis stehen lassen.',
    'Die drei Originalzitate (Edison 1889, Tesla 1919 und 1931) habe ich übersetzt; den Wortlaut kenne ich nur aus '
    'Abdrucken. Vor dem Vortrag im Original prüfen.',
], letters=False))
b.append(para('**2. Optimierungspotenzial**', after=40))
b.append(punkte([
    'Folien 4, 5, 14 und 15 haben deutlich mehr als 20 Wörter. Nicht vorlesen; pro Folie einen Satz frei sprechen.',
    '16 Inhaltsfolien in 10 Minuten sind knapp. Folie 12 ist die Reserve zum Kürzen.',
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
    ('War der Stromkrieg wirklich ein Duell Edison gegen Tesla?',
     'Es war der Kampf ihrer beiden Systeme: Edisons Gleichstrom gegen Teslas Wechselstrom. Geführt haben ihn aber vor '
     'allem die Firmen: Edison General Electric gegen Westinghouse, das mit Teslas Patenten arbeitete. Tesla lieferte '
     'die Erfindung, Westinghouse Fabriken, Netze und Geld.'),
    ('Hat Edison Tesla um 50’000 Dollar betrogen?',
     'Belegt ist das nicht. In Teslas eigenem Bericht von 1919 machte «der Manager» das Versprechen, nicht Edison. '
     'Sicher ist nur: Die beiden trennten sich nach sechs Monaten, und Edison wollte vom Wechselstrom nichts wissen.'),
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

b.append(para('Vorbereitung: Was lesen, was schauen?', b=True, sz=28, after=200))
b.append(heading('1. Zuerst lesen (zusammen etwa 4 Stunden)'))
b.append(table([
    ['van Kaick und Bachert (2016), Der Radiologe', 'Hast du als PDF. Vier Seiten auf Deutsch: Tesla, Westinghouse, '
     'die Einheit Tesla. Ca. 15 Min.'],
    ['Edison Papers: «The Current Wars»', 'Kurzer Überblick der Rutgers University. Hintergrund für Folien 5–9. '
     'Online, ca. 15 Min.'],
    ['Tesla: «My Inventions», Teil IV (1919)', 'Teslas eigene Erinnerung an seine Zeit bei Edison, mit der Stelle zu '
     'den 50’000 Dollar. Frei im Netz, englisch, ca. 20 Min.'],
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
