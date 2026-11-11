"""Playbook für die Präsentation, offene Punkte und Vorbereitungsliste (FHNW-Layout).

Aufruf: python3 scripts/playbook.py   → vorbereitung/Playbook_und_Vorbereitung_Stromkrieg.docx
"""
from fhnw_docx import PB, ROOT, build, heading, items, para, table

b = []
b.append(para('Der Stromkrieg – Playbook für die Präsentation', b=True, sz=36, jc='center', after=60))
b.append(para('10 Minuten · 18 Folien (mit Diskussion: 25) · 10. Dezember 2026', jc='center', after=240))

b.append(heading('Deine Kernbotschaft'))
b.append(para('Den Stromkrieg gewann nicht die bessere Propaganda und nicht das grössere Genie, sondern das billigere '
              'System mit dem stärkeren Kapital – und genau so wurden die USA zur Industriemacht.', after=200))

b.append(heading('Ablauf: Folie für Folie'))
b.append(table([
    ['0:00–0:40', '1–3 Intro, Titel, Auburn',
     'Intro steht, während alle Platz nehmen. Ein Klick: Der Titel morpht, Edison, Westinghouse und Tesla erscheinen. '
     'Dann: «6. August 1890: Zum ersten Mal stirbt ein Mensch auf dem elektrischen Stuhl, mit einem Generator von '
     'Westinghouse. Beschafft hatte ihn Harold P. Brown gegen den Willen von Westinghouse.»',
     '«Daraus ergibt sich meine Leitfrage.»'],
    ['0:40–1:10', '4 Leitfrage, These',
     'Leitfrage vorlesen. Klick: Die These erscheint. «Entschieden haben Kosten, Kapital und Technik.»',
     '«Dafür braucht es eine Minute Technik.»'],
    ['1:10–2:00', '5 Spannung',
     'Hohe Spannung heisst wenig Verlust. Nur Wechselstrom liess sich damals mit dem Transformator umspannen.',
     '«Wer stand hinter den beiden Systemen?»'],
    ['2:00–2:50', '6 Firmen',
     'Drei Konzerne und ein Bankier: Edison General Electric, Westinghouse, Thomson-Houston, J. P. Morgan. Tesla liefert '
     'Westinghouse 1888 die Motorpatente.',
     '«Wie lief der Streit ab?»'],
    ['2:50–3:55', '7 Verlauf I',
     'Erst Wettbewerb, ab 1888 Angstkampagne: Brown tötet Hunde, dann der elektrische Stuhl.',
     '«Entschieden hat die Kampagne aber nichts.»'],
    ['3:55–5:00', '8 Verlauf II',
     'Wendepunkt 1892: Edison General Electric fusioniert mit Thomson-Houston zu General Electric. 1893 Chicago, 1896 Niagara.',
     '«Warum also Wechselstrom?»'],
    ['5:00–5:55', '9 Warum',
     'Kosten und Kapital. Die Angstkampagne scheiterte, weil Westinghouse sie mit besserer Technik entkräftete.',
     '«Was heisst das für die USA?»'],
    ['5:55–6:45', '10 Grafik',
     'Zwischen 1880 und 1900 überholen die USA Grossbritannien. 1913: 32 % der Weltindustrie.',
     '«Was steckt hinter dieser Kurve?»'],
    ['6:45–7:25', '11 Zahlen',
     'Elektromotoren in Fabriken: 5 % → 78 %. Der Output pro Arbeitsstunde wächst danach mehr als doppelt so schnell.',
     '«Viele kennen die Geschichte anders.»'],
    ['7:25–8:15', '12 Mythen',
     'Der Elefant, «Tesla gegen Edison», der zerrissene Vertrag: Alle drei halten der Prüfung nicht stand.',
     '«Was bleibt davon bis heute?»'],
    ['8:15–9:05', '13 Und heute?',
     'Wechselstrom im Netz, Gleichstrom kehrt mit HGÜ zurück, Schweizer Spur über BBC und ABB, neue Standardkriege bei '
     'Ladesteckern.',
     '«Zurück zur Leitfrage.»'],
    ['9:05–10:00', '14–15 Fazit, Danke',
     'Beide Teile der Leitfrage beantworten: Der Wechselstrom gewann über Kosten und Kapital. Billiger Strom und '
     'Grosskonzerne trugen den Aufstieg der USA.',
     '«Danke – Fragen?»'],
    ['–', '16–18 Quellen', 'Nur zeigen, wenn jemand nach Quellen fragt.', ''],
], [1050, 1500, 4652, 2050], head=['Zeit', 'Folie', 'Das sagst du', 'Überleitung'], sz=18))
b.append(para('Die ausführlichen Sprechernotizen stehen in der PowerPoint (Ansicht → Notizen).', sz=18, after=200))

b.append(heading('Version mit Diskussion'))
b.append(para('Datei «…_mit_Diskussion»: Nach dem Fazit folgen sieben Folien (15–21): Kapitel «Diskussion», Handout Seite 1 '
              'und 2 (Fragen ans Plenum), die Gruppenaufträge (stehen lassen, solange die Gruppen arbeiten) und je eine '
              'Lösungsfolie pro Gruppe. Auf den Lösungsfolien erscheint jede Antwort erst per Klick: zuerst Antworten '
              'sammeln, dann klicken. Danach kommen «Danke» und die Quellen.', after=200))

b.append(heading('Checkliste'))
b.append(items([
    'Lerne pro Folie nur den Satz aus «Das sagst du», nicht die ganzen Notizen.',
    'Übe dreimal laut mit Stoppuhr. Wirst du zu lang, kürze Folie 11 oder lass sie weg.',
    'Die Animationen laufen automatisch; nur die These auf Folie 4 und die Lösungen der Diskussion kommen per Klick. '
    'Morph und Animationen funktionieren nur in PowerPoint (2019 oder Microsoft 365); im PDF ist alles schon sichtbar.',
    'Diese Zahlen müssen sitzen: 1882, 1888, 1890, 1892, 1893, 1896 · 32 % · 5 % → 78 %.',
    'Installiere die Schriften Libre Baskerville und Source Sans 3 (kostenlos bei Google Fonts) und nimm zusätzlich '
    'das PDF der Folien mit. Im PDF sehen die Folien auf jedem Computer richtig aus.',
    'Optional: echte Porträts von Edison, Westinghouse und Tesla in die Ovale der Titelfolie einsetzen; die Anleitung '
    'steht in den Notizen der Titelfolie.',
    'Verteile das Handout zu Beginn. Kläre vorher mit dem Dozenten, ob nach dem Vortrag Zeit für die Gruppendiskussion '
    'bleibt. Wenn nicht: Präsentation ohne Diskussion nehmen und die Gruppenfragen als «zum Weiterdenken» ankündigen.',
    'Wenn du den Faden verlierst: zurück zur Leitfrage.',
], letters=False))
b.append(PB)

b.append(heading('Offene Punkte vor der Abgabe'))
b.append(para('Das kann nur jemand mit den Büchern in der Hand erledigen. Ohne diese Punkte sind die Belege formal korrekt, '
              'aber nicht nachprüfbar.', after=100))
b.append(table([
    ['Seitenzahlen nachtragen',
     'Im Handout stehen Belege wie [1]. Ergänze bei Büchern und Aufsätzen die Seite, z. B. [1, S. 141]. Betroffen sind '
     'Jonnes [1] (Great Barrington 1886, Fusion 1892, Chicago 1893, Niagara 1896), Carlson [2] (Patente 1888, '
     'Lizenzverzicht 1891), Hughes [3] (Brown, Generator), Essig [7] (Hinrichtung 1890), Bairoch [10] (Tabelle mit den '
     'Anteilen) und Devine [11] (5 % → 78 %; 1,3 % → 3,1 %).'],
    ['Devine geht schnell',
     'Der Aufsatz steht frei im Netz: gwern.net/doc/economics/automation/1983-devine.pdf. Suche dort «78 percent» und '
     '«1.3 percent».'],
    ['Cole & Chandler prüfen',
     'Der Satz im Handout («ignorierte, bestritt, machte gegenstandslos») stammt aus Zusammenfassungen der Autoren. '
     'Prüfe ihn an deinem PDF und trage die Seite bei [8] nach.'],
    ['Online-Quellen einmal öffnen',
     'Edison Papers «The Current Wars» [4]: Stimmt die Datierung «späte 1880er, frühe 1890er»? ABB-Artikel [12]: Steht '
     'das Zitat mit dem Elefanten wörtlich so da? Joint Office [17] und IEC [14]: Titel der Seite prüfen. Diese Seiten '
     'konnte ich nicht selbst öffnen.'],
    ['Porträts',
     'Auf der Titelfolie stehen noch Buchstaben statt Bilder. Gemeinfreie Porträts gibt es auf Wikimedia Commons.'],
], [2400, 6852], first_center=False))
b.append(para('', after=160))

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
    ['Edison Papers: «The Current Wars»', 'Kurzer Überblick der Rutgers University. Hintergrund für Folien 6–8. '
     'Online, ca. 15 Min.'],
    ['Edison Papers: «Myth Buster: Topsy the Elephant»', 'Für die Mythen-Folie und Gruppe 2. Online, ca. 10 Min.'],
    ['Cole & Chandler (2019)', 'Hast du als PDF. Nur Einleitung und Schluss lesen (ca. 1 Std.): Warum Edisons '
     'Angstkampagne scheiterte. Für Folie 9.'],
    ['Jonnes: *Empires of Light* (2003)', 'Das Standardbuch zum Stromkrieg, gut lesbar. Nicht alles lesen: im '
     'Inhaltsverzeichnis die Kapitel zu 1886–1896 suchen (Great Barrington, elektrischer Stuhl, Chicago, Niagara). '
     'Ca. 2 Std. Dabei gleich die Seitenzahlen notieren.'],
    ['Deine Artikel aus Spiegel und Tages-Anzeiger', 'Deutschsprachiger Einstieg. Lies sie mit einer Frage im Kopf: '
     'Erzählen sie ein Duell zweier Genies oder die Geschichte vom Elefanten? Das ist dein Beispiel für Quellenkritik.'],
], [3000, 6252], head=['Was', 'Wofür / wie lange'], first_center=False))
b.append(para('', after=160))
b.append(heading('2. Schauen'))
b.append(table([
    ['ZDF Terra X: «Nikola Tesla – Magier der Elektrizität» (2024)', 'Auf Deutsch, in der ZDF-Mediathek. Guter Einstieg. '
     'Achte darauf, wie stark Tesla als einsames Genie gezeigt wird; das passt zur Mythen-Folie.'],
    ['PBS American Experience: «Edison» (2015)', 'Zeigt Edison als Unternehmer; der Stromkrieg ist ein Teil davon. '
     'Englisch, 114 Min. Kann ausserhalb der USA gesperrt sein.'],
    ['PBS American Experience: «Tesla» (2016)', 'Teslas Zeit bei Westinghouse. Englisch, ca. 1 Std.'],
    ['Spielfilm «Edison – Ein Leben voller Licht» (Originaltitel *The Current War*, 2017/2019)',
     'Nur zur Unterhaltung. Zeigt gut, wie Filme die Geschichte zuspitzen. Nicht als Quelle zitieren.'],
], [3000, 6252], head=['Was', 'Wofür / wie lange'], first_center=False))
b.append(para('', after=160))
b.append(heading('3. Vertiefen, falls Zeit bleibt'))
b.append(table([
    ['Carlson: *Tesla* (2013)', 'Die Kapitel zu Westinghouse 1888–1896. Belegt Lizenzverzicht und Patentverkauf.'],
    ['Hughes (1958): Aufsatz über Harold P. Brown', 'Tierversuche und die Beschaffung des Generators. 23 Seiten.'],
    ['Hughes: *Networks of Power* (1983)', 'Das wissenschaftliche Standardwerk zur Elektrifizierung. Das Kapitel zum '
     '«Battle of the Systems» reicht.'],
    ['David (1992): «Heroes, Herds and Hysteresis»', 'Wenn du mehr zu Standardkriegen sagen willst. Anspruchsvoll; '
     'die Einleitung reicht.'],
], [3000, 6252], head=['Was', 'Wofür'], first_center=False))
b.append(para('Bücher findest du über swisscovery (FHNW-Bibliothek), Aufsätze meist über den Online-Zugang der '
              'Bibliothek. Alle vollständigen Angaben stehen im Handout und in der langen Ausarbeitung.', sz=18, after=200))

b.append(heading('Zeitplan bis zum Vortrag'))
b.append(table([
    ['bis 22. November', 'Liste 1 lesen, Liste 2 schauen, Seitenzahlen nachtragen (siehe «Offene Punkte»).'],
    ['bis 3. Dezember', 'Sprechernotizen in eigene Worte umschreiben, Porträts einsetzen, offene Fragen klären.'],
    ['4.–9. Dezember', 'Dreimal laut üben, Handout drucken, Folien als PDF sichern, Technik prüfen.'],
], [2000, 7252], first_center=False))

build(ROOT / 'vorbereitung' / 'Playbook_und_Vorbereitung_Stromkrieg.docx', b, 'Der Stromkrieg – Playbook und Vorbereitung')
