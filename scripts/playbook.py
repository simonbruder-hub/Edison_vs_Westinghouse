"""Playbook für die Präsentation und Vorbereitungsliste (gleiches FHNW-Layout).

Aufruf: python3 scripts/playbook.py <vorlage.docx> vorbereitung/Playbook_und_Vorbereitung_Stromkrieg.docx
"""
import sys

from fhnw_docx import PB, build, heading, items, para, table

src, out = sys.argv[1], sys.argv[2]

b = []
b.append(para('Der Stromkrieg – Playbook für die Präsentation', b=True, sz=36, jc='center', after=60))
b.append(para('10 Minuten · 13 Folien · 10. Dezember 2026', jc='center', after=240))

b.append(heading('Deine Kernbotschaft'))
b.append(para('Den Stromkrieg gewann nicht die bessere Propaganda und nicht das grössere Genie, sondern das billigere '
              'System mit dem stärkeren Kapital – und genau so wurden die USA zur Industriemacht.', after=200))

b.append(heading('Ablauf: Folie für Folie'))
b.append(table([
    ['0:00–0:45', '1–2 Titel, Auburn',
     '«6. August 1890: Zum ersten Mal stirbt ein Mensch auf dem elektrischen Stuhl – mit Strom von Westinghouse. '
     'Den Generator hatte aber Edisons Lager beschafft.» Dann die Leitfrage: Warum gewann trotzdem der Wechselstrom?',
     '«Dafür braucht es eine Minute Technik.»'],
    ['0:45–1:45', '3 Spannung',
     'Hohe Spannung = wenig Verlust. Nur Wechselstrom liess sich damals mit dem Transformator umspannen.',
     '«Wer stand hinter den beiden Systemen?»'],
    ['1:45–3:00', '4 Akteure',
     'Edison, Westinghouse, Tesla, Brown – eigentlich stritten aber Firmen und ihre Geldgeber.',
     '«Wie lief der Streit ab?»'],
    ['3:00–4:15', '5 Verlauf I',
     'Erst Wettbewerb, ab 1888 Propaganda: Brown tötet Hunde, dann der elektrische Stuhl.',
     '«Entschieden hat die Propaganda aber nichts.»'],
    ['4:15–5:30', '6 Verlauf II',
     'Wendepunkt 1892: J. P. Morgan gründet GE, Edison verliert die Kontrolle. 1893 Chicago, 1896 Niagara.',
     '«Warum also Wechselstrom?»'],
    ['5:30–6:30', '7 Warum',
     'Kosten und Kapital. Die Angstkampagne scheiterte. Schweizer Bezug: Oerlikon und BBC Baden, heute ABB.',
     '«Was heisst das für die USA?»'],
    ['6:30–7:30', '8 Grafik',
     'Genau in dieser Zeit überholen die USA Grossbritannien. 1913: 32 % der Weltindustrie.',
     '«Was steckt hinter dieser Kurve?»'],
    ['7:30–8:30', '9 Zahlen',
     'Elektromotoren in Fabriken: 5 % → 78 %. GE gehört 1896 zu den ersten Werten des Dow Jones.',
     '«Viele kennen die Geschichte anders.»'],
    ['8:30–9:30', '10 Mythen',
     'Topsy, «Tesla gegen Edison», der zerrissene Vertrag – alle drei sind nicht belegt.',
     '«Was bleibt davon bis heute?»'],
    ['9:30–10:00', '11–12 Fazit, Danke',
     'Gewonnen hat das billigere System mit dem stärkeren Kapital. Heute: Wechselstrom aus jeder Steckdose, neue '
     'Standardkriege (Ladestecker), und Gleichstrom kommt mit HGÜ zurück.',
     '«Danke – Fragen?»'],
    ['–', '13 Quellen', 'Nur zeigen, wenn jemand nach Quellen fragt.', ''],
], [1050, 1500, 4652, 2050], head=['Zeit', 'Folie', 'Das sagst du', 'Überleitung'], sz=18))
b.append(para('Die ausführlichen Sprechernotizen stehen in der PowerPoint (Ansicht → Notizen).', sz=18, after=200))

b.append(heading('Checkliste'))
b.append(items([
    'Lerne pro Folie nur den Satz aus «Das sagst du», nicht die ganzen Notizen.',
    'Übe dreimal laut mit Stoppuhr. Wirst du zu lang, kürze Folie 9 oder lass sie weg.',
    'Diese Zahlen müssen sitzen: 1882, 1888, 1890, 1892, 1893, 1896 · 32 % · 5 % → 78 %.',
    'Installiere die Schriften Libre Baskerville und Source Sans 3 (kostenlos bei Google Fonts) oder nimm zusätzlich '
    'ein PDF der Folien mit. Im PDF sehen die Folien auf jedem Computer richtig aus.',
    'Verteile das Handout zu Beginn, damit alle die Begriffe vor sich haben. Kläre vorher mit dem Dozenten, ob nach '
    'dem Vortrag Zeit für die Gruppendiskussion bleibt.',
    'Wenn du den Faden verlierst: zurück zur Leitfrage.',
], letters=False))
b.append(para('', after=120))

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
     'Weniger heftig. 1891 zeigte die Übertragung von Lauffen nach Frankfurt über 175 km, dass Drehstrom funktioniert – '
     'mit der Maschinenfabrik Oerlikon. Im selben Jahr gründete Charles Brown in Baden die BBC, heute ABB.'),
    ('Was hat das mit der Weltmacht USA zu tun?',
     'Billiger Strom trieb die Fabriken an, und aus dem Streit entstanden Grosskonzerne wie GE. Beides war typisch für '
     'den Aufstieg der USA zur grössten Industriemacht.'),
    ('Woher weisst du, dass die Topsy-Geschichte falsch ist?',
     'Aus den Edison Papers der Rutgers University, einem Archiv mit Edisons Briefen und Akten. Es gibt keinen Beleg für '
     'Edisons Beteiligung; nur seine Filmfirma filmte die Tötung 1903.'),
]:
    b.append(para(q, b=True, after=20, keep=True))
    b.append(para(a, after=120))
b.append(para('Weisst du etwas nicht: «Gute Frage, das habe ich nicht untersucht – ich schaue nach.» Das ist besser als raten.',
              after=0))
b.append(PB)

b.append(para('Vorbereitung: Was lesen, was schauen?', b=True, sz=28, after=200))
b.append(heading('1. Zuerst lesen (zusammen etwa 4 Stunden)'))
b.append(table([
    ['Edison Papers: «The Current Wars»', 'Kurzer, zuverlässiger Überblick der Rutgers University. Grundlage für Folien 4–6. '
     'Online, ca. 15 Min.'],
    ['Edison Papers: «Myth Buster: Topsy the Elephant»', 'Für die Mythen-Folie und Gruppe 2. Online, ca. 10 Min.'],
    ['Cole & Chandler (2019)', 'Hast du als PDF. Nur Einleitung und Schluss lesen (ca. 1 Std.): Warum Edisons '
     'Angstkampagne scheiterte. Für Folie 7.'],
    ['Jonnes: *Empires of Light* (2003)', 'Das Standardbuch zum Stromkrieg, gut lesbar. Nicht alles lesen: im '
     'Inhaltsverzeichnis die Kapitel zu 1886–1896 suchen (Great Barrington, elektrischer Stuhl, Chicago, Niagara). '
     'Ca. 2 Std.'],
    ['Deine Artikel aus Spiegel und Tages-Anzeiger', 'Deutschsprachiger Einstieg. Lies sie mit einer Frage im Kopf: '
     'Erzählen sie ein Duell zweier Genies oder den Topsy-Mythos? Das ist dein Beispiel für Quellenkritik.'],
], [3000, 6252], head=['Was', 'Wofür / wie lange'], first_center=False))
b.append(para('', after=160))
b.append(heading('2. Schauen'))
b.append(table([
    ['ZDF Terra X: «Nikola Tesla – Magier der Elektrizität» (2024)', 'Auf Deutsch, in der ZDF-Mediathek. Guter Einstieg. '
     'Achte darauf, wie stark Tesla als einsames Genie gezeigt wird – passt zur Mythen-Folie.'],
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
    ['Hughes (1958): Aufsatz über Harold P. Brown', 'Falls Fragen zu den Tierversuchen kommen. 23 Seiten.'],
    ['David (1992): «Heroes, Herds and Hysteresis»', 'Wenn du mehr zu Standardkriegen sagen willst. Anspruchsvoll – '
     'die Einleitung reicht.'],
], [3000, 6252], head=['Was', 'Wofür'], first_center=False))
b.append(para('Bücher findest du über swisscovery (FHNW-Bibliothek), Aufsätze meist über den Online-Zugang der '
              'Bibliothek. Alle vollständigen Angaben stehen in der langen Ausarbeitung.', sz=18, after=200))

b.append(heading('Zeitplan bis zum Vortrag'))
b.append(table([
    ['bis 8. November', 'Liste 1 lesen und Liste 2 schauen, Stichworte notieren.'],
    ['bis 29. November', 'Sprechernotizen in eigene Worte umschreiben, offene Fragen klären.'],
    ['1.–9. Dezember', 'Dreimal laut üben, Handout drucken, Folien als PDF sichern, Technik prüfen.'],
], [2000, 7252], first_center=False))

build(src, out, b, 'Der Stromkrieg – Playbook und Vorbereitung')
