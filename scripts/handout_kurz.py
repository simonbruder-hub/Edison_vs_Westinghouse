"""Kurzes Handout im Stil von Simons FHNW-Handout (Kopf-/Fusszeile und Logo aus dessen .docx).

Aufruf: python3 scripts/handout_kurz.py <vorlage.docx> abgabe/Handout_Stromkrieg_Bruder.docx
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent.parent
src, out = sys.argv[1], sys.argv[2]

RPR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>{b}{i}<w:color w:val="000000"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:val="de-CH"/>'


def run(text, b=False, i=False, sz=20):
    rpr = RPR.format(b='<w:b/>' if b else '', i='<w:i/>' if i else '', sz=sz)
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def rich(text, sz=20):
    """**fett** und *kursiv* im Fliesstext."""
    parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
    return ''.join(run(p[2:-2], b=True, sz=sz) if p.startswith('**') else
                   run(p[1:-1], i=True, sz=sz) if p.startswith('*') else
                   run(p, sz=sz) for p in parts if p)


def para(text='', b=False, sz=20, jc=None, after=120, ind=None, keep=False):
    ppr = f'<w:spacing w:before="0" w:after="{after}" w:line="264" w:lineRule="auto"/>'
    if ind:
        ppr += f'<w:ind w:left="{ind[0]}" w:hanging="{ind[1]}"/>'
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    if keep:
        ppr = '<w:keepNext/>' + ppr
    body = run(text, b=True, sz=sz) if b else rich(text, sz)
    return f'<w:p><w:pPr>{ppr}</w:pPr>{body}</w:p>'


def heading(text):
    return para(text, b=True, after=80, keep=True)


def items(lst, letters=True):
    out = []
    for k, t in enumerate(lst):
        mark = f'{"abcdefgh"[k]})' if letters else '–'
        out.append(f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:spacing w:after="60" w:line="264" w:lineRule="auto"/>'
                   f'<w:ind w:left="720" w:hanging="360"/></w:pPr>'
                   f'{run(mark)}<w:r><w:tab/></w:r>{rich(t)}</w:p>')
    return ''.join(out)


def table(rows, widths, head=None):
    def cell(t, w, bold=False, center=False):
        jc = '<w:jc w:val="center"/>' if center else ''
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/>{jc}</w:pPr>'
                f'{run(t, b=True) if bold else rich(t)}</w:p></w:tc>')
    border = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders>'
         f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % w for w in widths) + '</w:tblGrid>')
    if head:
        x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ''.join(cell(h, w, True, True) for h, w in zip(head, widths)) + '</w:tr>'
    for r in rows:
        x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(r[0], widths[0], False, True) + ''.join(cell(t, w) for t, w in zip(r[1:], widths[1:])) + '</w:tr>'
    return x + '</w:tbl>'


def image(rid, cx_in, cy_in, pid, name, descr):
    cx, cy = int(cx_in * 914400), int(cy_in * 914400)
    return (f'<w:p><w:pPr><w:spacing w:before="120" w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{pid}" name="{name}" descr="{escape(descr)}"/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{pid}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


PB = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

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
for k, r in enumerate(refs, 1):
    b.append(f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs><w:spacing w:after="60" w:line="240" w:lineRule="auto"/>'
             f'<w:ind w:left="567" w:hanging="567"/></w:pPr>{run(f"[{k}]", sz=18)}<w:r><w:tab/></w:r>{rich(r, sz=18)}</w:p>')
b.append(para('Alle Online-Quellen abgerufen am 1.10.2026.', sz=18, after=0))

# ---------------- Zusammenbau ----------------
tmp = Path(tempfile.mkdtemp())
with zipfile.ZipFile(src) as z:
    z.extractall(tmp)
for p in tmp.rglob('*'):
    if p.is_symlink():
        p.unlink()
doc = (tmp / 'word/document.xml').read_text(encoding='utf-8')
sect = re.search(r'<w:sectPr\b.*?</w:sectPr>', doc, re.S).group(0)
doc = re.sub(r'<w:body>.*</w:body>', '<w:body>' + ''.join(b) + sect + '</w:body>', doc, flags=re.S)
(tmp / 'word/document.xml').write_text(doc, encoding='utf-8')

# Filmbilder entfernen, Grafik einfügen
rels_p = tmp / 'word/_rels/document.xml.rels'
rels = rels_p.read_text(encoding='utf-8')
for m in re.finditer(r'<Relationship [^>]*Type="[^"]*/image" Target="media/([^"]+)"/>', rels):
    (tmp / 'word/media' / m.group(1)).unlink(missing_ok=True)
rels = re.sub(r'<Relationship [^>]*Type="[^"]*/image"[^>]*/>', '', rels)
rels = rels.replace('</Relationships>', '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/chart.png"/></Relationships>')
rels_p.write_text(rels, encoding='utf-8')
shutil.copy(ROOT / 'assets/chart_weltindustrie.png', tmp / 'word/media/chart.png')
ct = (tmp / '[Content_Types].xml').read_text(encoding='utf-8')
if 'Extension="png"' not in ct:
    ct = ct.replace('<Types ', '<Types ').replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
    (tmp / '[Content_Types].xml').write_text(ct, encoding='utf-8')

# Fusszeile: Dozent
for f in (tmp / 'word').glob('footer*.xml'):
    f.write_text(f.read_text(encoding='utf-8').replace('Stefan Czarnecki', 'Stephan Schwarz'), encoding='utf-8')
# Dokumenteigenschaften
core = tmp / 'docProps/core.xml'
if core.exists():
    c = core.read_text(encoding='utf-8')
    c = re.sub(r'<dc:title>.*?</dc:title>', '<dc:title>Der Stromkrieg – Handout</dc:title>', c)
    core.write_text(c, encoding='utf-8')

Path(out).unlink(missing_ok=True)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for p in sorted(tmp.rglob('*')):
        if p.is_file() and not p.relative_to(tmp).as_posix().startswith('[trash]'):
            z.write(p, p.relative_to(tmp).as_posix())
shutil.rmtree(tmp)
print('OK', out)
