"""Füllt die Folienstruktur mit dem Referat und erzeugt zwei Fassungen: mit und ohne Diskussionsblock.

Eingabe:  assets/folien_struktur.pptx (25 Folien aus der Slidesgo-Vorlage «Reconstruction Era and the Gilded Age»)
Ausgabe:  abgabe/Praesentation_Stromkrieg_Bruder.pptx und …_mit_Diskussion.pptx
Aufruf:   python3 scripts/praesentation_vorlage.py   (nach handout_kurz.py)

Quellen: Markup {schluessel} wird zu [n]. Die Nummern stammen aus assets/zitierreihenfolge.json und sind
dieselben wie im Handout. Eine Quelle, die im Handout fehlt, bricht den Lauf ab.
"""
import copy
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from animation import Timeline, rename, set_timing, set_transition
from quellen import ROOT, gemeinsame_zitate

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
NAVY, NAVY2, PAPER, WEISS = (RGBColor(0x0A, 0x1A, 0x35), RGBColor(0x0A, 0x1C, 0x5E), RGBColor(0xC2, 0xCF, 0xDE),
                             RGBColor(0xFF, 0xFF, 0xFF))
SERIF, SANS = 'Libre Baskerville', 'Source Sans 3'
STRUKTUR = ROOT / 'assets' / 'folien_struktur.pptx'
OUT_OHNE = ROOT / 'abgabe' / 'Praesentation_Stromkrieg_Bruder.pptx'
OUT_MIT = ROOT / 'abgabe' / 'Praesentation_Stromkrieg_Bruder_mit_Diskussion.pptx'

(I_INTRO, I_TITEL, I_AUBURN, I_LEITFRAGE, I_SPANNUNG, I_FIRMEN, I_V1, I_V2, I_WARUM, I_CHART, I_ZAHLEN, I_MYTHEN,
 I_HEUTE, I_FAZIT, I_KAPITEL, I_S1, I_S2, I_GRUPPEN, I_G1, I_G2, I_G3, I_DANKE, I_Q1, I_Q2, I_Q3) = range(25)
DISKUSSION = range(I_KAPITEL, I_G3 + 1)

Z = gemeinsame_zitate()
T = Z.text


# ---------- Bausteine ----------
def set_paras(tf, texts, size=None, rich=False):
    """Ersetzt den Text und behält das Format der Vorlage (Absatz i nutzt Vorlagenabsatz i).

    {schluessel} wird zur Quellennummer; mit rich=True wird *kursiv* ausgewertet (IEEE-Einträge).
    """
    body = tf._txBody
    old = body.findall(A + 'p')
    for k, text in enumerate(texts):
        p = copy.deepcopy(old[min(k, len(old) - 1)])
        runs = p.findall(A + 'r')
        for el in list(p):
            if el.tag in (A + 'r', A + 'br', A + 'fld'):
                p.remove(el)
        if runs:
            base = copy.deepcopy(runs[0])
        else:
            base = etree.Element(A + 'r')
            etree.SubElement(base, A + 't')
        rpr = base.find(A + 'rPr')
        if rpr is None:
            rpr = etree.Element(A + 'rPr')
            base.insert(0, rpr)
        rpr.set('lang', 'de-CH')
        for h in rpr.findall(A + 'hlinkClick'):
            rpr.remove(h)
        if size:
            rpr.set('sz', str(int(size * 100)))
        end = p.find(A + 'endParaRPr')
        parts = re.split(r'(\*[^*]+\*)', T(text)) if rich else [T(text)]
        for part in parts:
            if not part:
                continue
            r = copy.deepcopy(base)
            if rich:
                if part.startswith('*'):
                    r.find(A + 'rPr').set('i', '1')
                    part = part[1:-1]
                else:
                    r.find(A + 'rPr').attrib.pop('i', None)
            r.find(A + 't').text = part
            if end is not None:
                end.addprevious(r)
            else:
                p.append(r)
        body.append(p)
    for p in old:
        body.remove(p)


def shapes_by_id(slide):
    out = {}

    def walk(shs):
        for sh in shs:
            out[sh.shape_id] = sh
            if sh.shape_type == 6:
                walk(sh.shapes)
    walk(slide.shapes)
    return out


def fill(slide, mapping, notes):
    """mapping: Form-ID → Text | [Absätze] | (Text oder Absätze, Schriftgrösse)."""
    sh = shapes_by_id(slide)
    for sid, val in mapping.items():
        size = None
        if isinstance(val, tuple):
            val, size = val
        set_paras(sh[sid].text_frame, val if isinstance(val, list) else [val], size)
    slide.notes_slide.notes_text_frame.text = T(notes)


def fill_table(table, rows, size=None):
    """size: eine Grösse für alle Zeilen oder eine Liste mit einer Grösse pro Zeile."""
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            set_paras(table.cell(r, c).text_frame, [text], size[r] if isinstance(size, list) else size)


def reset(slide):
    """Entfernt alles, was ein früherer Lauf dieses Skripts hinzugefügt hat (die Struktur stammt aus dem alten Foliensatz)."""
    for sh in list(slide.shapes):
        if not sh.name.startswith('Google Shape;') and sh.name not in ('!!Titel', '!!Untertitel'):
            sh._element.getparent().remove(sh._element)
    for old in slide._element.findall(P + 'timing'):
        slide._element.remove(old)


def remove_ids(slide, ids):
    sh = shapes_by_id(slide)
    for sid in ids:
        if sid in sh:
            sh[sid]._element.getparent().remove(sh[sid]._element)


def place(shape, x, y, w=None, h=None, anchor=None):
    shape.left, shape.top = Inches(x), Inches(y)
    if w is not None:
        shape.width = Inches(w)
    if h is not None:
        shape.height = Inches(h)
    if anchor:
        shape.text_frame._txBody.find(A + 'bodyPr').set('anchor', anchor)


def style_run(run, font, size, color=NAVY, bold=False, italic=False):
    run.font.name, run.font.size, run.font.bold, run.font.italic = font, Pt(size), bold, italic
    run.font.color.rgb = color


def textbox(slide, name, x, y, w, h, paras, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP):
    """paras: Liste von (Text, Schrift, Grösse, fett, kursiv)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    rename(tb, name)
    tf = tb.text_frame
    tf.word_wrap, tf.vertical_anchor = True, anchor
    tf.margin_left = tf.margin_right = Inches(0.02)
    for k, (text, font, size, bold, italic) in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = T(text)
        style_run(r, font, size, bold=bold, italic=italic)
    return tb


def slot(slide, name, x, y, text, size=22):
    """Zeichen im Rahmenfeld der Vorlage (0,81 × 0,68 Zoll, zwischen zwei kurzen Linien)."""
    return textbox(slide, name, x, y, 0.81, 0.68, [(text, SERIF, size, True, False)], anchor=MSO_ANCHOR.MIDDLE)


def medallion(slide, name, x, y, w, h, letter, bg=None, letter_color=NAVY, size=40):
    """Ovaler Bilderrahmen mit doppelter Linie. Später per Formfüllung durch ein Foto ersetzbar."""
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    rename(shp, name)
    shp.shadow.inherit = False
    shp.fill.solid()
    if bg is None:  # halbtransparentes Weiss auf dem Papierhintergrund
        shp.fill.fore_color.rgb = WEISS
        etree.SubElement(shp.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr')), qn('a:alpha')).set('val', '45000')
    else:
        shp.fill.fore_color.rgb = bg
    shp.line.color.rgb = NAVY if bg is None else PAPER
    shp.line.width = Pt(4.5 if bg is None else 3)
    shp.line._get_or_add_ln().set('cmpd', 'dbl')
    tf = shp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = letter
    style_run(r, SERIF, size, color=letter_color, bold=bg is not None)
    return shp


def space_after(tf, pts):
    for p in tf._txBody.findall(A + 'p'):
        ppr = p.find(A + 'pPr')
        if ppr is None:
            ppr = etree.Element(A + 'pPr')
            p.insert(0, ppr)
        spc = ppr.find(A + 'spcAft')
        if spc is None:  # Schema-Reihenfolge: lnSpc, spcBef, spcAft, …
            spc = etree.Element(A + 'spcAft')
            prev = [c for c in ppr if c.tag in (A + 'lnSpc', A + 'spcBef')]
            if prev:
                prev[-1].addnext(spc)
            else:
                ppr.insert(0, spc)
        for c in list(spc):
            spc.remove(c)
        etree.SubElement(spc, A + 'spcPts').set('val', str(int(pts * 100)))


def zweispalter(slide, title, l_label, links, r_label, rechts, notes, l_size=14, r_size=13, label_size=14):
    """Zwei Spalten mit Überschrift und Trennlinie (Leitfrage/These, Fazit, Fragen/Antworten)."""
    fill(slide, {2340: title, 2342: (links, l_size), 2341: (rechts, r_size)}, notes)
    sh = shapes_by_id(slide)
    for sid, x in ((2342, 1.25), (2341, 5.28)):
        place(sh[sid], x, 2.15, 3.47, 2.35, anchor='t')
        space_after(sh[sid].text_frame, 10)
    for name, x, label in (('Label links', 1.33, l_label), ('Label rechts', 5.36, r_label)):
        textbox(slide, name, x, 1.72, 3.39, 0.36, [(label, SERIF, label_size, True, False)], align=PP_ALIGN.LEFT)
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.0), Inches(1.8), Inches(5.0), Inches(4.45))
    rename(line, 'Trennlinie')
    line.line.color.rgb, line.line.width = NAVY, Pt(1)


def qa_slide(slide, title, l_label, fragen, r_label, antworten, notes):
    """Frage links (sichtbar), Antwort rechts: Jeder Absatz erscheint per Klick."""
    zweispalter(slide, title, l_label, fragen, r_label, antworten, notes)
    tl = Timeline()
    for k in range(len(antworten)):
        tl.click().add('float', 2341, dur=700, para=k)
    set_timing(slide, tl)


def staggered(slide, columns, kind='float', step=400, dur=600):
    tl = Timeline().auto()
    for k, col in enumerate(columns):
        for spid in col:
            tl.add(kind, spid, at=k * step, dur=dur)
    set_timing(slide, tl)


def drop_slides(prs, idxs):
    lst = prs.slides._sldIdLst
    ids = list(lst)
    for i in sorted(idxs, reverse=True):
        prs.part.drop_rel(ids[i].rId)
        lst.remove(ids[i])


prs = Presentation(STRUKTUR)
S = prs.slides
assert len(S) == 25, 'assets/folien_struktur.pptx muss 25 Folien haben.'
for slide in S:
    reset(slide)

# ---------- 1 Intro: nur der Titel; mit dem ersten Klick morpht er auf die Titelfolie ----------
s = S[I_INTRO]
fill(s, {890: ('Der Stromkrieg', 60), 891: 'Edison vs. Westinghouse – und die Rolle von Nikola Tesla'},
     'Startbild, während alle Platz nehmen. Erster Klick: Der Titel morpht auf die Titelfolie; danach erscheinen '
     'Edison, Westinghouse und Tesla automatisch.')
sh = shapes_by_id(s)
place(sh[890], 1.5, 1.35, 7.0, 2.0)
place(sh[891], 1.5, 3.45, 7.0)
set_timing(s, Timeline().auto().add('fade', 890, dur=1200).add('float', 891, at=900, dur=800))

# ---------- 2 Titel: Edison gegen Westinghouse, Tesla als Partner von Westinghouse ----------
s = S[I_TITEL]
fill(s, {895: ('Modul «Weltmacht USA»', 13), 890: ['Der Stromkrieg', '- 1886–1892 -'],
         891: ('Simon Bruder · Dozent: Stephan Schwarz · 10.12.2026', 13)},
     """[0:00–0:40, Folien 1–3] Begrüssung: «Ich spreche über den Stromkrieg: den Streit zwischen den Firmen von Edison und Westinghouse um Amerikas Stromnetz – und über die Rolle von Nikola Tesla.»
Animation: Nach dem Morph fliegen Edison (links) und Westinghouse (rechts) ein, dann erscheint Tesla als kleiner Kreis bei Westinghouse. Das ist Absicht: Tesla war nicht Edisons Gegner, sondern lieferte Westinghouse die Motorpatente. Erst danach weiterklicken.
Porträts einsetzen (optional): Oval anklicken → Rechtsklick → Form formatieren → Füllung → Bild- oder Texturfüllung → Einfügen; danach den Buchstaben im Oval löschen. Animation und Morph bleiben erhalten. Gemeinfreie Porträts gibt es auf Wikimedia Commons.""")
sh = shapes_by_id(s)
place(sh[895], 3.6, 1.11, 2.8)
for slide in (S[I_INTRO], S[I_TITEL]):
    t = shapes_by_id(slide)
    rename(t[890], '!!Titel')
    rename(t[891], '!!Untertitel')
edison = medallion(s, '!!Edison', 1.42, 1.62, 1.05, 1.35, 'E', size=36)
edison_name = textbox(s, 'Edison Name', 0.995, 3.02, 1.9, 0.5,
                      [('Thomas A. Edison', SERIF, 11, True, False), ('Gleichstrom', SANS, 11, False, True)])
westinghouse = medallion(s, '!!Westinghouse', 7.53, 1.62, 1.05, 1.35, 'W', size=36)
westinghouse_name = textbox(s, 'Westinghouse Name', 6.955, 3.02, 2.2, 0.5,
                            [('George Westinghouse', SERIF, 11, True, False),
                             ('Wechselstrom · mit Nikola Tesla', SANS, 11, False, True)])
tesla = medallion(s, '!!Tesla', 8.24, 1.48, 0.55, 0.55, 'T', bg=NAVY2, letter_color=WEISS, size=20)
set_timing(s, Timeline().auto()
           .add('fly_left', edison.shape_id, at=0, dur=900).add('fade', edison_name.shape_id, at=500, dur=500)
           .add('fly_right', westinghouse.shape_id, at=1000, dur=900).add('fade', westinghouse_name.shape_id, at=1500, dur=500)
           .add('grow_turn', tesla.shape_id, at=2100, dur=800))

# ---------- 3 Einstieg ----------
s = S[I_AUBURN]
fill(s, {1864: ('- Auburn, 1890 -', 54),
         1865: ('6. August: erste Hinrichtung auf dem elektrischen Stuhl – mit einem Generator von Westinghouse {essig}', 14)},
     """[0:00–0:40] Einstieg
Auburn, Staat New York, 6. August 1890: William Kemmler wird als erster Mensch auf dem elektrischen Stuhl hingerichtet {essig}. Der Generator stammt von Westinghouse. Beschafft hatte ihn aber Harold P. Brown über Dritte und gegen den Willen von Westinghouse {hughes58}. Brown wollte zeigen: Wechselstrom tötet.
Überleitung: «Daraus ergibt sich meine Leitfrage.»""")
set_timing(s, Timeline().auto().add('zoom', 1864, dur=900).add('float', 1865, at=800, dur=700))

# ---------- 4 Leitfrage und These (Wortlaut wie im Handout) ----------
s = S[I_LEITFRAGE]
zweispalter(s, 'Leitfrage und These', 'Leitfrage',
            ['Warum setzte sich in den USA der Wechselstrom gegen Edisons Gleichstrom durch – und was zeigt der '
             'Konflikt über den Aufstieg der USA zur führenden Industriemacht?'],
            'These',
            ['Entschieden haben nicht Propaganda oder ein einzelnes Genie, sondern Kosten, Kapital und Technik.',
             'Aus dem Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.'],
            """[0:40–1:10] Leitfrage und These
Leitfrage vorlesen. Dann klicken: Die These erscheint.
These in eigenen Worten: Entschieden haben nicht Propaganda oder ein einzelnes Genie, sondern Kosten, Kapital und Technik. Aus dem Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.
Zur Datierung, falls jemand fragt: Der Stromkrieg fällt in die späten 1880er- und frühen 1890er-Jahre {edpCW}; ich rechne von 1886 (erste Wechselstromnetze) bis 1892 (Gründung von General Electric).
Überleitung: «Dafür braucht es eine Minute Technik.»""", l_size=16, r_size=15)
set_timing(s, Timeline().auto().add('fade', 2342, dur=700)
           .click().add('float', 2341, dur=700, para=0).add('float', 2341, at=500, dur=700, para=1))

# ---------- 5 Technik (Tabelle) ----------
s = S[I_SPANNUNG]
fill(s, {3480: 'Warum die Spannung entscheidet'}, """[1:10–2:00] Technik
Kernidee: Strom lässt sich nur mit hoher Spannung billig über weite Strecken schicken.
Der Verlust in der Leitung ist I² · R: Er wächst mit dem Quadrat der Stromstärke. Zehnfache Spannung heisst bei gleicher Leistung ein Zehntel des Stroms und nur noch ein Hundertstel des Verlusts.
Um 1890 liess sich nur Wechselstrom mit dem Transformator hoch- und heruntersetzen. Edisons Gleichstrom mit rund 110 Volt brauchte deshalb alle paar Kilometer ein Kraftwerk und dicke Kupferleitungen {pearl,jonnes}.
Überleitung: «Wer stand hinter den beiden Systemen?»""")
frame = shapes_by_id(s)[3481]
tbl = frame.table
for tr in tbl._tbl.tr_lst[4:]:  # Zeile «Schwachpunkt» entfällt
    tbl._tbl.remove(tr)
fill_table(tbl, [
    ['', 'Gleichstrom (Edison)', 'Wechselstrom (Westinghouse)'],
    ['Spannung', 'rund 110 V, damals nicht transformierbar', 'Transformator: hoch im Netz, tief beim Kunden'],
    ['Leitungsverlust', 'hoch: viel Strom, dicke Kupferleitungen', 'klein: hohe Spannung, wenig Strom'],
    ['Reichweite', 'ein Stadtquartier (Pearl Street 1882: 0,65 km²) {pearl}', 'Niagara–Buffalo 1896: rund 40 km {jonnes}'],
], size=13)
for k, row in enumerate(tbl.rows):
    row.height = Inches(0.48 if k == 0 else 0.6)
frame.height = Inches(0.48 + 3 * 0.6)
merksatz = textbox(s, 'Merksatz', 0.79, 3.98, 8.43, 0.45,
                   [('Bei gleicher Leistung: 10-fache Spannung → 1/10 des Stroms → 1/100 des Verlusts (I² · R)', SERIF, 13, True, False)])
set_timing(s, Timeline().auto().add('float', merksatz.shape_id, at=600, dur=700))

# ---------- 6 Firmen und Geldgeber ----------
s = S[I_FIRMEN]
remove_ids(s, (2450, 2455, 2458, 2467))  # Symbolbilder der Vorlage
fill(s, {
    2429: 'Firmen und Geldgeber statt Genies',
    2434: ('Edison', 15), 2430: ('Edison General Electric: Gleichstrom aus vielen kleinen Kraftwerken {jonnes}', 13),
    2436: ('Westinghouse', 15), 2431: ('Wechselstromnetze ab 1886; ab 1888 mit Teslas Motorpatenten {carlson}', 13),
    2435: ('Thomson-Houston', 15), 2432: ('dritter grosser Anbieter unter Charles A. Coffin; auch Wechselstrom {jonnes}', 13),
    2437: ('J. P. Morgan', 15), 2433: ('Bankier; seine Bank fädelt 1892 die Fusion zu General Electric ein {jonnes}', 13),
}, """[2:00–2:50] Akteure
Es stritten nicht zwei Erfinder, sondern drei Konzerne und ihre Geldgeber {jonnes}.
Edison General Electric: Edisons Firmen setzten auf Gleichstrom. Edison hatte viel Geld und seinen Ruf in dieses System gesteckt; er hielt Hochspannung aber auch ehrlich für gefährlich.
Westinghouse Electric: George Westinghouse, bekannt durch die Druckluftbremse, baute ab 1886 Wechselstromnetze. 1888 erwarb er die Motorpatente von Nikola Tesla {carlson}. Damit konnte Wechselstrom auch Fabriken antreiben.
Thomson-Houston unter Charles A. Coffin: der dritte grosse Anbieter; er verkaufte ebenfalls Wechselstrom.
J. P. Morgan: Bankier und früher Geldgeber Edisons. Seine Bank fädelte 1892 die Fusion zu General Electric mit ein.
Überleitung: «Wie lief der Streit ab?»""")


def vier_felder(slide, zeichen):
    """Verbreitert die Textfelder des Vier-Felder-Layouts und setzt Zeichen in die Rahmenfelder."""
    sh = shapes_by_id(slide)
    for sid in (2434, 2430, 2435, 2432):  # linke Spalte
        place(sh[sid], 2.31, sh[sid].top / 914400, 2.7)
    for sid in (2436, 2431, 2437, 2433):  # rechte Spalte (rechtsbündig)
        place(sh[sid], 4.99, sh[sid].top / 914400, 2.7)
    felder = [(1.33, 1.86, 2434, 2430), (7.86, 1.86, 2436, 2431), (1.33, 3.44, 2435, 2432), (7.86, 3.44, 2437, 2433)]
    spalten = []
    for (x, y, titel, text), (z, size) in zip(felder, zeichen):
        spalten.append((slot(slide, f'Zeichen {z}', x, y, z, size).shape_id, titel, text))
    staggered(slide, spalten)


vier_felder(s, [('E', 22), ('W', 22), ('T-H', 18), ('M', 22)])


# ---------- 7 und 8 Verlauf (Zeitstrahl; Jahreszahlen stehen in den Rahmenfeldern) ----------
def zeitstrahl(slide, title, spalten, notes):
    remove_ids(slide, (4403, 4386, 4394, 4391))  # Symbolbilder der Vorlage
    ids = [(4368, 4361, 4362, 4371), (4370, 4365, 4366, 4374), (4367, 4359, 4360, 4377), (4369, 4363, 4364, 4380)]
    mapping = {4358: (title, 25)}
    for (jahr, label, text, _), (j, l, t) in zip(ids, spalten):
        mapping.update({jahr: (j, 17), label: (l, 13), text: (t, 13)})
    fill(slide, mapping, notes)
    sh = shapes_by_id(slide)
    for jahr, label, text, rahmen in ids:
        place(sh[rahmen], sh[rahmen].left / 914400, 2.0)
        place(sh[jahr], sh[jahr].left / 914400, 2.135, anchor='ctr')
        place(sh[label], sh[label].left / 914400, 2.92)
        place(sh[text], sh[text].left / 914400, 3.36, h=1.2, anchor='t')
    for pfeil in (4383, 4384, 4385):
        sh[pfeil].top = Inches(2.34)
    return [(j, l, t) for j, l, t, _ in ids]


spalten = zeitstrahl(S[I_V1], '1882–1890: Wettbewerb und Angstkampagne', [
    ('1882', 'Pearl Street', 'Edisons Gleichstrom-Kraftwerk in New York: 85 Kunden, 400 Lampen {pearl}'),
    ('1886', 'Great Barrington', 'erstes Wechselstromnetz der USA mit Transformatoren {jonnes}'),
    ('1888', 'Angstkampagne', 'Harold P. Brown tötet öffentlich Hunde mit Wechselstrom {hughes58}'),
    ('1890', 'Elektrischer Stuhl', 'erste Hinrichtung, mit einem Generator von Westinghouse {essig}'),
], """[2:50–3:55] Verlauf, Teil 1
Phase 1, offener Wettbewerb: 1882 eröffnet Edison in New York das Kraftwerk an der Pearl Street, mit anfangs rund 85 Kunden und 400 Lampen {pearl}. 1886 läuft in Great Barrington das erste Wechselstromnetz der USA mit Transformatoren; im selben Jahr gründet Westinghouse seine Elektrofirma {jonnes}.
Phase 2, Angstkampagne: 1888 tötet der Ingenieur Harold P. Brown öffentlich Hunde mit Wechselstrom, unterstützt aus Edisons Labor {hughes58}. Wichtig: Es war Brown, nicht Edison selbst.
1890 folgt die erste Hinrichtung auf dem elektrischen Stuhl {essig}. Sie verläuft qualvoll, kein «sauberer» Tod.
Überleitung: «Entschieden hat die Kampagne aber nichts.»""")
staggered(S[I_V1], spalten, step=350)
zeitstrahl(S[I_V2], '1891–1896: Markt und Kapital entscheiden', [
    ('1891', 'Lizenzverzicht', 'Westinghouse in Finanznot; Tesla verzichtet auf seine Gebühren {carlson}'),
    ('1892', 'General Electric', 'Edison General Electric fusioniert mit Thomson-Houston {jonnes}'),
    ('1893', 'Chicago', 'Westinghouse beleuchtet die Weltausstellung {jonnes}'),
    ('1896', 'Niagara–Buffalo', 'Strom über rund 40 km; GE und Westinghouse teilen ihre Patente {jonnes}'),
], """[3:55–5:00] Verlauf, Teil 2
Phase 3: Jetzt entscheiden Markt und Kapital.
1891 gerät Westinghouse in Finanznot; Tesla verzichtet auf seine Lizenzgebühren {carlson}.
Wendepunkt 1892: Edison General Electric fusioniert mit dem Rivalen Thomson-Houston zu General Electric. Eingefädelt haben das Banken um J. P. Morgan. Edisons Name verschwindet aus dem Firmennamen, die Führung übernimmt Thomson-Houston, und GE setzt auf Wechselstrom {jonnes}. Damit ist der Streit um das System entschieden.
1893 beleuchtet Westinghouse die Weltausstellung in Chicago. 1896 fliesst Strom von den Niagarafällen rund 40 km nach Buffalo; im selben Jahr teilen GE und Westinghouse ihre Patente {jonnes}. Aus Rivalen wird ein Duopol.
Überleitung: «Warum also Wechselstrom?»""")


# ---------- 9 Warum der Wechselstrom gewann ----------
def drei_spalten(slide, title, spalten, notes):
    """Drei Spalten mit den Ziffern 1–3 in den Rahmenfeldern."""
    remove_ids(slide, (2394, 2399, 2400))  # Symbolbilder der Vorlage
    mapping = {2378: title}
    for (titel, text), (t, d) in zip(((2382, 2379), (2383, 2380), (2384, 2381)), spalten):
        mapping.update({titel: (t, 15), text: (d, 14)})
    fill(slide, mapping, notes)
    cols = []
    for k, (x, titel, text) in enumerate(((1.81, 2382, 2379), (4.59, 2383, 2380), (7.38, 2384, 2381))):
        cols.append((slot(slide, f'Ziffer {k + 1}', x, 2.0, str(k + 1), 24).shape_id, titel, text))
    staggered(slide, cols, step=450)


drei_spalten(S[I_WARUM], 'Kosten und Kapital statt Angstkampagne', [
    ('Kosten', 'Hohe Spannung, wenig Verlust: wenige grosse Kraftwerke statt eines pro Quartier {jonnes}'),
    ('Kapital', '1892 geht Edisons Firma in General Electric auf; auch GE setzt auf Wechselstrom {jonnes}'),
    ('Angst scheitert', 'Westinghouse ignoriert die Kampagne, bestreitet sie und entkräftet sie mit besserer Technik {cole}'),
], """[5:00–5:55] Warum der Wechselstrom gewann
Erstens Kosten: Mit hoher Spannung kam Strom mit wenig Verlust über weite Strecken. Wenige grosse Kraftwerke reichten; Edisons Gleichstrom brauchte alle paar Kilometer ein Kraftwerk {jonnes}.
Zweitens Kapital: 1892 ging Edisons Firma in General Electric auf. Auch GE setzte auf Wechselstrom {jonnes}.
Drittens: Die Angstkampagne scheiterte. Cole und Chandler zeigen, dass Westinghouse sie zuerst ignorierte, dann die Vorwürfe bestritt und sie schliesslich mit besserer Technik gegenstandslos machte {cole}.
Überleitung: «Was heisst das für die USA?»""")

# ---------- 10 Grafik Weltindustrie ----------
s = S[I_CHART]
remove_ids(s, (4458, 4459, 4460, 4461, 4462, 4463))  # Legende der Vorlage; die Grafik beschriftet die Linien direkt
fill(s, {4456: 'Die USA überholen Grossbritannien',
         4457: ('Anteil an der Weltindustrieproduktion in %. Die Linien verbinden die Stichjahre; es sind Schätzwerte. '
                'Eigene Darstellung nach {bairoch}.', 11)},
     """[5:55–6:45] Bedeutung für die USA (1)
Bezug zum Modul «Weltmacht USA». Die Daten stammen vom Wirtschaftshistoriker Paul Bairoch {bairoch}; es sind Schätzwerte für vier Stichjahre, die Linien verbinden diese Punkte.
1880 lag Grossbritannien mit 22,9 % klar vor den USA mit 14,7 %. Zwischen 1880 und 1900, in den Jahrzehnten des Stromkriegs, ziehen die USA vorbei. 1913 stellen sie 32 % der Weltindustrieproduktion, mehr als Grossbritannien (13,6 %) und Deutschland (14,8 %) zusammen.
Vorsicht bei der Deutung: Die Grafik zeigt einen zeitlichen Zusammenhang, keinen Beweis. Strom war eine von mehreren Ursachen, neben Stahl, Eisenbahn und dem grossen Binnenmarkt.
Überleitung: «Was steckt hinter dieser Kurve?»""")
sh = shapes_by_id(s)
pic = sh[4464]
s.part.related_part(pic._element.blipFill.find(A + 'blip').get(R_EMBED))._blob = \
    (ROOT / 'assets' / 'chart_weltindustrie_folie.png').read_bytes()
place(pic, 1.2, 1.36, 7.6, 3.0)
place(sh[4457], 1.07, 4.4, 7.85, 0.36)
set_timing(s, Timeline().auto().add('wipe_left', 4464, dur=2000).add('fade', 4457, at=1800, dur=600))

# ---------- 11 Zwei Kennzahlen ----------
fill(S[I_ZAHLEN], {
    3953: ('Strom baut die Fabriken um', 27), 3952: ('Zwei Kennzahlen zur US-Industrie {devine}', 14),
    3954: '5 % → 78 %', 3955: ('Anteil der Elektromotoren an der Antriebsleistung, 1899 → 1929', 14),
    3956: '1,3 % → 3,1 %', 3957: ('Wachstum des Outputs pro Arbeitsstunde pro Jahr, vor → nach 1919', 14),
}, """[6:45–7:25] Bedeutung für die USA (2)
Beide Zahlen stammen von Devine {devine}.
1899 lieferten Elektromotoren knapp 5 % der Antriebsleistung der US-Industrie, 1929 rund 78 %.
Der Output pro Arbeitsstunde wuchs vor 1919 um 1,3 % pro Jahr, danach um 3,1 %. Grund: Erst als die Fabriken um viele einzelne Elektromotoren herum neu gebaut wurden, zahlte sich die Technik aus.
Wenn die Zeit knapp wird: Diese Folie kürzen.
Überleitung: «Viele kennen die Geschichte anders.»""")
staggered(S[I_ZAHLEN], [(3953, 3952), (3954, 3955), (3956, 3957)], step=500)

# ---------- 12 Mythen (drei Tabellen: Mythos / Befund) ----------
s = S[I_MYTHEN]
fill(s, {4469: 'Drei Mythen – und was die Quellen sagen'}, """[7:25–8:15] Mythen und Quellenkritik
Erstens der Elefant: Ein ABB-Artikel schreibt, Edison habe öffentlich «Experimente mit Hunden, Pferden und sogar einem Elefanten» gezeigt {abb}. Die Tierversuche machte aber Harold P. Brown {hughes58}. Und die Elefantin Topsy starb erst 1903, über zehn Jahre nach dem Stromkrieg; Edison war nicht beteiligt, nur seine Filmfirma filmte die Tötung {topsy}.
Zweitens «Tesla gegen Edison»: Es stritten Firmen und ihre Geldgeber. Tesla lieferte Westinghouse 1888 die Motorpatente {carlson}.
Drittens der zerrissene Vertrag: Belegt sind nur der Verzicht auf Lizenzgebühren 1891 und der spätere Verkauf der Patente {carlson}.
Überleitung: «Was bleibt davon bis heute?»""")
sh = shapes_by_id(s)
for sid, y, rows in (
    (4470, 1.45, [['Mythos: Edisons «Experimente mit Hunden, Pferden und sogar einem Elefanten» {abb}'],
                  ['Befund: Die Tiere tötete Harold P. Brown; Topsy starb erst 1903, ohne Edisons Zutun {hughes58,topsy}']]),
    (4471, 2.6, [['Mythos: Der Stromkrieg war ein Duell «Tesla gegen Edison»'],
                 ['Befund: Es stritten Firmen und ihre Geldgeber; Tesla lieferte Westinghouse 1888 die Motorpatente {carlson}']]),
    (4472, 3.75, [['Mythos: Tesla zerriss aus Grossmut seinen Vertrag mit Westinghouse'],
                  ['Befund: Belegt sind der Verzicht auf Lizenzgebühren (1891) und der spätere Verkauf der Patente {carlson}']]),
):
    fill_table(sh[sid].table, rows, size=[12, 13])
    for row in sh[sid].table.rows:
        row.height = Inches(0.42)
    sh[sid].top, sh[sid].height = Inches(y), Inches(0.84)
set_timing(s, Timeline().auto().add('fade', 4470, dur=600).add('fade', 4471, at=500, dur=600).add('fade', 4472, at=1000, dur=600))

# ---------- 13 Und heute? ----------
s = S[I_HEUTE]
remove_ids(s, (2450, 2455, 2458, 2467))
fill(s, {
    2429: 'Und heute?',
    2434: ('Wechselstrom im Netz', 14), 2430: ('weltweit aus der Steckdose: Schweiz 230 V, USA 120 V {iec}', 13),
    2436: ('Gleichstrom ist zurück', 14), 2431: ('Akkus, Solarzellen, Elektronik; HGÜ für lange Leitungen seit 1954 {gotland}', 13),
    2435: ('Schweizer Spur', 14), 2432: ('C. E. L. Brown gründet 1891 in Baden die BBC; seit 1988 ABB {steig,abbHistory}', 13),
    2437: ('Neue Standardkriege', 14), 2433: ('Ladestecker in den USA; USB-C-Pflicht in der EU seit Ende 2024 {nacs,eu}', 13),
}, """[8:15–9:05] Und heute?
Wechselstrom im Netz: Weltweit kommt Wechselstrom aus der Steckdose, in der Schweiz mit 230 Volt, in den USA mit 120 Volt {iec}. Die tiefe US-Spannung stammt aus der Anfangszeit mit Edisons rund 110 Volt.
Gleichstrom ist zurück: Akkus, Solarzellen und Elektronik arbeiten mit Gleichstrom. Für sehr lange Leitungen nutzt man HGÜ; die erste kommerzielle Verbindung moderner Bauart ging 1954 in Schweden in Betrieb, gebaut von ASEA {gotland}.
Schweizer Spur: Charles E. L. Brown von der Maschinenfabrik Oerlikon gründete 1891 in Baden die BBC {steig}. BBC und ASEA schlossen sich 1988 zu ABB zusammen {abbHistory}. Nicht verwechseln mit Harold P. Brown.
Neue Standardkriege: In den USA stellten ab 2023 fast alle Autohersteller auf den Ladestecker der Firma Tesla um {nacs}. In der EU ist USB-C für Handys seit Ende 2024 Pflicht {eu}.
Überleitung: «Zurück zur Leitfrage.»""")
vier_felder(s, [('~', 26), ('=', 26), ('CH', 18), ('vs.', 18)])

# ---------- 14 Fazit: beide Teile der Leitfrage ----------
s = S[I_FAZIT]
zweispalter(s, 'Fazit: ein prägender Standardkrieg', 'Warum Wechselstrom?',
            ['Er war billiger: hohe Spannung, wenige grosse Kraftwerke, weite Strecken.',
             'Entschieden hat das Kapital, nicht die Angstkampagne: 1892 ging Edisons Firma in General Electric auf.'],
            'Was zeigt das über die USA?',
            ['Billiger Strom und Grosskonzerne trugen den Aufstieg zur grössten Industriemacht.',
             'Standards werden früh gesetzt und bleiben lange: Pfadabhängigkeit {david92}.'],
            """[9:05–10:00] Fazit
Zurück zur Leitfrage, erster Teil: Der Wechselstrom gewann, weil er billiger war und weil das Kapital 1892 darauf setzte. Die Angstkampagne hat daran nichts geändert.
Zweiter Teil: Billiger Strom und Grosskonzerne wie General Electric und Westinghouse trugen den Aufstieg der USA zur grössten Industriemacht.
Der Stromkrieg ist ein früher, prägender Standardkrieg. Er ist nicht der erste: Britische Bahngesellschaften stritten schon in den 1840er-Jahren um die Spurweite.
Schluss: «Standards werden früh gesetzt und bleiben lange. Darum kommt in den USA bis heute eine andere Spannung aus der Steckdose als bei uns.»""",
            l_size=14, r_size=14, label_size=13)
set_timing(s, Timeline().auto().add('float', 2342, dur=800).add('float', 2341, at=700, dur=800))

# ---------- Diskussionsblock (nur Fassung «mit Diskussion») ----------
s = S[I_KAPITEL]
fill(s, {1874: '?', 1873: 'Diskussion', 1875: 'Handout · Fragen und Gruppen 1–3'},
     """[ab 10:00] Diskussion – Dauer vorher mit dem Dozenten absprechen.
Ablauf: zuerst die Fragen von Handout Seite 1 und 2 kurz im Plenum, dann die Gruppenaufträge von Seite 3.""")
set_timing(s, Timeline().auto().add('grow_turn', 1874, dur=800).add('fade', 1873, at=600, dur=600)
           .add('float', 1875, at=1000, dur=600))
qa_slide(S[I_S1], 'Handout S. 1: Strom im Alltag', 'Fragen',
         ['a) Welche Geräte in deinem Alltag laufen mit Gleichstrom, welche mit Wechselstrom?',
          'b) Kennst du ein aktuelles Beispiel, bei dem sich Firmen um einen technischen Standard streiten?'],
         'Antworten',
         ['a) Meist beides: Das Netz liefert Wechselstrom. Elektronik, LED-Lampen und alles mit Akku arbeiten intern '
          'mit Gleichstrom; das Netzteil wandelt um.',
          'b) Ladestecker für E-Autos in den USA {nacs}, USB-C gegen Lightning {eu}; früher VHS gegen Betamax.'],
         """Fragen a) und b) ins Plenum stellen, zwei, drei Antworten sammeln. Dann pro Klick eine Lösung einblenden.
Ergänzung zu a): Direkt mit Wechselstrom laufen klassische Motoren und Heizgeräte, zum Beispiel der Wasserkocher. Solarzellen liefern Gleichstrom.""")
qa_slide(S[I_S2], 'Handout S. 2: Bedeutung für die USA', 'Fragen',
         ['a) Warum war billiger Strom für den Aufstieg der USA so wichtig?',
          'b) Edison war der berühmtere Erfinder – trotzdem verlor er. Was sagt das über die Rolle von Geld und Firmen aus?'],
         'Antworten',
         ['a) Elektromotoren trieben die Fabriken an: 1899 knapp 5 %, 1929 rund 78 % der Antriebsleistung {devine}. '
          'Die Produktivität stieg.',
          'b) Ruhm reicht nicht: Kosten und Kapital entschieden. 1892 ging Edisons Firma in General Electric auf; '
          'die Führung übernahm Thomson-Houston {jonnes}.'],
         """Wie bei Seite 1: erst Antworten sammeln, dann pro Klick einblenden.
a) knüpft an die Folie mit den zwei Kennzahlen an. b) knüpft an 1892 an: Fusion zu General Electric. Typisch für das Gilded Age: Banken und Grosskonzerne bestimmen den Markt, nicht einzelne Erfinder.""")
drei_spalten(S[I_GRUPPEN], 'Handout S. 3: Diskussion in Gruppen', [
    ('Gruppe 1', 'Angst als Werbung: Wo arbeiten Werbung oder Kampagnen heute mit Angst?'),
    ('Gruppe 2', 'Elefantin Topsy: Warum halten sich falsche Geschichten so lange?'),
    ('Gruppe 3', 'Standardkriege: Wer gewinnt – die bessere Technik oder die stärkere Firma?'),
], """Gruppen einteilen; die Aufträge stehen auf Handout Seite 3. Diese Folie stehen lassen, solange die Gruppen diskutieren.
Danach pro Gruppe die Lösungsfolie: Die Gruppe stellt zuerst ihr Ergebnis vor, dann Klick für Klick ergänzen.""")
qa_slide(S[I_G1], 'Gruppe 1: Angst als Werbung', 'Auftrag',
         ['Harold P. Brown tötete öffentlich Tiere, um Wechselstrom gefährlich wirken zu lassen. Wo gibt es heute '
          'Werbung oder Kampagnen, die mit Angst arbeiten?'],
         'Lösung',
         ['Beispiele: Werbung, die ein Konkurrenzprodukt schlechtmacht; Angstkampagnen im Wahlkampf.',
          'Angst wirkt kurzfristig, setzt sich aber selten gegen klaren Nutzen und tiefere Kosten durch.',
          'So auch im Stromkrieg: Westinghouse entkräftete die Vorwürfe mit besserer Technik {cole}.'],
         'Gruppe 1 stellt zuerst vor. Danach Klick für Klick ergänzen.')
qa_slide(S[I_G2], 'Gruppe 2: Die Elefantin Topsy', 'Auftrag',
         ['Der ABB-Artikel erzählt die Geschichte vom Elefanten, die Edison Papers widerlegen sie. Warum halten sich '
          'solche Geschichten so lange?'],
         'Lösung',
         ['Die Geschichte hat einen klaren Bösewicht und bleibt hängen.',
          'Edisons Filmfirma filmte die Tötung 1903. Das wirkt wie ein Beweis, ist aber keiner {topsy}.',
          'Websites und Firmen übernehmen sie ungeprüft {abb}. Lehre: Primärquellen prüfen.'],
         'Gruppe 2 stellt zuerst vor. Danach Klick für Klick ergänzen.')
qa_slide(S[I_G3], 'Gruppe 3: Standardkriege heute', 'Auftrag',
         ['Welche Standardkriege kennt ihr heute (Ladestecker, Streaming, Betriebssysteme)? Wer gewinnt meistens: '
          'die bessere Technik oder die stärkere Firma?'],
         'Lösung',
         ['Beispiele: VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker, iOS gegen Android.',
          'Oft gewinnt, wer mehr Partner, tiefere Kosten und mehr Kapital hat.',
          'Ein gesetzter Standard bleibt lange: Pfadabhängigkeit {david92}. Beim Stromkrieg sprachen Technik und '
          'Kosten für den Wechselstrom.'],
         'Gruppe 3 stellt zuerst vor. Danach Klick für Klick ergänzen.')

# ---------- Danke ----------
s = S[I_DANKE]
fill(s, {4515: 'Danke!', 4516: ['Fragen?', 'Simon Bruder · Modul «Weltmacht USA» (FHNW)', 'Dozent: Stephan Schwarz · 10.12.2026']},
     'Danke – Fragen? Die Quellen stehen auf den nächsten drei Folien; nur bei Nachfragen zeigen.')
remove_ids(s, (4845, 4848, 4851, 4854, 4855, 4860))  # Social-Media-Symbole der Vorlage
set_timing(s, Timeline().auto().add('zoom', 4515, dur=800).add('float', 4516, at=600, dur=700))

# ---------- Quellen: vollständige Liste, Nummern wie im Handout ----------
eintraege = [f'[{nr}] {text}' for nr, text in Z.eintraege()]
assert len(eintraege) <= 18, 'Mehr als 18 Quellen: eine vierte Quellenfolie in die Struktur aufnehmen.'
for k, idx in enumerate((I_Q1, I_Q2, I_Q3)):
    s = S[idx]
    teil = eintraege[6 * k:6 * k + 6]
    fill(s, {2340: f'Quellen ({k + 1}/3)'},
         'Anhang – nur bei Nachfragen zeigen. Vollständige Liste im IEEE-Format; die Nummern sind dieselben wie im Handout.')
    sh = shapes_by_id(s)
    for sid, x, zeilen in ((2342, 1.25, teil[:3]), (2341, 5.28, teil[3:])):
        set_paras(sh[sid].text_frame, zeilen or [''], 10, rich=True)
        place(sh[sid], x, 1.72, 3.47, 3.0, anchor='t')
        space_after(sh[sid].text_frame, 7)
    textbox(s, 'Hinweis', 1.25, 1.27, 7.5, 0.3, [('IEEE-Format · Nummern wie im Handout', SANS, 11, False, True)])

# ---------- Übergänge ----------
for i, slide in enumerate(S):
    if i not in (I_TITEL, I_V2):
        set_transition(slide, 'fade')
set_transition(S[I_TITEL], 'morph', 1600)  # Intro → Titel
set_transition(S[I_V2], 'morph', 1200)     # Zeitstrahl Teil 1 → Teil 2

# ---------- Aufräumen ----------
for slide in S:  # Die Vorlage nennt an einzelnen Stellen die ältere Schrift «Source Sans Pro».
    for el in slide._element.iter(A + 'latin', A + 'ea', A + 'cs', A + 'sym'):
        if el.get('typeface') == 'Source Sans Pro':
            el.set('typeface', SANS)
prs.core_properties.title = 'Der Stromkrieg – Edison vs. Westinghouse und die Rolle von Nikola Tesla'
prs.core_properties.author = 'Simon Bruder'

prs.save(OUT_MIT)
print('OK', OUT_MIT, len(prs.slides), 'Folien')
prs = Presentation(OUT_MIT)
drop_slides(prs, DISKUSSION)
prs.save(OUT_OHNE)
print('OK', OUT_OHNE, len(prs.slides), 'Folien')
