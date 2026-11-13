"""Füllt die Folienstruktur mit dem Referat und erzeugt zwei Fassungen: mit und ohne Diskussionsblock.

Eingabe:  assets/folien_struktur.pptx (25 Folien aus der Slidesgo-Vorlage «Reconstruction Era and the Gilded Age»),
          die Bilder aus latex/handout/bilder und das Titelvideo (scripts/titelvideo.py)
Ausgabe:  abgabe/Praesentation_Stromkrieg_Bruder.pptx und …_mit_Diskussion.pptx
Aufruf:   python3 scripts/praesentation_vorlage.py

Quellen: Markup {schluessel} wird zu [n]. Die Nummern folgen der ersten Nennung im Handout
(latex/handout/main.tex). Eine Quelle, die im Handout fehlt, bricht den Lauf ab.
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
BILDER = ROOT / 'latex' / 'handout' / 'bilder'
VIDEO, POSTER = ROOT / 'assets' / 'titelvideo.mp4', ROOT / 'assets' / 'titelvideo_poster.jpg'

(I_INTRO, I_TITEL, I_AUBURN, I_LEITFRAGE, I_SPANNUNG, I_FIRMEN, I_V1, I_V2, I_WARUM, I_CHART, I_ZAHLEN, I_MYTHEN,
 I_HEUTE, I_FAZIT, I_KAPITEL, I_S1, I_S2, I_GRUPPEN, I_G1, I_G2, I_G3, I_DANKE, I_Q1, I_Q2, I_Q3) = range(25)
# Die Vorlagenfolien I_S1, I_S2 und I_GRUPPEN werden nicht mehr gebraucht (keine Aufgaben, keine Gruppenarbeit).

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


def bild(slide, name, datei, x, y, w=None, h=None):
    """Bild mit feinem Rahmen in der Farbe der Vorlage."""
    pic = slide.shapes.add_picture(str(BILDER / datei), Inches(x), Inches(y),
                                   Inches(w) if w else None, Inches(h) if h else None)
    rename(pic, name)
    pic.line.color.rgb = NAVY
    pic.line.width = Pt(1.5)
    return pic


def nachweis(slide, name, text, x, y, w, align=PP_ALIGN.LEFT, farbe=NAVY):
    """Bildnachweis direkt am Bildrand (klein, kursiv)."""
    tb = textbox(slide, name, x, y, w, 0.24, [(text, SANS, 8, False, True)], align=align)
    for run in tb.text_frame.paragraphs[0].runs:
        run.font.color.rgb = farbe
    return tb


def kopie(prs, quelle):
    """Hängt eine Kopie einer Textfolie (ohne Bilder) ans Ende an."""
    neu = prs.slides.add_slide(quelle.slide_layout)
    for sh in list(neu.shapes):
        sh._element.getparent().remove(sh._element)
    for el in list(quelle.shapes._spTree)[2:]:
        neu.shapes._spTree.append(copy.deepcopy(el))
    return neu


def ordne(prs, folien):
    """Legt die Reihenfolge der Folien fest; nicht genannte Folien entfallen."""
    lst = prs.slides._sldIdLst
    el_von = {prs.part.related_part(el.rId): el for el in list(lst)}
    behalten = [el_von[f.part] for f in folien]
    for el in list(lst):
        lst.remove(el)
        if not any(el is b for b in behalten):
            prs.part.drop_rel(el.rId)
    for el in behalten:
        lst.append(el)


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
F_PRESSE = kopie(prs, S[I_Q1])   # neue Folie mit der Karikatur
F_Q4 = kopie(prs, S[I_Q3])       # vierte Quellenfolie
F_ZITATE = kopie(prs, S[I_Q2])   # neue Folie mit zwei Primärquellen

# ---------- 1 Frontfolie: Bild von 1889 als kurzes Video, darüber der Titel ----------
s = S[I_INTRO]
remove_ids(s, (896,))  # Ornamente der Vorlage; das Bild füllt den inneren Rahmen
fill(s, {890: ('Der Stromkrieg', 40), 891: ('Edison vs. Tesla – und die Rolle von George Westinghouse', 16)},
     """[0:00–0:15] Frontfolie. Sie steht, während alle Platz nehmen.
Erster Klick: Das Video startet (11 Sekunden, ohne Ton). Das Bild wird dunkel, dann gehen die Lampen nacheinander an. Dazu frei sprechen: «New York im Jahr 1889. Elektrisches Licht verändert die Stadt. Aber welcher Strom soll fliessen: Edisons Gleichstrom oder Teslas Wechselstrom?»
Zweiter Klick: Der Titel wandert auf die Titelfolie.
Bild: Stich von William Allen Rogers für «Harper's Weekly», 1889 {natgeo}. Das Abdunkeln und Aufleuchten ist eine eigene Bearbeitung; das Video beginnt und endet mit dem unveränderten Bild.
Falls das Video nicht startet: einfach weiterklicken, das Bild bleibt stehen.""")
BX, BY, BW, BH = 0.52, 0.489, 8.96, 4.647  # Bildfläche im inneren Rahmen (Seitenverhältnis des Videos)
BAND = 3.70
film = s.shapes.add_movie(str(VIDEO), Inches(BX), Inches(BY), Inches(BW), Inches(BH),
                          poster_frame_image=str(POSTER), mime_type='video/mp4')
rename(film, 'Titelvideo')
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(BX), Inches(BAND), Inches(BW), Inches(BY + BH - BAND))
rename(band, 'Titelband')
band.shadow.inherit = False
band.fill.solid()
band.fill.fore_color.rgb = NAVY
etree.SubElement(band.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr')), qn('a:alpha')).set('val', '82000')
band.line.fill.background()
tree = s.shapes._spTree
for k, el in enumerate((film._element, band._element)):  # nach hinten: Video, Band, darüber die Texte
    tree.remove(el)
    tree.insert(2 + k, el)
sh = shapes_by_id(s)
place(sh[890], 0.9, BAND + 0.04, 8.2, 0.78, anchor='ctr')
place(sh[891], 0.9, BAND + 0.80, 8.2, 0.36, anchor='ctr')
for sid, farbe in ((890, WEISS), (891, PAPER)):
    for para in sh[sid].text_frame.paragraphs:
        for run in para.runs:
            run.font.color.rgb = farbe
nachweis(s, 'Bildnachweis', 'Bild: W. A. Rogers, «Harper’s Weekly», 1889 · De Agostini/Getty Images {natgeo}',
         5.2, BAND + 1.17, 4.2, PP_ALIGN.RIGHT, PAPER)
set_timing(s, Timeline().click().add('video', film.shape_id, dur=11040))

# ---------- 2 Titel: Edison gegen Tesla, Westinghouse als Teslas Partner ----------
s = S[I_TITEL]
fill(s, {895: ('Modul «Weltmacht USA»', 13), 890: ['Der Stromkrieg', 'Edison vs. Tesla'],
         891: ('Simon Bruder · Dozent: Stephan Schwarz · 10.12.2026', 13)},
     """[0:15–0:25] Begrüssung: «Ich spreche über den Stromkrieg: den Streit zwischen Thomas Edison und Nikola Tesla um Amerikas Stromnetz – und über die Rolle von George Westinghouse.»
Animation: Nach dem Morph fliegt Edison von links ein, dann Tesla von rechts. Zuletzt erscheint W bei Tesla: Westinghouse kaufte 1888 Teslas Patente und brachte sein System auf den Markt {kaick}. Erst danach weiterklicken.
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
tesla = medallion(s, '!!Tesla', 7.53, 1.62, 1.05, 1.35, 'T', size=36)
tesla_name = textbox(s, 'Tesla Name', 6.955, 3.02, 2.2, 0.5,
                     [('Nikola Tesla', SERIF, 11, True, False), ('Wechselstrom · mit Westinghouse', SANS, 11, False, True)])
westinghouse = medallion(s, '!!Westinghouse', 8.24, 1.48, 0.55, 0.55, 'W', bg=NAVY2, letter_color=WEISS, size=20)
set_timing(s, Timeline().auto()
           .add('fly_left', edison.shape_id, at=0, dur=900).add('fade', edison_name.shape_id, at=500, dur=500)
           .add('fly_right', tesla.shape_id, at=1000, dur=900).add('fade', tesla_name.shape_id, at=1500, dur=500)
           .add('grow_turn', westinghouse.shape_id, at=2100, dur=800))

# ---------- 3 Einstieg ----------
s = S[I_AUBURN]
remove_ids(s, (1407,))  # Ornamente der Vorlage; links steht das Bild
fill(s, {1864: ('- Auburn, 1890 -', 48),
         1865: ('6. August: erste Hinrichtung auf dem elektrischen Stuhl – mit einem Generator von Westinghouse {essig}', 14)},
     """[0:25–0:50] Einstieg
Auburn, Staat New York, 6. August 1890: William Kemmler wird als erster Mensch auf dem elektrischen Stuhl hingerichtet {essig}. Der Generator stammt von Westinghouse. Beschafft hatte ihn aber Harold P. Brown über Dritte und gegen den Willen von Westinghouse {hughes58}. Brown wollte zeigen: Wechselstrom tötet, also das System, auf das Westinghouse mit Teslas Patenten setzte.
Zum Bild: kolorierte Illustration der ersten Hinrichtung auf dem elektrischen Stuhl {natgeo}.
Überleitung: «Daraus ergibt sich meine Leitfrage.»""")
sh = shapes_by_id(s)
stuhl = bild(s, 'Bild Stuhl', 'elektrischer_stuhl_1890.jpg', 1.0, 0.95, h=3.6)
nachweis(s, 'Bildnachweis', 'Bild: Granger/Cordon Press {natgeo}', 1.0, 4.58, 2.6)
place(sh[1864], 4.0, 1.5, 5.0)
place(sh[1865], 4.0, 3.08, 5.0)
sh[1866].left, sh[1866].width = Inches(3.86), Inches(5.28)
set_timing(s, Timeline().auto().add('fade', stuhl.shape_id, dur=800).add('zoom', 1864, at=400, dur=900)
           .add('float', 1865, at=1200, dur=700))

# ---------- 4 Leitfrage und These (Wortlaut wie im Handout) ----------
s = S[I_LEITFRAGE]
zweispalter(s, 'Leitfrage und These', 'Leitfrage',
            ['Edison gegen Tesla: Warum setzte sich Teslas Wechselstrom gegen Edisons Gleichstrom durch – und was '
             'zeigt der Konflikt über den Aufstieg der USA zur führenden Industriemacht?'],
            'These',
            ['Nicht der berühmtere Erfinder gewann, sondern das billigere System mit dem stärkeren Geldgeber: Teslas '
             'Wechselstrom, finanziert von Westinghouse.',
             'Aus dem Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.'],
            """[0:50–1:15] Leitfrage und These
Leitfrage vorlesen. Dann klicken: Die These erscheint.
These in eigenen Worten: Edison war der berühmtere Erfinder, Tesla hatte die bessere Technik für weite Strecken. Gewonnen hat das billigere System mit dem stärkeren Geldgeber. Aus dem Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.
Zur Datierung, falls jemand fragt: Der Stromkrieg fällt in die späten 1880er- und frühen 1890er-Jahre {edpCW}.
Überleitung: «Wer waren die beiden?»""", l_size=15, r_size=14)
set_timing(s, Timeline().auto().add('fade', 2342, dur=700)
           .click().add('float', 2341, dur=700, para=0).add('float', 2341, at=500, dur=700, para=1))

# ---------- 5 Technik (Tabelle) ----------
s = S[I_SPANNUNG]
fill(s, {3480: 'Warum die Spannung entscheidet'}, """[2:40–3:20] Technik
Kernidee: Strom lässt sich nur mit hoher Spannung billig über weite Strecken schicken.
Der Verlust in der Leitung ist I² · R: Er wächst mit dem Quadrat der Stromstärke. Zehnfache Spannung heisst bei gleicher Leistung ein Zehntel des Stroms und nur noch ein Hundertstel des Verlusts.
Um 1890 liess sich nur Wechselstrom mit dem Transformator hoch- und heruntersetzen. Edisons Gleichstrom mit rund 110 Volt brauchte deshalb alle paar Kilometer ein Kraftwerk und dicke Kupferleitungen {pearl,jonnes}.
Dazu kam Teslas Drehfeldmotor: Mit ihm konnte Wechselstrom auch Maschinen in Fabriken antreiben {kaick}.
Überleitung: «Wie lief der Streit ab?»""")
frame = shapes_by_id(s)[3481]
tbl = frame.table
for tr in tbl._tbl.tr_lst[5:]:  # Kopfzeile und vier Zeilen bleiben
    tbl._tbl.remove(tr)
fill_table(tbl, [
    ['', 'Gleichstrom (Edison)', 'Wechselstrom (Tesla)'],
    ['Spannung', 'rund 110 V, damals nicht transformierbar', 'Transformator: hoch im Netz, tief im Haus'],
    ['Leitungsverlust', 'hoch: viel Strom, dicke Kupferleitungen', 'klein: hohe Spannung, wenig Strom'],
    ['Reichweite', 'ein Quartier (Pearl Street: 0,65 km²) {pearl}', 'Niagara–Buffalo 1896: rund 40 km {jonnes}'],
    ['Motor', 'Gleichstrommotoren', 'Teslas Drehfeldmotor (Patente 1888) {kaick}'],
], size=13)
for col, breite in zip(tbl.columns, (1.57, 3.43, 3.43)):  # jede Zelle bleibt einzeilig
    col.width = Inches(breite)
for k, row in enumerate(tbl.rows):
    row.height = Inches(0.46 if k == 0 else 0.56)
frame.height = Inches(0.46 + 4 * 0.56)
merksatz = textbox(s, 'Merksatz', 0.79, 4.4, 8.43, 0.45,
                   [('Bei gleicher Leistung: 10-fache Spannung → 1/10 des Stroms → 1/100 des Verlusts (I² · R)', SANS, 14, True, False)])
set_timing(s, Timeline().auto().add('float', merksatz.shape_id, at=600, dur=700))

# ---------- 6 Bildquelle: Karikatur von Currier & Ives ----------
s = F_PRESSE
remove_ids(s, (2342,))
fill(s, {2340: 'So sah es die Presse: ein Zweikampf',
         2341: (['Schon damals machte die Presse aus dem Wettlauf um das Licht ein Duell: Edison (links) gegen '
                 'Charles Brush, den Erfinder der Bogenlampe {natgeo}.',
                 'Genauso wird der Stromkrieg bis heute erzählt: als Duell Edison gegen Tesla.'], 15)},
     """[7:05–7:25] Bildquelle
Karikatur von Currier & Ives aus dem späten 19. Jahrhundert: Edison (links) und Charles Brush, der Erfinder der Bogenlampe, ringen um die Beleuchtung von New York {natgeo}.
Quellenkritik: Die Zeichnung macht aus einem Wettbewerb von Firmen einen Zweikampf zweier Männer. Genauso wird auch der Stromkrieg bis heute erzählt: als persönliches Duell Edison gegen Tesla. Dahinter standen aber immer auch Firmen und Geldgeber.
Klick: zweiter Satz. Überleitung: «Aus solchen Zuspitzungen entstehen Mythen.»""")
sh = shapes_by_id(s)
place(sh[2341], 5.45, 1.95, 3.55, 2.6, anchor='t')
space_after(sh[2341].text_frame, 12)
karikatur = bild(s, 'Bild Karikatur', 'karikatur_edison_brush.jpg', 1.0, 1.72, w=4.1)
nachweis(s, 'Bildnachweis', 'Bild: Currier & Ives, spätes 19. Jh. · MPI/Getty Images {natgeo}', 1.0, 4.81, 4.1)
set_timing(s, Timeline().auto().add('fade', karikatur.shape_id, dur=700).add('float', 2341, at=500, dur=700, para=0)
           .click().add('float', 2341, dur=700, para=1))

# ---------- 5 Zwei Erfinder und ihre Geldgeber ----------
s = S[I_FIRMEN]
remove_ids(s, (2450, 2455, 2458, 2467))  # Symbolbilder der Vorlage
fill(s, {
    2429: 'Zwei Erfinder und ihre Geldgeber',
    2434: ('Thomas A. Edison', 15), 2430: ('Erfinder und Unternehmer; Gleichstrom aus vielen kleinen Kraftwerken {jonnes}', 13),
    2436: ('Nikola Tesla', 15), 2431: ('Ingenieur aus Smiljan; 1884 sechs Monate bei Edison, dann eigene Wege {kaick}', 13),
    2435: ('J. P. Morgan', 15), 2432: ('Bankier; finanziert Edison und fädelt 1892 die Fusion zu General Electric ein {jonnes}', 13),
    2437: ('George Westinghouse', 15), 2433: ('Unternehmer; kauft 1888 Teslas Patente und baut die Netze {kaick}', 13),
}, """[1:15–2:00] Die beiden Erfinder
Thomas Edison: Erfinder und Unternehmer, berühmt durch die Glühlampe. Seine Firmen setzten auf Gleichstrom aus vielen kleinen Kraftwerken {jonnes}.
Nikola Tesla: Ingenieur aus Smiljan im heutigen Kroatien. Er kam 1884 mit einer Empfehlung aus Paris zu Edison nach New York. Die beiden waren charakterlich zu verschieden, und Edison wollte vom Wechselstrom nichts wissen; nach sechs Monaten trennten sie sich {kaick}.
Hinter beiden stand Geld: J. P. Morgan finanzierte Edison und fädelte 1892 die Fusion zu General Electric ein {jonnes}. George Westinghouse kaufte im Juli 1888 Teslas Patente und baute die Netze {kaick}.
Überleitung: «Wie dachten die beiden übereinander?»""")


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


vier_felder(s, [('E', 22), ('T', 22), ('M', 22), ('W', 22)])

# ---------- 6 Zwei Primärquellen: Edison über Wechselstrom, Tesla über Edison ----------
s = F_ZITATE
zweispalter(s, 'Zwei Blicke aufeinander', 'Edison über Wechselstrom, 1889',
            ['«Nichts rechtfertigt Hochspannung und Wechselstrom, weder wissenschaftlich noch geschäftlich.» {edison89}'],
            'Tesla über Edison, 1931',
            ['«Hätte Edison eine Nadel im Heuhaufen suchen müssen, er hätte Halm für Halm umgedreht.»',
             '«Ein wenig Theorie und Rechnung hätten ihm 90 Prozent der Arbeit erspart.» {nyt31}'],
            """[2:00–2:40] Zwei Primärquellen
Links: Edison 1889 in der Zeitschrift «North American Review» {edison89}. Original, Wortlaut vor dem Vortrag prüfen: «There is no plea which will justify the use of high-tension and alternating currents, either in a scientific or a commercial sense.» Deutung: Edison hielt Wechselstrom für lebensgefährlich, und sein Geschäft hing am Gleichstrom.
Rechts: Tesla 1931 in der New York Times, einen Tag nach Edisons Tod {nyt31}. Original, Wortlaut prüfen: «If Edison had a needle to find in a haystack he would proceed at once with the diligence of the bee to examine straw after straw until he found the object of his search … a little theory and calculation would have saved him ninety per cent of his labor.» Deutung: Tesla sah sich als Theoretiker und Edison als Tüftler, der alles ausprobiert.
Beide Zitate sind übersetzt und gekürzt. Klick: Teslas Zitat erscheint.
Überleitung: «Wer hatte technisch recht?»""", l_size=16, r_size=16, label_size=13)
textbox(s, 'Hinweis Übersetzung', 1.25, 4.62, 7.5, 0.25, [('Zitate übersetzt', SANS, 9, False, True)])
set_timing(s, Timeline().auto().add('fade', 2342, dur=700)
           .click().add('float', 2341, dur=700, para=0).add('float', 2341, at=600, dur=700, para=1))


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


spalten = zeitstrahl(S[I_V1], '1882–1890: Vom Mitarbeiter zum Rivalen', [
    ('1882', 'Pearl Street', 'Edisons Gleichstrom-Kraftwerk in New York: 85 Kunden, 400 Lampen {pearl}'),
    ('1884', 'Tesla bei Edison', 'sechs Monate in Edisons Firma, dann trennen sie sich {kaick}'),
    ('1888', 'Teslas Patente', 'Westinghouse kauft im Juli Teslas Wechselstrompatente {kaick}'),
    ('1890', 'Elektrischer Stuhl', 'Edisons Lager will zeigen: Wechselstrom tötet {hughes58,essig}'),
], """[3:20–4:15] Verlauf, Teil 1
1882 eröffnet Edison in New York das Kraftwerk an der Pearl Street, mit anfangs rund 85 Kunden und 400 Lampen {pearl}.
1884 kommt Tesla zu Edison. Nach sechs Monaten trennen sie sich; Edison will vom Wechselstrom nichts wissen {kaick}. Tesla gründet eigene Firmen und entwickelt seinen Wechselstrommotor.
1888 kauft Westinghouse im Juli Teslas Wechselstrompatente {kaick}. Im selben Jahr startet Edisons Lager die Angstkampagne: Der Ingenieur Harold P. Brown tötet öffentlich Hunde mit Wechselstrom, unterstützt aus Edisons Labor {hughes58}. Wichtig: Es war Brown, nicht Edison selbst.
1890 folgt die erste Hinrichtung auf dem elektrischen Stuhl {essig}. Sie verläuft qualvoll, kein «sauberer» Tod.
Überleitung: «Entschieden hat die Kampagne aber nichts.»""")
staggered(S[I_V1], spalten, step=350)
zeitstrahl(S[I_V2], '1891–1896: Teslas System setzt sich durch', [
    ('1891', 'Lizenzverzicht', 'Westinghouse in Finanznot; Tesla verzichtet auf seine Gebühren {carlson}'),
    ('1892', 'General Electric', 'Edisons Firma geht in GE auf; auch GE setzt auf Wechselstrom {jonnes}'),
    ('1893', 'Chicago', 'Teslas System versorgt die Weltausstellung {kaick}'),
    ('1896', 'Niagara–Buffalo', 'Teslas System bringt Strom über rund 40 km {kaick,jonnes}'),
], """[4:15–5:10] Verlauf, Teil 2
Jetzt entscheiden Markt und Kapital.
1891 gerät Westinghouse in Finanznot; Tesla verzichtet auf seine Lizenzgebühren {carlson}.
Wendepunkt 1892: Edison General Electric fusioniert mit dem Rivalen Thomson-Houston zu General Electric. Eingefädelt haben das Banken um J. P. Morgan. Edisons Name verschwindet aus dem Firmennamen, die Führung übernimmt Thomson-Houston, und GE setzt auf Wechselstrom {jonnes}. Damit hat Edison den Streit um das System verloren.
1893 versorgt Westinghouse die Weltausstellung in Chicago mit Teslas Wechselstromsystem; im selben Jahr wird Teslas System für die Niagarafälle gewählt {kaick}. 1896 fliesst Strom von dort rund 40 km nach Buffalo; GE und Westinghouse teilen ihre Patente {jonnes}. Aus Rivalen wird ein Duopol.
Überleitung: «Warum also Teslas System?»""")


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


drei_spalten(S[I_WARUM], 'Warum Teslas System gewann', [
    ('Kosten', 'Hohe Spannung, wenig Verlust: wenige grosse Kraftwerke statt eines pro Quartier {jonnes}'),
    ('Kapital', 'Westinghouse bringt Teslas System auf den Markt; 1892 geht Edisons Firma in GE auf {kaick,jonnes}'),
    ('Angst scheitert', 'Westinghouse ignoriert die Kampagne, bestreitet sie und entkräftet sie mit besserer Technik {cole}'),
], """[5:10–5:55] Warum Teslas System gewann
Erstens Kosten: Mit hoher Spannung kam Strom mit wenig Verlust über weite Strecken. Wenige grosse Kraftwerke reichten; Edisons Gleichstrom brauchte alle paar Kilometer ein Kraftwerk {jonnes}.
Zweitens Kapital: Westinghouse brachte Teslas System in die Praxis {kaick}. 1892 ging Edisons Firma in General Electric auf, und auch GE setzte auf Wechselstrom {jonnes}.
Drittens: Die Angstkampagne scheiterte. Cole und Chandler zeigen, dass Westinghouse sie zuerst ignorierte, dann die Vorwürfe bestritt und sie schliesslich mit besserer Technik gegenstandslos machte {cole}.
Überleitung: «Was heisst das für die USA?»""")

# ---------- 10 Grafik Weltindustrie ----------
s = S[I_CHART]
remove_ids(s, (4458, 4459, 4460, 4461, 4462, 4463))  # Legende der Vorlage; die Grafik beschriftet die Linien direkt
fill(s, {4456: 'Die USA überholen Grossbritannien',
         4457: ('Anteil an der Weltindustrieproduktion in %. Linien verbinden Stichjahre (Schätzwerte). '
                'Eigene Darstellung nach {bairoch}.', 11)},
     """[5:55–6:35] Bedeutung für die USA (1)
Bezug zum Modul «Weltmacht USA». Die Daten stammen vom Wirtschaftshistoriker Paul Bairoch {bairoch}; es sind Schätzwerte für vier Stichjahre, die Linien verbinden diese Punkte.
1880 lag Grossbritannien mit 22,9 % klar vor den USA mit 14,7 %. Zwischen 1880 und 1900, in den Jahrzehnten des Stromkriegs, ziehen die USA vorbei. 1913 stellen sie 32 % der Weltindustrieproduktion, mehr als Grossbritannien (13,6 %) und Deutschland (14,8 %) zusammen.
Vorsicht bei der Deutung: Die Grafik zeigt einen zeitlichen Zusammenhang, keinen Beweis. Strom war eine von mehreren Ursachen, neben Stahl, Eisenbahn und dem grossen Binnenmarkt.
Überleitung: «Was steckt hinter dieser Kurve?»""")
sh = shapes_by_id(s)
pic = sh[4464]
s.part.related_part(pic._element.blipFill.find(A + 'blip').get(R_EMBED))._blob = \
    (ROOT / 'assets' / 'chart_weltindustrie_folie.png').read_bytes()
place(pic, 1.2, 1.36, 7.6, 3.0)
place(sh[4457], 0.79, 4.4, 8.43, 0.36)
set_timing(s, Timeline().auto().add('wipe_left', 4464, dur=2000).add('fade', 4457, at=1800, dur=600))

# ---------- 11 Zwei Kennzahlen ----------
fill(S[I_ZAHLEN], {
    3953: ('Strom baut die Fabriken um', 27), 3952: ('Zwei Kennzahlen zur US-Industrie {devine}', 14),
    3954: '5 % → 78 %', 3955: ('Anteil der Elektromotoren an der Antriebsleistung, 1899 → 1929', 14),
    3956: '1,3 % → 3,1 %', 3957: ('Wachstum des Outputs pro Arbeitsstunde pro Jahr, vor → nach 1919', 14),
}, """[6:35–7:05] Bedeutung für die USA (2)
Beide Zahlen stammen von Devine {devine}.
1899 lieferten Elektromotoren knapp 5 % der Antriebsleistung der US-Industrie, 1929 rund 78 %.
Der Output pro Arbeitsstunde wuchs vor 1919 um 1,3 % pro Jahr, danach um 3,1 %. Grund: Erst als die Fabriken um viele einzelne Elektromotoren herum neu gebaut wurden, zahlte sich die Technik aus.
Wenn die Zeit knapp wird: Diese Folie kürzen.
Überleitung: «Wie haben die Zeitgenossen den Streit gesehen?»""")
sh = shapes_by_id(S[I_ZAHLEN])
for sid in (3953, 3952, 3955, 3957):
    place(sh[sid], 1.8, sh[sid].top / 914400, 6.4)
staggered(S[I_ZAHLEN], [(3953, 3952), (3954, 3955), (3956, 3957)], step=500)

# ---------- 12 Mythen (drei Tabellen: Mythos / Befund) ----------
s = S[I_MYTHEN]
fill(s, {4469: 'Drei Mythen – und was die Quellen sagen'}, """[7:25–8:15] Mythen und Quellenkritik
Erstens der Elefant: Ein ABB-Artikel schreibt, Edison habe öffentlich «grausame Experimente mit Hunden, Pferden und sogar einem Elefanten» gezeigt {abb}. Die Tierversuche machte aber Harold P. Brown, unterstützt von Edisons Labor {hughes58}. Und die Elefantin Topsy starb erst 1903, über zehn Jahre nach dem Stromkrieg; Edison war nicht beteiligt, nur seine Filmfirma filmte die Tötung {topsy}.
Zweitens die 50’000 Dollar: Oft heisst es, Edison habe Tesla für bessere Dynamos 50’000 Dollar versprochen und danach behauptet, das sei ein Scherz gewesen. In Teslas eigenem Bericht von 1919 steht aber: «The Manager had promised me fifty thousand dollars on the completion of this task but it turned out to be a practical joke.» Es war also «der Manager», nicht Edison {tesla19}. Die bekannte Szene mit Edison ist eine spätere Zuspitzung.
Drittens der zerrissene Vertrag: Belegt sind nur der Verzicht auf Lizenzgebühren 1891 und der spätere Verkauf der Patente {carlson}.
Überleitung: «Was bleibt von Edison und Tesla bis heute?»""")
sh = shapes_by_id(s)
for sid, y, rows in (
    (4470, 1.45, [['Mythos: Edisons «Experimente mit Hunden, Pferden und sogar einem Elefanten» {abb}'],
                  ['Befund: Die Tiere tötete Harold P. Brown; Topsy starb erst 1903, ohne Edisons Zutun {hughes58,topsy}']]),
    (4471, 2.6, [['Mythos: Edison prellte Tesla um 50’000 Dollar'],
                 ['Befund: In Teslas eigenem Bericht (1919) versprach «der Manager» das Geld, nicht Edison {tesla19}']]),
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
    2434: ('Teslas Wechselstrom', 14), 2430: ('aus jeder Steckdose: Schweiz 230 V, USA 120 V {iec}', 13),
    2436: ('Edisons Comeback', 14), 2431: ('Akkus, Solarzellen, Elektronik; HGÜ für lange Leitungen seit 1954 {gotland}', 13),
    2435: ('Tesla als Name', 14), 2432: ('Einheit der Flussdichte seit 1960; Autohersteller seit 2003 {kaick}', 13),
    2437: ('Neuer Standardkrieg', 14), 2433: ('Der Ladestecker der Firma Tesla wird in den USA zur Norm SAE J3400 {nacs}', 13),
}, """[8:15–9:05] Und heute?
Teslas Wechselstrom: Aus jeder Steckdose kommt Wechselstrom, in der Schweiz mit 230 Volt, in den USA mit 120 Volt {iec}. Die tiefe US-Spannung stammt aus der Anfangszeit mit Edisons rund 110 Volt.
Edisons Comeback: Akkus, Solarzellen und Elektronik arbeiten mit Gleichstrom. Für sehr lange Leitungen nutzt man die Hochspannungs-Gleichstrom-Übertragung (HGÜ); die erste kommerzielle Verbindung moderner Bauart ging 1954 in Schweden in Betrieb {gotland}.
Tesla als Name: Seit 1960 heisst die Einheit der magnetischen Flussdichte Tesla; ein MRI-Gerät im Spital arbeitet mit 0,35 bis 3 Tesla. Der Autohersteller Tesla wurde 2003 gegründet und nach ihm benannt {kaick}.
Neuer Standardkrieg: Der Ladestecker dieser Firma wird in den USA seit Ende 2023 zur Norm SAE J3400 {nacs}. So kämpft der Name Tesla wieder um einen Standard.
Überleitung: «Zurück zur Leitfrage.»""")
vier_felder(s, [('AC', 18), ('DC', 18), ('T', 22), ('vs.', 18)])

# ---------- 14 Fazit: beide Teile der Leitfrage ----------
s = S[I_FAZIT]
zweispalter(s, 'Fazit: Edison gegen Tesla', 'Warum Teslas System?',
            ['Es war billiger: hohe Spannung, wenige grosse Kraftwerke, weite Strecken.',
             'Nicht der berühmtere Erfinder gewann, sondern das System mit dem stärkeren Geldgeber.'],
            'Was zeigt das über die USA?',
            ['Billiger Strom und Grosskonzerne trugen den Aufstieg zur grössten Industriemacht.',
             'Standards werden früh gesetzt und bleiben lange: Pfadabhängigkeit {david92}.'],
            """[9:05–10:00] Fazit
Zurück zur Leitfrage, erster Teil: Teslas Wechselstrom gewann, weil er billiger war und weil Westinghouse ihn auf den Markt brachte; 1892 setzte auch General Electric darauf. Die Angstkampagne hat daran nichts geändert.
Zweiter Teil: Billiger Strom und Grosskonzerne wie General Electric und Westinghouse trugen den Aufstieg der USA zur grössten Industriemacht.
Schluss: «Edison hatte den Namen, Tesla die Technik. Gewonnen hat das System mit dem Geld dahinter. Darum kommt bis heute Teslas Wechselstrom aus unserer Steckdose.»""",
            l_size=16, r_size=16, label_size=14)
set_timing(s, Timeline().auto().add('float', 2342, dur=800).add('float', 2341, at=700, dur=800))

# ---------- Diskussionsblock (nur Fassung «mit Diskussion»): drei Fragen ans Plenum ----------
s = S[I_KAPITEL]
fill(s, {1874: '?', 1873: 'Diskussion', 1875: 'Drei Fragen ans Plenum'},
     """[ab 10:00] Diskussion – nur, wenn Zeit bleibt. Dauer vorher mit dem Dozenten absprechen.
Pro Frage zuerst zwei, drei Stimmen sammeln, dann die möglichen Antworten Klick für Klick einblenden.""")
set_timing(s, Timeline().auto().add('grow_turn', 1874, dur=800).add('fade', 1873, at=600, dur=600)
           .add('float', 1875, at=1000, dur=600))
qa_slide(S[I_G1], 'Frage 1: Angst als Werbung', 'Frage',
         ['Harold P. Brown tötete öffentlich Tiere, um Wechselstrom gefährlich wirken zu lassen. Wo gibt es heute '
          'Werbung oder Kampagnen, die mit Angst arbeiten?'],
         'Mögliche Antworten',
         ['Beispiele: Werbung, die ein Konkurrenzprodukt schlechtmacht; Angstkampagnen im Wahlkampf.',
          'Angst wirkt kurzfristig, setzt sich aber selten gegen klaren Nutzen und tiefere Kosten durch.',
          'So auch im Stromkrieg: Westinghouse entkräftete die Vorwürfe mit besserer Technik {cole}.'],
         'Zuerst Stimmen aus dem Plenum sammeln. Danach Klick für Klick ergänzen.')
qa_slide(S[I_G2], 'Frage 2: Tesla als Kultfigur', 'Frage',
         ['Lange galt Edison als der grosse Erfinder. Heute gilt Tesla oft als verkanntes Genie. Woher kommt '
          'dieses Bild?'],
         'Mögliche Antworten',
         ['Die Geschichte vom armen Genie gegen den reichen Unternehmer bleibt hängen.',
          'Filme, Internet und der Autohersteller Tesla machen den Namen bekannt {kaick}.',
          'Die Quellen zeigen ein nüchterneres Bild, etwa bei den 50’000 Dollar {tesla19}. Lehre: Primärquellen prüfen.'],
         'Zuerst Stimmen aus dem Plenum sammeln. Danach Klick für Klick ergänzen.')
qa_slide(S[I_G3], 'Frage 3: Standardkriege heute', 'Frage',
         ['Welche Standardkriege kennt ihr heute (Ladestecker, Streaming, Betriebssysteme)? Wer gewinnt meistens: '
          'die bessere Technik oder die stärkere Firma?'],
         'Mögliche Antworten',
         ['Beispiele: VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker für E-Autos {nacs}.',
          'Oft gewinnt, wer mehr Partner, tiefere Kosten und mehr Kapital hat.',
          'Ein gesetzter Standard bleibt lange: Pfadabhängigkeit {david92}. Beim Stromkrieg sprachen Technik und '
          'Kosten für den Wechselstrom.'],
         'Zuerst Stimmen aus dem Plenum sammeln. Danach Klick für Klick ergänzen.')

# ---------- Danke ----------
s = S[I_DANKE]
fill(s, {4515: 'Danke!', 4516: ['Fragen?', 'Simon Bruder · Modul «Weltmacht USA» (FHNW)', 'Dozent: Stephan Schwarz · 10.12.2026']},
     'Danke – Fragen? Quellen und Bildnachweis stehen auf den nächsten vier Folien; nur bei Nachfragen zeigen.')
remove_ids(s, (4845, 4848, 4851, 4854, 4855, 4860))  # Social-Media-Symbole der Vorlage
set_timing(s, Timeline().auto().add('zoom', 4515, dur=800).add('float', 4516, at=600, dur=700))

# ---------- Quellen: vollständige Liste, Nummern wie im Handout; am Schluss der Bildnachweis ----------
eintraege = [f'[{nr}] {text}' for nr, text in Z.eintraege()]
QUELLENFOLIEN = (S[I_Q1], S[I_Q2], S[I_Q3], F_Q4)
assert len(eintraege) <= 6 * 3 + 3, 'Zu viele Quellen: eine weitere Quellenfolie anlegen.'
BILDNACHWEIS = [
    'Bildnachweis (alle Bilder aus {natgeo})',
    'Frontfolie: W. A. Rogers, Stich für «Harper’s Weekly», 1889; De Agostini/Getty Images. Abdunkeln und '
    'Aufleuchten im Video: eigene Bearbeitung.',
    'Auburn 1890: kolorierte Illustration; Granger/Cordon Press.',
    'Karikatur: Currier & Ives, spätes 19. Jh.; MPI/Getty Images.',
    'Grafik zur Weltindustrie: eigene Darstellung nach {bairoch}.',
]
for k, s in enumerate(QUELLENFOLIEN):
    letzte = k == len(QUELLENFOLIEN) - 1
    teil = eintraege[6 * k:6 * k + 6]
    fill(s, {2340: f'Quellen ({k + 1}/{len(QUELLENFOLIEN)})' + (' und Bildnachweis' if letzte else '')},
         'Anhang – nur bei Nachfragen zeigen. Vollständige Liste im IEEE-Format; die Nummern sind dieselben wie im Handout.')
    sh = shapes_by_id(s)
    links, rechts = (teil[:3], teil[3:] + BILDNACHWEIS) if letzte else (teil[:3], teil[3:])
    for sid, x, zeilen in ((2342, 1.25, links), (2341, 5.28, rechts)):
        set_paras(sh[sid].text_frame, zeilen or [''], 10, rich=True)
        place(sh[sid], x, 1.64, 3.47, 3.2, anchor='t')
        space_after(sh[sid].text_frame, 7)
    textbox(s, 'Hinweis', 1.25, 1.24, 7.5, 0.3, [('IEEE-Format · Nummern wie im Handout', SANS, 11, False, True)])

# ---------- Übergänge ----------
for i, slide in enumerate(S):
    if i not in (I_TITEL, I_V2):
        set_transition(slide, 'fade')
set_transition(S[I_TITEL], 'morph', 1200)  # Frontfolie → Titel
set_transition(S[I_V2], 'morph', 1000)     # Zeitstrahl Teil 1 → Teil 2

# ---------- Aufräumen ----------
for slide in S:  # Die Vorlage nennt an einzelnen Stellen die ältere Schrift «Source Sans Pro».
    for el in slide._element.iter(A + 'latin', A + 'ea', A + 'cs', A + 'sym'):
        if el.get('typeface') == 'Source Sans Pro':
            el.set('typeface', SANS)
prs.core_properties.title = 'Der Stromkrieg – Edison vs. Tesla und die Rolle von George Westinghouse'
prs.core_properties.author = 'Simon Bruder'

HAUPTTEIL = [S[i] for i in (I_INTRO, I_TITEL, I_AUBURN, I_LEITFRAGE, I_FIRMEN)] + [F_ZITATE] + \
            [S[i] for i in (I_SPANNUNG, I_V1, I_V2, I_WARUM, I_CHART, I_ZAHLEN)] + [F_PRESSE] + \
            [S[i] for i in (I_MYTHEN, I_HEUTE, I_FAZIT)]
DISKUSSION = [S[i] for i in (I_KAPITEL, I_G1, I_G2, I_G3)]
SCHLUSS = [S[I_DANKE], *QUELLENFOLIEN]
ordne(prs, HAUPTTEIL + DISKUSSION + SCHLUSS)
prs.save(OUT_MIT)
print('OK', OUT_MIT, len(prs.slides), 'Folien')
prs = Presentation(OUT_MIT)
drop_slides(prs, range(len(HAUPTTEIL), len(HAUPTTEIL) + len(DISKUSSION)))
prs.save(OUT_OHNE)
print('OK', OUT_OHNE, len(prs.slides), 'Folien')
