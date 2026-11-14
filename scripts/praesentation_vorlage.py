"""Füllt die Folienstruktur mit dem Referat und erzeugt zwei Fassungen: mit und ohne Diskussionsblock.

Eingabe:  assets/folien_struktur.pptx (25 Folien aus der Slidesgo-Vorlage «Reconstruction Era and the Gilded Age»),
          die Bilder aus latex/handout/bilder und das Titelvideo (scripts/titelvideo.py)
Ausgabe:  abgabe/Praesentation_Stromkrieg_Bruder.pptx und …_mit_Diskussion.pptx
Aufruf:   python3 scripts/praesentation_vorlage.py
Sprechernotizen: die Karteikarten aus scripts/karten.py.

Quellen: Markup {schluessel} wird zu [n]. Die Nummern folgen der ersten Nennung im Handout
(latex/handout/main.tex). Eine Quelle, die im Handout fehlt, bricht den Lauf ab.
"""
import copy
import json
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
from karten import notiz

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
PORTRAETS = ROOT / 'assets' / 'portraets'  # edison.jpg, tesla.jpg und nachweis.json (gemeinfreie Fotos)
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


def portraet(name):
    """Porträt aus assets/portraets, falls vorhanden; sonst None (dann steht dort ein Buchstabe)."""
    datei = PORTRAETS / f'{name}.jpg'
    return datei if datei.exists() else None


def bildfuellung(shp, datei, fokus=0.25):
    """Füllt eine Form (Oval, Kreis) mit einem Foto, zugeschnitten auf das Seitenverhältnis der Form.
    Ist das Foto zu hoch, fällt oben nur der Anteil «fokus» des Überstands weg, damit das Gesicht bleibt."""
    from PIL import Image
    with Image.open(datei) as im:
        bw, bh = im.size
    _, rid = shp.part.get_or_add_image_part(str(datei))
    ziel, ist = shp.width / shp.height, bw / bh
    if ist > ziel:  # zu breit: links und rechts gleich viel weg
        rand = {'l': (1 - ziel / ist) / 2, 'r': (1 - ziel / ist) / 2}
    else:           # zu hoch: oben wenig, unten mehr weg
        rest = 1 - ist / ziel
        rand = {'t': rest * fokus, 'b': rest * (1 - fokus)}
    sppr = shp._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill'):
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)
    bf = etree.Element(qn('a:blipFill'), rotWithShape='1')
    etree.SubElement(bf, qn('a:blip')).set(qn('r:embed'), rid)
    src = etree.SubElement(bf, qn('a:srcRect'))
    for k, v in rand.items():
        src.set(k, str(round(v * 100000)))
    etree.SubElement(etree.SubElement(bf, qn('a:stretch')), qn('a:fillRect'))
    sppr.find(qn('a:prstGeom')).addnext(bf)


def rund_portraet(slide, name, x, y, foto, d=0.66):
    """Rundes Porträt mitten im Rahmenfeld der Vorlage (0,81 × 0,68 Zoll) statt des Buchstabens."""
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + (0.81 - d) / 2), Inches(y + (0.68 - d) / 2),
                                 Inches(d), Inches(d))
    rename(shp, name)
    shp.shadow.inherit = False
    shp.line.color.rgb = NAVY
    shp.line.width = Pt(1.25)
    bildfuellung(shp, foto, fokus=0.1)  # im Kreis mehr Platz über dem Kopf
    return shp


def medallion(slide, name, x, y, w, h, letter, bg=None, letter_color=NAVY, size=40, foto=None):
    """Ovaler Bilderrahmen mit doppelter Linie: mit Foto, falls eines vorliegt, sonst mit Buchstabe."""
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
    if foto:
        bildfuellung(shp, foto)
        return shp
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
# Zusätzliche Folien als Kopien von Vorlagenfolien (vor dem Füllen anlegen)
F_ARGUMENTE = kopie(prs, S[I_Q2])   # Zwei Argumente (zwei Spalten)
F_EUROPA = kopie(prs, S[I_WARUM])   # USA und Europa (drei Spalten)
F_PRESSE = kopie(prs, S[I_Q1])      # Karikatur
F_Q4 = kopie(prs, S[I_Q3])          # vierte Quellenfolie
F_BILD = kopie(prs, S[I_Q3])        # Bildnachweis


def K(key):
    """Karteikarte als Sprechernotiz (scripts/karten.py)."""
    return notiz(key, T)


# ---------- 1 Frontfolie: Bild von 1889 als kurzes Video, darüber der Titel ----------
s = S[I_INTRO]
remove_ids(s, (896,))  # Ornamente der Vorlage; das Bild füllt den inneren Rahmen
fill(s, {890: ('Der Stromkrieg', 40), 891: ('Edison vs. Tesla – und die Rolle von George Westinghouse', 16)}, K('front'))
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
         891: ('Simon Bruder · Dozent: Stephan Schwarz · 10.12.2026', 13)}, K('titel'))
sh = shapes_by_id(s)
place(sh[895], 3.6, 1.11, 2.8)
for slide in (S[I_INTRO], S[I_TITEL]):
    t = shapes_by_id(slide)
    rename(t[890], '!!Titel')
    rename(t[891], '!!Untertitel')
edison = medallion(s, '!!Edison', 1.42, 1.62, 1.05, 1.35, 'E', size=36, foto=portraet('edison'))
edison_name = textbox(s, 'Edison Name', 0.995, 3.02, 1.9, 0.5,
                      [('Thomas A. Edison', SERIF, 11, True, False), ('Gleichstrom', SANS, 11, False, True)])
tesla = medallion(s, '!!Tesla', 7.53, 1.62, 1.05, 1.35, 'T', size=36, foto=portraet('tesla'))
tesla_name = textbox(s, 'Tesla Name', 6.955, 3.02, 2.2, 0.5,
                     [('Nikola Tesla', SERIF, 11, True, False), ('Wechselstrom · mit Westinghouse', SANS, 11, False, True)])
westinghouse = medallion(s, '!!Westinghouse', 8.24, 1.48, 0.55, 0.55, 'W', bg=NAVY2, letter_color=WEISS, size=20)
set_timing(s, Timeline().auto()
           .add('fly_left', edison.shape_id, at=0, dur=900).add('fade', edison_name.shape_id, at=500, dur=500)
           .add('fly_right', tesla.shape_id, at=1000, dur=900).add('fade', tesla_name.shape_id, at=1500, dur=500)
           .add('grow_turn', westinghouse.shape_id, at=2100, dur=800))

# ---------- 3 Einstieg: Auburn 1890 ----------
s = S[I_AUBURN]
remove_ids(s, (1407,))  # Ornamente der Vorlage; links steht das Bild
fill(s, {1864: ('- Auburn, 1890 -', 48),
         1865: ('6. August: erste Hinrichtung auf dem elektrischen Stuhl – mit einem Generator von Westinghouse {essig}', 14)},
     K('auburn'))
sh = shapes_by_id(s)
stuhl = bild(s, 'Bild Stuhl', 'elektrischer_stuhl_1890.jpg', 1.0, 0.95, h=3.6)
nachweis(s, 'Bildnachweis', 'Bild: Granger/Cordon Press {natgeo}', 1.0, 4.58, 2.6)
place(sh[1864], 4.0, 1.5, 5.0)
place(sh[1865], 4.0, 3.08, 5.0)
sh[1866].left, sh[1866].width = Inches(3.86), Inches(5.28)
set_timing(s, Timeline().auto().add('fade', stuhl.shape_id, dur=800).add('zoom', 1864, at=400, dur=900)
           .add('float', 1865, at=1200, dur=700))

# ---------- 4 Leitfrage und These (kurz; die Langfassung steht im Handout) ----------
s = S[I_LEITFRAGE]
zweispalter(s, 'Leitfrage und These', 'Leitfrage',
            ['Warum setzte sich Teslas Wechselstrom gegen Edisons Gleichstrom durch – und was zeigt das über den '
             'Aufstieg der USA?'],
            'These',
            ['Nicht der berühmtere Erfinder gewann, sondern das billigere System mit dem stärkeren Geldgeber.'],
            K('leitfrage'), l_size=17, r_size=17)
set_timing(s, Timeline().auto().add('fade', 2342, dur=700).click().add('float', 2341, dur=700, para=0))


# ---------- 5 Zwei Erfinder und ihre Geldgeber (vier Felder) ----------
def vier_felder(slide, zeichen):
    """Verbreitert die Textfelder des Vier-Felder-Layouts und setzt Zeichen oder Porträts in die Rahmenfelder."""
    sh = shapes_by_id(slide)
    for sid in (2434, 2430, 2435, 2432):  # linke Spalte
        place(sh[sid], 2.31, sh[sid].top / 914400, 2.7)
    for sid in (2436, 2431, 2437, 2433):  # rechte Spalte (rechtsbündig)
        place(sh[sid], 4.99, sh[sid].top / 914400, 2.7)
    felder = [(1.33, 1.86, 2434, 2430), (7.86, 1.86, 2436, 2431), (1.33, 3.44, 2435, 2432), (7.86, 3.44, 2437, 2433)]
    spalten = []
    for (x, y, titel, text), (z, size, *foto) in zip(felder, zeichen):
        if foto and foto[0]:
            form = rund_portraet(slide, f'Porträt {z}', x, y, foto[0])
        else:
            form = slot(slide, f'Zeichen {z}', x, y, z, size)
        spalten.append((form.shape_id, titel, text))
    staggered(slide, spalten)


s = S[I_FIRMEN]
remove_ids(s, (2450, 2455, 2458, 2467))  # Symbolbilder der Vorlage
fill(s, {
    2429: 'Zwei Erfinder und ihre Geldgeber',
    2434: ('Thomas A. Edison', 15), 2430: ('Erfinder und Unternehmer, setzt auf Gleichstrom {jonnes}', 13),
    2436: ('Nikola Tesla', 15), 2431: ('1884 sechs Monate bei Edison, dann eigener Weg mit Wechselstrom {kaick}', 13),
    2435: ('J. P. Morgan', 15), 2432: ('Bankier, finanziert Edison {jonnes}', 13),
    2437: ('George Westinghouse', 15), 2433: ('Unternehmer, kauft 1888 Teslas Patente {kaick}', 13),
}, K('erfinder'))
vier_felder(s, [('E', 22, portraet('edison')), ('T', 22, portraet('tesla')), ('M', 22), ('W', 22)])

# ---------- 6 Zwei Systeme: Stärken und Schwächen (Tabelle) ----------
s = S[I_SPANNUNG]
fill(s, {3480: 'Zwei Systeme: Stärken und Schwächen'}, K('systeme'))
frame = shapes_by_id(s)[3481]
tbl = frame.table
for tr in tbl._tbl.tr_lst[5:]:  # Kopfzeile und vier Zeilen bleiben
    tbl._tbl.remove(tr)
fill_table(tbl, [
    ['', 'Gleichstrom (DC) – Edison', 'Wechselstrom (AC) – Tesla'],
    ['Umspannen', 'damals nicht möglich (rund 110 V)', 'Transformator: hoch im Netz, tief im Haus'],
    ['Reichweite', 'ein Quartier (Pearl Street) {pearl}', 'weit: Niagara–Buffalo, rund 40 km {jonnes}'],
    ['Sicherheit', 'niedrige Spannung, gilt als sicher', 'Hochspannung: tödliche Unfälle 1889 {essig}'],
    ['Motor, Speicher', 'erprobte Motoren, Akkus', 'Teslas Drehfeldmotor erst ab 1888 {kaick}'],
], size=13)
for col, breite in zip(tbl.columns, (1.57, 3.43, 3.43)):  # jede Zelle bleibt einzeilig
    col.width = Inches(breite)
for k, row in enumerate(tbl.rows):
    row.height = Inches(0.46 if k == 0 else 0.56)
frame.height = Inches(0.46 + 4 * 0.56)
merksatz = textbox(s, 'Merksatz', 0.79, 4.4, 8.43, 0.45,
                   [('Merksatz: 10-fache Spannung → 1/10 des Stroms → 1/100 des Verlusts (I² · R)', SANS, 14, True, False)])
set_timing(s, Timeline().auto().add('float', merksatz.shape_id, at=600, dur=700))

# ---------- 7 Zwei Argumente: Sicherheit gegen Kosten ----------
s = F_ARGUMENTE
zweispalter(s, 'Sicherheit gegen Kosten', 'Edison: Sicherheit',
            ['1889: Nichts rechtfertige Hochspannung und Wechselstrom {edison89}.',
             'Anlass: tödliche Unfälle an Leitungen über New York {essig}.'],
            'Westinghouse und Tesla: Kosten',
            ['Hohe Spannung spart Kupfer und Kraftwerke.',
             'Strom reist verlustarm über weite Strecken, auch von Wasserkraftwerken wie Niagara {kaick}.'],
            K('argumente'), l_size=16, r_size=16, label_size=13)
textbox(s, 'Hinweis', 1.25, 4.62, 7.5, 0.25, [('Edison sinngemäss übersetzt', SANS, 9, False, True)])
set_timing(s, Timeline().auto().add('fade', 2342, dur=700)
           .click().add('float', 2341, dur=700, para=0).add('float', 2341, at=600, dur=700, para=1))


# ---------- 8 und 9 Verlauf (Zeitstrahl; Jahreszahlen stehen in den Rahmenfeldern) ----------
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
], K('verlauf1'))
staggered(S[I_V1], spalten, step=350)
zeitstrahl(S[I_V2], '1891–1896: Teslas System setzt sich durch', [
    ('1891', 'Lizenzverzicht', 'Westinghouse in Finanznot; Tesla verzichtet auf seine Gebühren {carlson}'),
    ('1892', 'General Electric', 'Edisons Firma geht in GE auf; auch GE setzt auf Wechselstrom {jonnes}'),
    ('1893', 'Chicago', 'Teslas System versorgt die Weltausstellung {kaick}'),
    ('1896', 'Niagara–Buffalo', 'Teslas System bringt Strom über rund 40 km {kaick,jonnes}'),
], K('verlauf2'))


# ---------- 10 Warum Teslas System gewann, 11 USA und Europa (drei Spalten) ----------
def drei_spalten(slide, title, spalten, notes, zeichen=('1', '2', '3'), groesse=24, text_size=14):
    """Drei Spalten mit Zeichen (Ziffern oder Kürzel) in den Rahmenfeldern."""
    remove_ids(slide, (2394, 2399, 2400))  # Symbolbilder der Vorlage
    mapping = {2378: title}
    for (titel, text), (t, d) in zip(((2382, 2379), (2383, 2380), (2384, 2381)), spalten):
        mapping.update({titel: (t, 15), text: (d, text_size)})
    fill(slide, mapping, notes)
    cols = []
    for k, (x, titel, text) in enumerate(((1.81, 2382, 2379), (4.59, 2383, 2380), (7.38, 2384, 2381))):
        cols.append((slot(slide, f'Zeichen {zeichen[k]}', x, 2.0, zeichen[k], groesse).shape_id, titel, text))
    staggered(slide, cols, step=450)


drei_spalten(S[I_WARUM], 'Warum Teslas System gewann', [
    ('Kosten', 'Hohe Spannung, wenig Verlust: wenige grosse Kraftwerke statt eines pro Quartier {jonnes}'),
    ('Kapital', 'Westinghouse bringt Teslas System auf den Markt; 1892 geht Edisons Firma in GE auf {kaick,jonnes}'),
    ('Kampagne verpufft', 'Westinghouse ignoriert die Kampagne, bestreitet sie und entkräftet sie mit Technik {cole}'),
], K('warum'))
drei_spalten(F_EUROPA, 'USA und Europa', [
    ('Gilded Age', 'Kaum Regeln, Bankiers wie Morgan, Patentprozesse vor Gericht {jonnes}'),
    ('Pioniere in Europa', 'Transformator: Gaulard und Gibbs; Drehfeld: Ferraris; 1891 Lauffen–Frankfurt (MFO) '
                           '{jonnes,carlson,steig}'),
    ('Zwei Normen', 'USA 120 V / 60 Hz, Schweiz 230 V / 50 Hz {iec}'),
], K('europa'), zeichen=('USA', 'EU', 'Hz'), groesse=18, text_size=13)

# ---------- 12 Grafik Weltindustrie ----------
s = S[I_CHART]
remove_ids(s, (4458, 4459, 4460, 4461, 4462, 4463))  # Legende der Vorlage; die Grafik beschriftet die Linien direkt
fill(s, {4456: 'Die USA überholen Grossbritannien',
         4457: ('Anteil an der Weltindustrieproduktion in %. Linien verbinden Stichjahre (Schätzwerte). '
                'Eigene Darstellung nach {bairoch}.', 11)}, K('grafik'))
sh = shapes_by_id(s)
pic = sh[4464]
s.part.related_part(pic._element.blipFill.find(A + 'blip').get(R_EMBED))._blob = \
    (ROOT / 'assets' / 'chart_weltindustrie_folie.png').read_bytes()
place(pic, 1.2, 1.36, 7.6, 3.0)
place(sh[4457], 0.79, 4.4, 8.43, 0.36)
set_timing(s, Timeline().auto().add('wipe_left', 4464, dur=2000).add('fade', 4457, at=1800, dur=600))

# ---------- 13 Bildquelle: Karikatur von Currier & Ives ----------
s = F_PRESSE
remove_ids(s, (2342,))
fill(s, {2340: 'Vor Tesla: Edison gegen Brush',
         2341: (['1880, Jahre vor Tesla: Edison (links) gegen Charles Brush, den Erfinder der Bogenlampe {natgeo}.',
                 '1892 landen beide Firmen in General Electric {ech}.',
                 'Erzählt wird ein Duell, heute Edison gegen Tesla. Entschieden haben Firmen und Kapital.'],
                15)}, K('karikatur'))
sh = shapes_by_id(s)
place(sh[2341], 5.45, 1.95, 3.55, 2.6, anchor='t')
space_after(sh[2341].text_frame, 12)
karikatur = bild(s, 'Bild Karikatur', 'karikatur_edison_brush.jpg', 1.0, 1.72, w=4.1)
nachweis(s, 'Bildnachweis', 'Bild: Currier & Ives, «The Electric Light», 1880 {met} · MPI/Getty Images {natgeo}',
         1.0, 4.81, 4.1)
set_timing(s, Timeline().auto().add('fade', karikatur.shape_id, dur=700).add('float', 2341, at=500, dur=700, para=0)
           .click().add('float', 2341, dur=700, para=1).click().add('float', 2341, dur=700, para=2))

# ---------- 14 Mythen (drei Tabellen: Mythos / Befund) ----------
s = S[I_MYTHEN]
fill(s, {4469: 'Drei Mythen – und was die Quellen sagen'}, K('mythen'))
sh = shapes_by_id(s)
for sid, y, rows in (
    (4470, 1.45, [['Mythos: Edisons «Experimente mit Hunden, Pferden und sogar einem Elefanten» {abb}'],
                  ['Befund: Die Tiere tötete Harold P. Brown; Topsy starb erst 1903, ohne Edison {hughes58,topsy}']]),
    (4471, 2.6, [['Mythos: Edison prellte Tesla um 50’000 Dollar'],
                 ['Befund: Teslas eigener Bericht (1919) nennt «den Manager», nicht Edison {tesla19}']]),
    (4472, 3.75, [['Mythos: Tesla erfand den Wechselstrom allein'],
                  ['Befund: Den Transformator gab es schon in Europa; das Drehfeld fand Ferraris gleichzeitig {jonnes,carlson}']]),
):
    fill_table(sh[sid].table, rows, size=[12, 13])
    for row in sh[sid].table.rows:
        row.height = Inches(0.42)
    sh[sid].top, sh[sid].height = Inches(y), Inches(0.84)
set_timing(s, Timeline().auto().add('fade', 4470, dur=600).add('fade', 4471, at=500, dur=600).add('fade', 4472, at=1000, dur=600))

# ---------- 15 Und heute? ----------
s = S[I_HEUTE]
remove_ids(s, (2450, 2455, 2458, 2467))
fill(s, {
    2429: 'Und heute?',
    2434: ('Teslas Wechselstrom', 14), 2430: ('aus jeder Steckdose: Schweiz 230 V, USA 120 V {iec}', 13),
    2436: ('Gleichstrom zurück', 14), 2431: ('Akkus, E-Autos, HGÜ-Leitungen; SuedLink ab 2028 {gotland,bnetza}', 13),
    2435: ('Tesla als Name', 14), 2432: ('Einheit seit 1960, Autohersteller, Ladestecker-Norm {kaick,nacs}', 13),
    2437: ('Edison bis 2007', 14), 2433: ('Con Edison schaltet in Manhattan den letzten Gleichstrom ab {lee}', 13),
}, K('heute'))
vier_felder(s, [('AC', 18), ('DC', 18), ('T', 22), ('NY', 18)])

# ---------- 16 Fazit: beide Teile der Leitfrage ----------
s = S[I_FAZIT]
zweispalter(s, 'Fazit: Edison gegen Tesla', 'Warum Teslas System?',
            ['Es war billiger: weite Strecken, wenige grosse Kraftwerke.',
             'Edisons Sorge war echt; entschieden haben Kosten und Kapital.'],
            'Was zeigt das über die USA?',
            ['Billiger Strom und Grosskonzerne trugen den Aufstieg zur Industriemacht.',
             'Standards bleiben lange: Pfadabhängigkeit {david92}.'],
            K('fazit'), l_size=16, r_size=16, label_size=14)
set_timing(s, Timeline().auto().add('float', 2342, dur=800).add('float', 2341, at=700, dur=800))

# ---------- Diskussionsblock (nur Fassung «mit Diskussion»): drei Fragen ans Plenum ----------
s = S[I_KAPITEL]
fill(s, {1874: '?', 1873: 'Diskussion', 1875: 'Drei Fragen ans Plenum'}, K('diskussion'))
set_timing(s, Timeline().auto().add('grow_turn', 1874, dur=800).add('fade', 1873, at=600, dur=600)
           .add('float', 1875, at=1000, dur=600))
qa_slide(S[I_G1], 'Frage 1: Angst als Werbung', 'Frage',
         ['Harold P. Brown tötete öffentlich Tiere, um Wechselstrom gefährlich wirken zu lassen. Wo gibt es heute '
          'Werbung oder Kampagnen, die mit Angst arbeiten?'],
         'Mögliche Antworten',
         ['Beispiele: Werbung, die ein Konkurrenzprodukt schlechtmacht; Angstkampagnen im Wahlkampf.',
          'Angst wirkt kurzfristig, setzt sich aber selten gegen klaren Nutzen und tiefere Kosten durch.',
          'So auch im Stromkrieg: Westinghouse entkräftete die Vorwürfe mit besserer Technik {cole}.'],
         K('frage1'))
qa_slide(S[I_G2], 'Frage 2: Tesla als Kultfigur', 'Frage',
         ['Lange galt Edison als der grosse Erfinder. Heute gilt Tesla oft als verkanntes Genie. Woher kommt '
          'dieses Bild?'],
         'Mögliche Antworten',
         ['Die Geschichte vom armen Genie gegen den reichen Unternehmer bleibt hängen.',
          'Filme, Internet und der Autohersteller Tesla machen den Namen bekannt {kaick}.',
          'Die Quellen zeigen ein nüchterneres Bild, etwa bei den 50’000 Dollar {tesla19}. Lehre: Primärquellen prüfen.'],
         K('frage2'))
qa_slide(S[I_G3], 'Frage 3: Standardkriege heute', 'Frage',
         ['Welche Standardkriege kennt ihr heute (Ladestecker, Streaming, Betriebssysteme)? Wer gewinnt meistens: '
          'die bessere Technik oder die stärkere Firma?'],
         'Mögliche Antworten',
         ['Beispiele: VHS gegen Betamax, Blu-ray gegen HD DVD, Ladestecker für E-Autos {nacs}.',
          'Oft gewinnt, wer mehr Partner, tiefere Kosten und mehr Kapital hat.',
          'Ein gesetzter Standard bleibt lange: Pfadabhängigkeit {david92}. Beim Stromkrieg sprachen Technik und '
          'Kosten für den Wechselstrom.'],
         K('frage3'))

# ---------- Danke ----------
s = S[I_DANKE]
fill(s, {4515: 'Danke!', 4516: ['Fragen?', 'Simon Bruder · Modul «Weltmacht USA» (FHNW)', 'Dozent: Stephan Schwarz · 10.12.2026']},
     K('danke'))
remove_ids(s, (4845, 4848, 4851, 4854, 4855, 4860))  # Social-Media-Symbole der Vorlage
set_timing(s, Timeline().auto().add('zoom', 4515, dur=800).add('float', 4516, at=600, dur=700))

# ---------- Quellen: vollständige Liste, Nummern wie im Handout; danach der Bildnachweis ----------
eintraege = [f'[{nr}] {text}' for nr, text in Z.eintraege()]
QUELLENFOLIEN = (S[I_Q1], S[I_Q2], S[I_Q3], F_Q4)
PRO_FOLIE = -(-len(eintraege) // len(QUELLENFOLIEN))  # 26 Quellen: 7 pro Folie
assert PRO_FOLIE <= 7, 'Zu viele Quellen: eine weitere Quellenfolie anlegen.'


def spalten_teilen(teil):
    """Teilt die Einträge einer Folie so auf zwei Spalten auf, dass beide etwa gleich viel Text haben."""
    if len(teil) < 2:
        return len(teil)
    return min(range(1, len(teil)),
               key=lambda i: max(sum(map(len, teil[:i])), sum(map(len, teil[i:]))))


for k, s in enumerate(QUELLENFOLIEN):
    teil = eintraege[PRO_FOLIE * k:PRO_FOLIE * (k + 1)]
    links = spalten_teilen(teil)
    fill(s, {2340: f'Quellen ({k + 1}/{len(QUELLENFOLIEN)})'},
         'Anhang – nur bei Nachfragen zeigen. Vollständige Liste im IEEE-Format; die Nummern sind dieselben wie im Handout.')
    sh = shapes_by_id(s)
    for sid, x, zeilen in ((2342, 1.25, teil[:links]), (2341, 5.28, teil[links:])):
        set_paras(sh[sid].text_frame, zeilen or [''], 9.5 if PRO_FOLIE <= 6 else 9, rich=True)
        place(sh[sid], x, 1.6, 3.47, 3.3, anchor='t')
        space_after(sh[sid].text_frame, 5)
    textbox(s, 'Hinweis', 1.25, 1.24, 7.5, 0.3, [('IEEE-Format · Nummern wie im Handout', SANS, 11, False, True)])
s = F_BILD
PORTRAET_NACHWEIS = []
if portraet('edison') or portraet('tesla'):  # Fotos nur mit Nachweis (Fotograf, Jahr, Lizenz, Fundort)
    nachweis_datei = PORTRAETS / 'nachweis.json'
    assert nachweis_datei.exists(), 'Porträts ohne assets/portraets/nachweis.json: Bildnachweis ergänzen.'
    PORTRAET_NACHWEIS = ['Porträts: ' + '; '.join(json.loads(nachweis_datei.read_text(encoding='utf-8')).values()) + '.']
fill(s, {2340: 'Bildnachweis'}, 'Anhang – nur bei Nachfragen zeigen.')
sh = shapes_by_id(s)
for sid, x, zeilen in (
    (2342, 1.25, ['Frontfolie: William Allen Rogers, Stich für «Harper’s Weekly», 1889; De Agostini/Getty Images, '
                  'aus {natgeo}. Abdunkeln und Aufleuchten im Video: eigene Bearbeitung.',
                  'Auburn 1890: kolorierte Illustration; Granger/Cordon Press, aus {natgeo}.']),
    (2341, 5.28, ['Karikatur: Currier & Ives, «The Electric Light», Lithografie, 1880 {met}; MPI/Getty Images, aus {natgeo}.',
                  'Grafik zur Weltindustrie: eigene Darstellung nach {bairoch}.'] + PORTRAET_NACHWEIS),
):
    set_paras(sh[sid].text_frame, zeilen, 12, rich=True)
    place(sh[sid], x, 1.64, 3.47, 3.2, anchor='t')
    space_after(sh[sid].text_frame, 10)
textbox(s, 'Hinweis', 1.25, 1.24, 7.5, 0.3, [('Bilder nur für den Unterricht; die Rechte liegen bei den Agenturen', SANS, 11, False, True)])

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

HAUPTTEIL = [S[I_INTRO], S[I_TITEL], S[I_AUBURN], S[I_LEITFRAGE], S[I_FIRMEN], S[I_SPANNUNG], F_ARGUMENTE,
             S[I_V1], S[I_V2], S[I_WARUM], F_EUROPA, S[I_CHART], F_PRESSE, S[I_MYTHEN], S[I_HEUTE], S[I_FAZIT]]
DISKUSSION = [S[i] for i in (I_KAPITEL, I_G1, I_G2, I_G3)]
SCHLUSS = [S[I_DANKE], *QUELLENFOLIEN, F_BILD]
ordne(prs, HAUPTTEIL + DISKUSSION + SCHLUSS)
prs.save(OUT_MIT)
print('OK', OUT_MIT, len(prs.slides), 'Folien')
prs = Presentation(OUT_MIT)
drop_slides(prs, range(len(HAUPTTEIL), len(HAUPTTEIL) + len(DISKUSSION)))
prs.save(OUT_OHNE)
print('OK', OUT_OHNE, len(prs.slides), 'Folien')
