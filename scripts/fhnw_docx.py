"""Bausteine für Word-Dokumente im Layout von Simons FHNW-Handout.

Kopf-/Fusszeile und Logo stammen aus einer bestehenden .docx-Vorlage; nur der Inhalt wird ersetzt.
"""
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

RPR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>{b}{i}<w:color w:val="000000"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:val="de-CH"/>'
PB = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


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


def items(lst, letters=True, sz=20):
    out = []
    for k, t in enumerate(lst):
        mark = f'{"abcdefgh"[k]})' if letters else '–'
        out.append(f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:spacing w:after="60" w:line="264" w:lineRule="auto"/>'
                   f'<w:ind w:left="720" w:hanging="360"/></w:pPr>'
                   f'{run(mark, sz=sz)}<w:r><w:tab/></w:r>{rich(t, sz)}</w:p>')
    return ''.join(out)


def table(rows, widths, head=None, first_center=True, sz=20):
    def cell(t, w, bold=False, center=False):
        jc = '<w:jc w:val="center"/>' if center else ''
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/>{jc}</w:pPr>'
                f'{run(t, b=True, sz=sz) if bold else rich(t, sz)}</w:p></w:tc>')
    border = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders>'
         f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % w for w in widths) + '</w:tblGrid>')
    if head:
        x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ''.join(cell(h, w, True, True) for h, w in zip(head, widths)) + '</w:tr>'
    for r in rows:
        x += ('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(r[0], widths[0], False, first_center)
              + ''.join(cell(t, w) for t, w in zip(r[1:], widths[1:])) + '</w:tr>')
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


def references(refs, sz=18):
    """Nummerierte IEEE-Liste [1], [2], … mit hängendem Einzug."""
    return ''.join(f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs><w:spacing w:after="60" w:line="240" w:lineRule="auto"/>'
                   f'<w:ind w:left="567" w:hanging="567"/></w:pPr>{run(f"[{k}]", sz=sz)}<w:r><w:tab/></w:r>{rich(r, sz)}</w:p>'
                   for k, r in enumerate(refs, 1))


def build(src, out, body, title, chart=None):
    """Ersetzt den Inhalt der Vorlage `src` durch `body` (Liste von XML-Stücken) und schreibt `out`.

    Bilder der Vorlage werden entfernt; `chart` (PNG) wird als Beziehung rIdChart eingebunden.
    """
    tmp = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(src) as z:
            z.extractall(tmp)
        for p in tmp.rglob('*'):
            if p.is_symlink():
                p.unlink()
        doc = (tmp / 'word/document.xml').read_text(encoding='utf-8')
        sect = re.search(r'<w:sectPr\b.*?</w:sectPr>', doc, re.S).group(0)
        doc = re.sub(r'<w:body>.*</w:body>', lambda _: '<w:body>' + ''.join(body) + sect + '</w:body>', doc, flags=re.S)
        (tmp / 'word/document.xml').write_text(doc, encoding='utf-8')

        rels_p = tmp / 'word/_rels/document.xml.rels'
        rels = rels_p.read_text(encoding='utf-8')
        for m in re.finditer(r'<Relationship [^>]*Type="[^"]*/image" Target="media/([^"]+)"/>', rels):
            (tmp / 'word/media' / m.group(1)).unlink(missing_ok=True)
        rels = re.sub(r'<Relationship [^>]*Type="[^"]*/image"[^>]*/>', '', rels)
        if chart:
            rels = rels.replace('</Relationships>', '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/chart.png"/></Relationships>')
            shutil.copy(chart, tmp / 'word/media/chart.png')
            ct = (tmp / '[Content_Types].xml').read_text(encoding='utf-8')
            if 'Extension="png"' not in ct:
                (tmp / '[Content_Types].xml').write_text(ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'), encoding='utf-8')
        rels_p.write_text(rels, encoding='utf-8')

        for f in (tmp / 'word').glob('footer*.xml'):  # Fusszeile: Dozent
            f.write_text(f.read_text(encoding='utf-8').replace('Stefan Czarnecki', 'Stephan Schwarz'), encoding='utf-8')
        core = tmp / 'docProps/core.xml'
        if core.exists():
            c = core.read_text(encoding='utf-8')
            core.write_text(re.sub(r'<dc:title>.*?</dc:title>', f'<dc:title>{escape(title)}</dc:title>', c), encoding='utf-8')

        Path(out).unlink(missing_ok=True)
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for p in sorted(tmp.rglob('*')):
                if p.is_file() and not p.relative_to(tmp).as_posix().startswith('[trash]'):
                    z.write(p, p.relative_to(tmp).as_posix())
    finally:
        shutil.rmtree(tmp)
    print('OK', out)
