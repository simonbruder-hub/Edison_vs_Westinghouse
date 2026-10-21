"""PowerPoint-Animationen (p:timing) und Folienübergänge (Morph, Verblassen) als XML für python-pptx.

Eine Folie bekommt eine Zeitleiste aus Klickgruppen. Jede Gruppe startet per Klick oder – nur die erste –
automatisch mit der Folie. Effekte in einer Gruppe haben einen Startversatz `at` in Millisekunden.
"""
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NSDECL = f'xmlns:p="{P}" xmlns:a="{A}"'

# presetID/-subtype wie in PowerPoint, damit der Animationsbereich die Effekte richtig benennt
PRESETS = {
    'fade': (10, 0), 'fly_left': (2, 8), 'fly_right': (2, 2), 'fly_bottom': (2, 4),
    'float': (42, 0), 'zoom': (53, 16), 'grow_turn': (31, 0), 'wipe_left': (22, 8),
}


class Timeline:
    def __init__(self):
        self.groups = []

    def auto(self):
        assert not self.groups, 'Nur die erste Gruppe kann automatisch starten.'
        self.groups.append({'trigger': 'auto', 'effects': []})
        return self

    def click(self):
        self.groups.append({'trigger': 'click', 'effects': []})
        return self

    def add(self, kind, spid, at=0, dur=600, para=None):
        self.groups[-1]['effects'].append(dict(kind=kind, spid=spid, at=at, dur=dur, para=para))
        return self


class _Ids:
    def __init__(self):
        self.n = 2

    def __call__(self):
        self.n += 1
        return self.n


def _tgt(spid, para):
    if para is None:
        return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    return f'<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para}" end="{para}"/></p:txEl></p:spTgt></p:tgtEl>'


def _set_visible(ids, tgt):
    return (f'<p:set><p:cBhvr><p:cTn id="{ids()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'{tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>')


def _fade(ids, tgt, dur):
    return f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>'


def _wipe(ids, tgt, dur, direction):
    return f'<p:animEffect transition="in" filter="wipe({direction})"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>'


def _val(v):
    return f'<p:fltVal val="{v}"/>' if isinstance(v, (int, float)) else f'<p:strVal val="{v}"/>'


def _anim(ids, tgt, dur, attr, frm, to):
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt}'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val>{_val(frm)}</p:val></p:tav><p:tav tm="100000"><p:val>{_val(to)}</p:val></p:tav></p:tavLst></p:anim>')


def _behaviours(ids, kind, tgt, dur):
    if kind == 'fade':
        return _fade(ids, tgt, dur)
    if kind.startswith('fly_'):
        x0, y0 = {'fly_left': ('0-#ppt_w/2', '#ppt_y'), 'fly_right': ('1+#ppt_w/2', '#ppt_y'),
                  'fly_bottom': ('#ppt_x', '1+#ppt_h/2')}[kind]
        return _anim(ids, tgt, dur, 'ppt_x', x0, '#ppt_x') + _anim(ids, tgt, dur, 'ppt_y', y0, '#ppt_y')
    if kind == 'float':
        return (_fade(ids, tgt, dur) + _anim(ids, tgt, dur, 'ppt_x', '#ppt_x', '#ppt_x')
                + _anim(ids, tgt, dur, 'ppt_y', '#ppt_y+.1', '#ppt_y'))
    if kind == 'zoom':
        return (_anim(ids, tgt, dur, 'ppt_w', 0, '#ppt_w') + _anim(ids, tgt, dur, 'ppt_h', 0, '#ppt_h')
                + _fade(ids, tgt, dur))
    if kind == 'grow_turn':
        return (_anim(ids, tgt, dur, 'ppt_w', 0, '#ppt_w') + _anim(ids, tgt, dur, 'ppt_h', 0, '#ppt_h')
                + _anim(ids, tgt, dur, 'style.rotation', 90, 0) + _fade(ids, tgt, dur))
    if kind == 'wipe_left':
        return _wipe(ids, tgt, dur, 'left')
    raise ValueError(kind)


def _effect(ids, e, node_type):
    pid, sub = PRESETS[e['kind']]
    tgt = _tgt(e['spid'], e['para'])
    smooth = ' decel="100000"' if e['kind'].startswith('fly_') or e['kind'] == 'float' else ''
    ctn = ids()
    body = _set_visible(ids, tgt) + _behaviours(ids, e['kind'], tgt, e['dur'])
    return (f'<p:par><p:cTn id="{ctn}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0"'
            f' nodeType="{node_type}"{smooth}><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')


def timing_xml(tl, slide):
    ids = _Ids()
    groups = []
    for g in tl.groups:
        gid = ids()
        cond = '<p:cond delay="indefinite"/>'
        if g['trigger'] == 'auto':
            cond += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
        inner = []
        for k, at in enumerate(sorted({e['at'] for e in g['effects']})):
            effs = [e for e in g['effects'] if e['at'] == at]
            pars = ''.join(_effect(ids, e, ('clickEffect' if g['trigger'] == 'click' else 'afterEffect')
                                   if (k == 0 and j == 0) else 'withEffect') for j, e in enumerate(effs))
            inner.append(f'<p:par><p:cTn id="{ids()}" fill="hold"><p:stCondLst><p:cond delay="{at}"/></p:stCondLst>'
                         f'<p:childTnLst>{pars}</p:childTnLst></p:cTn></p:par>')
        groups.append(f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst>{cond}</p:stCondLst>'
                      f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')

    # Aufbauliste: Formen mit Text brauchen einen Eintrag (ganze Form oder absatzweise)
    bld, seen = [], set()
    tags = {int(el.get('id')): el.getparent().getparent().tag.split('}')[1]
            for el in slide.shapes._spTree.iter(f'{{{P}}}cNvPr')}
    for g in tl.groups:
        for e in g['effects']:
            if e['spid'] in seen or tags.get(e['spid']) != 'sp':
                continue
            seen.add(e['spid'])
            build = ' build="p"' if e['para'] is not None else ' animBg="1"'
            bld.append(f'<p:bldP spid="{e["spid"]}" grpId="0"{build}/>')
    bld_xml = f'<p:bldLst>{"".join(bld)}</p:bldLst>' if bld else ''

    return (f'<p:timing {NSDECL}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            f'{"".join(groups)}</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>{bld_xml}</p:timing>')


def _insert(slide, xml):
    """Fügt Übergang oder Zeitleiste in der vom Schema verlangten Reihenfolge ein."""
    sld = slide._element
    el = etree.fromstring(xml)
    anchor = sld.find(f'{{{P}}}clrMapOvr')
    if el.tag == f'{{{P}}}timing':
        after = [c for c in sld if c.tag in (f'{{{P}}}transition',) or c.tag.endswith('}AlternateContent')]
        anchor = after[-1] if after else anchor
    anchor.addnext(el)


def set_timing(slide, tl):
    for old in slide._element.findall(f'{{{P}}}timing'):
        slide._element.remove(old)
    _insert(slide, timing_xml(tl, slide))


def set_transition(slide, kind='fade', dur=700):
    sld = slide._element
    for old in list(sld):
        if old.tag == f'{{{P}}}transition' or old.tag.endswith('}AlternateContent'):
            sld.remove(old)
    if kind == 'fade':
        xml = f'<p:transition {NSDECL} spd="med"><p:fade/></p:transition>'
    elif kind == 'morph':
        xml = (f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" {NSDECL}>'
               f'<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
               f'<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="{dur}">'
               f'<p159:morph option="byObject"/></p:transition></mc:Choice>'
               f'<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    else:
        raise ValueError(kind)
    _insert(slide, xml)


def rename(shape, name):
    """Morph ordnet Objekte mit gleichem Namen einander zu; «!!» erzwingt die Zuordnung."""
    shape._element.xpath('./*[1]/p:cNvPr')[0].set('name', name)

