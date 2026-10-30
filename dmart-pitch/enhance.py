# -*- coding: utf-8 -*-
"""Inject designer-grade slide transitions + staggered entrance animations into the .pptx."""
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Curated, varied transitions (designer progression, not a single repeated one)
TRANS = {
 'morph':     '<p:transition xmlns:p="{P}"><p:morph transition="byObject"/></p:transition>',
 'push_r':    '<p:transition xmlns:p="{P}"><p:push><p:dir val="r"/></p:push></p:transition>',
 'wipe_r':    '<p:transition xmlns:p="{P}"><p:wipe><p:dir val="r"/></p:wipe></p:transition>',
 'split':     '<p:transition xmlns:p="{P}"><p:split><p:dir val="out"/><p:orient val="horz"/></p:split></p:transition>',
 'cover_d':   '<p:transition xmlns:p="{P}"><p:cover><p:dir val="d"/></p:cover></p:transition>',
 'zoom_in':   '<p:transition xmlns:p="{P}"><p:zoom><p:dir val="in"/></p:zoom></p:transition>',
 'uncover_l': '<p:transition xmlns:p="{P}"><p:uncover><p:dir val="l"/></p:uncover></p:transition>',
 'reveal_r':  '<p:transition xmlns:p="{P}"><p:reveal><p:dir val="r"/></p:reveal></p:transition>',
 'fade':      '<p:transition xmlns:p="{P}"><p:fade><p:outClr dir="out"/></p:fade></p:transition>',
}
ORDER = ['morph', 'push_r', 'wipe_r', 'split', 'cover_d', 'zoom_in', 'uncover_l', 'reveal_r', 'fade', 'zoom_in']


def build_timing(targets):
    """targets: list of (spid, delay_ms). Builds a p:timing block with staggered fade-in entrances."""
    out = [f'<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" fill="remove">'
           f'<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" fill="remove"><p:childTnLst>']
    tid = 3
    for (spid, delay) in targets:
        out.append(f'<p:par><p:cTn id="{tid}" fill="remove" presetID="1" presetClass="entrance" presetSubtype="0">')
        out.append(f'<p:stCondLst><p:cond evt="onBegin" delay="{delay}"/></p:stCondLst>')
        out.append('<p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr>')
        out.append(f'<p:cTnLst><p:cn tnId="{tid}" gdRef="0" fill="hold"/></p:cTnLst>')
        out.append(f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>')
        out.append('<p:attrNameLst><p:attrName>visibility</p:attrName></p:attrNameLst>')
        out.append('</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
        tid += 1
    out.append('</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
    return ''.join(out)


def enhance(path):
    prs = Presentation(path)
    for i, slide in enumerate(prs.slides):
        sld = slide._element
        # --- transition (anchor after p:cSld; clrMap is optional/inherited) ---
        old = sld.find(qn('p:transition'))
        if old is not None:
            sld.remove(old)
        old_t = sld.find(qn('p:timing'))
        if old_t is not None:
            sld.remove(old_t)
        tr = parse_xml(TRANS[ORDER[i % len(ORDER)]].format(P=P))
        cSld = sld.find(qn('p:cSld'))
        pos = list(sld).index(cSld)
        sld.insert(pos + 1, tr)
        # --- staggered entrance animations ---
        targets = []
        di = 0
        for sh in slide.shapes:
            if sh.left is None or sh.width is None:
                continue
            if (sh.width / 914400.0) >= 13.0:   # skip the full-width top accent bar
                continue
            targets.append((sh.shape_id, min(di * 70, 1500)))
            di += 1
        if targets:
            sld.append(parse_xml(build_timing(targets)))
    prs.save(path)
    print("ENHANCED:", path, "| slides:", len(prs.slides._sldIdLst))


if __name__ == "__main__":
    enhance("/home/user/DMart_TrolleyRest_Pitch.pptx")
