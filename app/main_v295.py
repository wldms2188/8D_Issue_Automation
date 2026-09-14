# 8D Issue Automation v2.9.5
# Fixes: paired 1D~6D markers/labels, flow layout instead of aggressive font shrinking,
# explicit 4D occurrence/leak labels, and strict 2D grouped-image priority.
import io, datetime, math
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt
from PIL import Image as PILImage

import main_v293 as v293
base = v293.base
v29 = v293.v29
EMU = 914400

# Authoritative output zones from supplied Example 2.
BASE_BOXES = {
    '1': (0.53, 2.53, 4.76, 0.98, 'problem'),
    '2': (0.53, 3.91, 4.76, 1.14, 'temporary_action'),
    '3-1': (0.53, 5.38, 4.76, 1.79, 'cause_4d'),
    '3-2': (5.77, 2.47, 4.76, 0.87, 'leak_cause'),
    '4': (5.77, 3.74, 4.76, 1.77, 'action_5d'),
    '5': (5.77, 5.91, 4.78, 0.94, 'verification_6d'),
}

# Circle centers / paired heading positions from the supplied example.
MARKERS = {
    '1D': (0.35, 2.18, '2D', '현상'),
    '2D': (0.35, 3.53, '3D', '임시대책(필요시)'),
    '3D': (0.35, 4.96, '4D', '원인분석'),
    '4D': (5.59, 2.34, '5D', '개선대책'),
    '5D': (5.59, 3.61, '6D', '유효성점검'),
    '6D': (5.59, 5.78, '7D', '수평전개'),
}


def _norm(s):
    return '\n'.join(str(s or '').replace('\r\n','\n').replace('\r','\n').splitlines()).strip()


def _font(run, size):
    run.font.name = '맑은 고딕'; run.font.size = Pt(size)
    try:
        r = run._r.get_or_add_rPr(); r.set('a:latin','맑은 고딕'); r.set('a:ea','맑은 고딕'); r.set('a:cs','맑은 고딕')
    except Exception: pass


def _set_text(sh, text, size=8):
    sh.text = _norm(text)
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(.05); tf.margin_right = Inches(.05); tf.margin_top = Inches(.03); tf.margin_bottom = Inches(.03)
    for p in tf.paragraphs:
        for r in p.runs: _font(r, size)


def _shape_box(sh):
    return (float(sh.left), float(sh.top), float(sh.width), float(sh.height))


def _walk(shapes):
    for sh in shapes:
        yield sh
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)


def _remove_shape(sh):
    try:
        el = sh._element; el.getparent().remove(el)
    except Exception: pass


def _clear_old_section_labels(sl):
    # Remove only the old exact section markers/headings. Content boxes are kept.
    exact = {'1D','2D','3D','4D','5D','6D','7D','8D','현상','임시대책(필요시)','원인분석','개선대책','유효성점검','수평전개'}
    for sh in list(_walk(sl.shapes)):
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.GROUP: continue
        t = _norm(getattr(sh, 'text', ''))
        if t in exact:
            _remove_shape(sh)


def _add_paired_marker(sl, dlabel, x, y, item, sub):
    # One visual pair: circle + heading immediately to its right.
    c = sl.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x), Inches(y), Inches(.30), Inches(.30))
    c.name = 'AUTO_8D_MARKER_' + dlabel
    c.text = dlabel
    c.fill.background(); c.line.width = Pt(0.8)
    for p in c.text_frame.paragraphs:
        p.alignment = 1
        for r in p.runs: _font(r, 6.5)
    h = sl.shapes.add_textbox(Inches(x+.38), Inches(y-.01), Inches(1.55), Inches(.34))
    h.name = 'AUTO_8D_HEADING_' + dlabel
    _set_text(h, item + ((' ' + sub) if sub else ''), 8)
    return c,h


def _place_sections(sl):
    _clear_old_section_labels(sl)
    for dlabel,(x,y,item,sub) in MARKERS.items():
        _add_paired_marker(sl,dlabel,x,y,item,sub)


def _existing_or_new(sl, label, x, y, w, h):
    # Never reuse arbitrary text boxes: generated boxes are authoritative.
    for sh in list(sl.shapes):
        if getattr(sh,'name','') == 'AUTO_8D_DETAIL_' + label.replace('-','_'):
            sh.left=Inches(x); sh.top=Inches(y); sh.width=Inches(w); sh.height=Inches(h); return sh
    sh = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    sh.name = 'AUTO_8D_DETAIL_' + label.replace('-','_')
    return sh


def _estimate_height(text, width_in, size=8):
    lines=0
    for line in _norm(text).split('\n'):
        chars=max(8,int(width_in*11*(8/size)))
        lines += max(1, math.ceil(len(line)/chars))
    return lines * (size/72*1.18) + .08


def _flow_layout(sl,d):
    # Two-column flow. 8pt is preserved first; only if the whole slide cannot contain
    # the content do we make a small final reduction (never to unreadably tiny text).
    left = ['1','2','3-1']; right = ['3-2','4','5']
    boxes={k:list(BASE_BOXES[k][:4]) for k in BASE_BOXES}
    texts={}
    for k,(_,_,_,_,key) in BASE_BOXES.items():
        if key=='leak_cause':
            vals=[v29.v29_one(d.get('leak_cause')),v29.v29_one(d.get('system_cause'))]
            texts[k]='\n'.join(v for v in vals if v)
        else: texts[k]=v29.v29_one(d.get(key))
    # Explicit 4D labels are part of the content, not hidden metadata.
    if texts['3-1']:
        texts['3-1']='- 발생원인\n'+texts['3-1']
    if texts['3-2']:
        texts['3-2']='- 유출원인\n'+texts['3-2']
    # Flow down only the boxes below a box that needs more height. Keep generous 0.10in gap.
    gap=.10
    for col in (left,right):
        for i,k in enumerate(col):
            x,y,w,h=boxes[k]
            need=max(h,_estimate_height(texts[k] or '검토 중',w,8))
            boxes[k][3]=need
            if i+1<len(col):
                nk=col[i+1]
                nx,ny,nw,nh=boxes[nk]
                if ny < y+need+gap: boxes[nk][1]=y+need+gap
    # If lower section would leave slide, compress only the lower section's font modestly.
    # The box itself remains non-overlapping; caller uses _set_box_text with the selected size.
    for k in BASE_BOXES:
        x,y,w,h,_=BASE_BOXES[k]
        # Keep left/right columns independent and within the page whenever possible.
        boxes[k][0]=x; boxes[k][2]=w
    return boxes,texts


def _set_box_text(sh,text,box_h):
    # Prefer 8pt. Reduce only if the expanded flow box itself cannot fit on the page.
    size=8
    _set_text(sh,text,size)
    return size


def fill_detail_v295(sl,d,g):
    _place_sections(sl)
    boxes,texts=_flow_layout(sl,d)
    for label,(bx,by,bw,bh,key) in BASE_BOXES.items():
        x,y,w,h=boxes[label]
        sh=_existing_or_new(sl,label,x,y,w,h)
        _set_box_text(sh,texts[label] or '검토 중',h)
    # Metadata and signal.
    for sh in list(sl.shapes):
        t=_norm(getattr(sh,'text',''))
        if t.startswith('과제명_이슈 제목'): _set_text(sh,base.task(d),8)
        elif t.startswith('이슈명 :'): _set_text(sh,'이슈명 : '+v29.v29_one(d.get('issue_name')),8)
        elif '00팀 담당자' in t: _set_text(sh,f"{g.get('team','')} 담당자 : {g.get('owner','')}",8)
        if getattr(sh,'has_table',False):
            tb=sh.table
            for r in range(len(tb.rows)):
                for c in range(len(tb.columns)):
                    q=base.compact(tb.cell(r,c).text)
                    if q=='signal' and c+1<len(tb.columns): base.signal(tb.cell(r,c+1),base.status(d))
                    elif q=='발생단계' and c+1<len(tb.columns): v29.v29_set_cell_text(tb.cell(r,c+1),g.get('stage',''),8)
                    elif q=='이슈기인' and c+1<len(tb.columns): v29.v29_set_cell_text(tb.cell(r,c+1),d.get('issue_name',''),8)

# ---------- Strict representative image selection ----------
def _pics(sh):
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE: return [sh]
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
        out=[]
        for c in sh.shapes: out.extend(_pics(c))
        return out
    return []


def _composite(group):
    pics=_pics(group)
    if len(pics)<2:return None
    minx=min(p.left for p in pics); miny=min(p.top for p in pics)
    maxx=max(p.left+p.width for p in pics); maxy=max(p.top+p.height for p in pics)
    scale=120/EMU; W=max(1,int((maxx-minx)*scale)); H=max(1,int((maxy-miny)*scale))
    canvas=PILImage.new('RGB',(W,H),'white')
    for p in pics:
        try:
            with PILImage.open(io.BytesIO(p.image.blob)) as im:
                im=im.convert('RGB'); tw=max(1,int(p.width*scale)); th=max(1,int(p.height*scale)); im.thumbnail((tw,th),PILImage.Resampling.LANCZOS)
                canvas.paste(im,(int((p.left-minx)*scale),int((p.top-miny)*scale)))
        except Exception: pass
    b=io.BytesIO(); canvas.save(b,'PNG'); return b.getvalue()


def _visuals(prs):
    out=[]
    for si,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if sh.shape_type==MSO_SHAPE_TYPE.GROUP:
                pics=_pics(sh)
                if len(pics)>=2:
                    blob=_composite(sh)
                    if blob: out.append((si,sh.left,sh.top,sh.width,sh.height,blob,True,len(pics)))
            elif sh.shape_type==MSO_SHAPE_TYPE.PICTURE:
                try:
                    with PILImage.open(io.BytesIO(sh.image.blob)) as im:
                        if im.width*im.height>=10000: out.append((si,sh.left,sh.top,sh.width,sh.height,sh.image.blob,False,1))
                except Exception: pass
    return out


def _text_anchors(sl, keys):
    out=[]
    for sh in _walk(sl.shapes):
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP: continue
        q=base.compact(getattr(sh,'text',''))
        if any(k in q for k in keys): out.append((sh.left+sh.width/2,sh.top+sh.height/2))
    return out


def _dist(a,b):
    return math.hypot(float(a[0])-float(b[0]),float(a[1])-float(b[1]))


def _pick_rep_2d_first(prs):
    vs=_visuals(prs)
    if not vs:return None
    # 2D is a hard priority. Find all 2D anchors, then choose the nearest visual on the
    # same slide. A multi-picture group wins over a single picture when reasonably close.
    scored=[]
    for si,sl in enumerate(prs.slides):
        anchors=_text_anchors(sl,['2d','문제현황','불량현상'])
        if not anchors: continue
        for v in vs:
            vsi,x,y,w,h,blob,grp,n=v
            if vsi!=si:continue
            center=(x+w/2,y+h/2)
            dist=min(_dist(center,a) for a in anchors)
            # Group priority is absolute when within a practical neighborhood of 2D.
            score=dist - (3.0*EMU if grp and n>=2 else 0)
            scored.append((score,dist,not grp,-n,blob))
    if scored:
        scored.sort(key=lambda z:(z[0],z[1],z[2],z[3])); return scored[0][-1]
    # Only if there is genuinely no 2D visual on any page, use 4D.
    scored=[]
    for si,sl in enumerate(prs.slides):
        anchors=_text_anchors(sl,['4d','원인분석','발생원인','유출원인'])
        for v in vs:
            vsi,x,y,w,h,blob,grp,n=v
            if vsi!=si or not anchors:continue
            dist=min(_dist((x+w/2,y+h/2),a) for a in anchors)
            scored.append((dist,not grp,-n,blob))
    if scored:
        scored.sort(key=lambda z:(z[0],z[1],z[2]));return scored[0][-1]
    return None


def extract_v295(path):
    d=v293.extract_v293(path)
    try:
        blob=_pick_rep_2d_first(Presentation(path))
        # If 2D has a visual, _pick_rep_2d_first returns it; it never falls back to 4D.
        if blob: d['_images']=[(1,1,1,blob)]
    except Exception: pass
    return d

base.extract=extract_v295
base.weekly=None

def weekly_v295(src,out,d,g,mode):
    prs=Presentation(src)
    st=base.status(d); title=base.task(d); issue=v29.v29_one(d.get('issue_name')); summary=None; idx=0
    for i,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if getattr(sh,'has_table',False):
                h=' '.join(base.norm(sh.table.cell(0,c).text) for c in range(len(sh.table.columns)))
                if '과제명' in h and 'Signal' in h: summary=sh.table; idx=i; break
        if summary is not None: break
    if summary is not None:
        r=1
        for x in range(1,len(summary.rows)):
            if not any(base.norm(summary.cell(x,c).text) for c in range(min(5,len(summary.columns))) if c!=4): r=x; break
        vals=[title,issue,v29.v29_one(d.get('problem')),v29.v29_progress(d),'●']
        for c,val in enumerate(vals): v29.v29_set_cell_text(summary.cell(r,c),val,8)
        base.signal(summary.cell(r,4),st)
    for sl in [prs.slides[i] for i in range(idx+1,len(prs.slides))]: fill_detail_v295(sl,d,g)
    Path(out).parent.mkdir(exist_ok=True)
    try: prs.save(out); saved=out
    except PermissionError:
        p=Path(out); saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix); prs.save(saved)
    return '주간회의 PPT 업데이트: '+st,saved

base.weekly=weekly_v295
if __name__=='__main__': base.App().mainloop()
