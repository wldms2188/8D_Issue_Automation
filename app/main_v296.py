# 8D Issue Automation v2.9.6
# Fixes requested after v2.9.5:
# 1) 8D SOURCE extraction is strictly page 1 only (text + representative image).
# 2) Weekly detail layout is rebuilt from content zones; marker + heading move with
#    their zone instead of remaining at old template coordinates.
# 3) Current section grows first; following sections are pushed down. Font remains
#    8pt unless the page physically cannot contain the expanded layout.
# 4) 4D always displays explicit - 발생원인 / - 유출원인 labels.
import io, datetime, math, tempfile, os
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from PIL import Image as PILImage

import main_v295 as v295
base = v295.base
v29 = v295.v29
EMU = 914400

# Example-2 destination zones. These are the starting geometry; vertical positions
# are allowed to move downward as the preceding section grows.
BOXES = {
    '1': [0.53, 2.53, 4.76, 0.98, 'problem'],
    '2': [0.53, 3.91, 4.76, 1.14, 'temporary_action'],
    '3-1': [0.53, 5.38, 4.76, 1.79, 'cause_4d'],
    '3-2': [5.77, 2.47, 4.76, 0.87, 'leak_cause'],
    '4': [5.77, 3.74, 4.76, 1.77, 'action_5d'],
    '5': [5.77, 5.91, 4.78, 0.94, 'verification_6d'],
}

# Marker/heading pair belongs to the content box immediately below it.
SECTION_BOX = {'1D':'1','2D':'2','3D':'3-1','4D':'3-2','5D':'4','6D':'5'}


def _norm(s):
    return '\n'.join(str(s or '').replace('\r\n','\n').replace('\r','\n').splitlines()).strip()


def _font(run, size=8):
    run.font.name='맑은 고딕'; run.font.size=Pt(size)
    try:
        r=run._r.get_or_add_rPr(); r.set('a:latin','맑은 고딕'); r.set('a:ea','맑은 고딕'); r.set('a:cs','맑은 고딕')
    except Exception: pass


def _set_text(sh,text,size=8):
    sh.text=_norm(text)
    tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.TOP
    tf.margin_left=Inches(.05); tf.margin_right=Inches(.05); tf.margin_top=Inches(.03); tf.margin_bottom=Inches(.03)
    for p in tf.paragraphs:
        for r in p.runs:_font(r,size)


def _walk(shapes):
    for sh in shapes:
        yield sh
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)


def _remove(sh):
    try:
        e=sh._element; e.getparent().remove(e)
    except Exception: pass


def _estimate_lines(text,width_in,size=8):
    # Conservative PowerPoint-style line estimate. Explicit blank lines count.
    n=0
    chars=max(8,int(width_in*10.5*(8/size)))
    for line in _norm(text).split('\n'):
        n += max(1,math.ceil(len(line)/chars))
    return n


def _needed_height(text,width_in,size=8,min_h=.45):
    return max(min_h, _estimate_lines(text,width_in,size)*(size/72*1.20)+.10)


def _texts(d):
    out={}
    for label,(_,_,w,_,key) in BOXES.items():
        if key=='cause_4d':
            out[label]='- 발생원인\n'+(v29.v29_one(d.get('cause_4d')) or '검토 중')
        elif key=='leak_cause':
            leak=v29.v29_one(d.get('leak_cause')) or '검토 중'
            sys=v29.v29_one(d.get('system_cause'))
            out[label]='- 유출원인\n'+leak+(('\n\n- 시스템/관리 원인\n'+sys) if sys else '')
        else:
            out[label]=v29.v29_one(d.get(key)) or '검토 중'
    return out


def _delete_generated(sl):
    for sh in list(_walk(sl.shapes)):
        try:
            n=str(sh.name)
            if n.startswith('AUTO_8D_') or n.startswith('AUTO295_'):
                _remove(sh)
        except Exception: pass


def _delete_old_section_text(sl):
    # Remove old labels/headings only. Existing tables and other user content remain.
    exact={'1D','2D','3D','4D','5D','6D','7D','8D','현상','임시대책(필요시)','원인분석','개선대책','유효성점검','수평전개'}
    for sh in list(_walk(sl.shapes)):
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP: continue
        if _norm(getattr(sh,'text','')) in exact:_remove(sh)


def _make_pair(sl,dlabel,x,y,item):
    # Marker and heading are generated together from the same x/y anchor.
    c=sl.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x), Inches(y), Inches(.30), Inches(.30))
    c.name='AUTO_8D_MARKER_'+dlabel
    c.text=dlabel
    try:c.fill.background(); c.line.width=Pt(.8)
    except Exception:pass
    for p in c.text_frame.paragraphs:
        p.alignment=1
        for r in p.runs:_font(r,6.5)
    h=sl.shapes.add_textbox(Inches(x+.38),Inches(y-.01),Inches(2.2),Inches(.34))
    h.name='AUTO_8D_HEADING_'+dlabel
    _set_text(h,item,8)


def _layout(sl,d):
    _delete_generated(sl); _delete_old_section_text(sl)
    texts=_texts(d)
    # Two independent vertical flows. A section's required height is applied before
    # placing the next section. This is the important difference from v2.9.5.
    left=['1','2','3-1']; right=['3-2','4','5']
    boxes={k:list(v[:4]) for k,v in BOXES.items()}
    gap=.14; heading_gap=.10; marker_x_left=.25; marker_x_right=5.49

    for col in (left,right):
        for i,k in enumerate(col):
            x,y,w,h=boxes[k]
            required=_needed_height(texts[k],w,8)
            boxes[k][3]=max(h,required)
            if i+1<len(col):
                nk=col[i+1]
                boxes[nk][1]=max(boxes[nk][1], boxes[k][1]+boxes[k][3]+gap+.30+heading_gap)

    # If content pushes beyond the bottom, make the affected lower boxes taller only
    # when possible; never allow a box to overlap the next box. We keep 8pt as default.
    # When page space is exhausted, only the text is reduced modestly to 7pt.
    slide_h=float(sl.part.presentation.slide_height)/EMU if hasattr(sl.part,'presentation') else 7.5
    for k in boxes:
        if boxes[k][1]+boxes[k][3] > slide_h-.10:
            boxes[k][3]=max(.55,slide_h-.10-boxes[k][1])

    # Marker/heading follows its own box. Therefore moving a content box also moves its
    # corresponding D marker and heading as a single unit.
    items={'1':'2D 현상','2':'3D 임시대책(필요시)','3-1':'4D 원인분석','3-2':'4D 원인분석','4':'5D 개선대책','5':'6D 유효성점검'}
    dlabels={'1':'1D','2':'2D','3-1':'3D','3-2':'4D','4':'5D','5':'6D'}
    for k,dlabel in dlabels.items():
        x,y,w,h=boxes[k]
        mx=marker_x_left if k in left else marker_x_right
        _make_pair(sl,dlabel,mx,y-.38,items[k])

    for label,(ox,oy,ow,oh,key) in BOXES.items():
        x,y,w,h=boxes[label]
        sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
        sh.name='AUTO_8D_DETAIL_'+label.replace('-','_')
        # 8pt is the standard. Only use 7pt if the physically bounded box is too short.
        size=8
        if _needed_height(texts[label],w,8) > h+.02:size=7
        _set_text(sh,texts[label],size)


def _clone_first_slide_only(path):
    """Make a temporary PPT containing only source slide 1, so inherited extractor
    cannot accidentally consume pages 2+ for 8D fields."""
    src=Presentation(path)
    out=Presentation()
    # Match page size.
    out.slide_width=src.slide_width; out.slide_height=src.slide_height
    blank=out.slide_layouts[6]
    new=out.slides.add_slide(blank)
    source=src.slides[0]
    for shape in source.shapes:
        try:
            new.shapes._spTree.insert_element_before(shape._element, 'p:extLst')
        except Exception:
            pass
    fd=tempfile.NamedTemporaryFile(delete=False,suffix='.pptx')
    fd.close(); out.save(fd.name)
    return fd.name


def _rep_page1(path):
    # Reuse v2.9.5 strict 2D-first selector, but it is called on the original PPT and
    # therefore can only return page-1 content when the selector is explicitly limited.
    prs=Presentation(path)
    if len(prs.slides)==0:return None
    sl=prs.slides[0]
    visuals=[]
    def pics(sh):
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:return [sh]
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
            z=[]
            for c in sh.shapes:z+=pics(c)
            return z
        return []
    def comp(g):
        ps=pics(g)
        if len(ps)<2:return None
        minx=min(p.left for p in ps);miny=min(p.top for p in ps)
        maxx=max(p.left+p.width for p in ps);maxy=max(p.top+p.height for p in ps)
        sc=120/EMU; W=max(1,int((maxx-minx)*sc));H=max(1,int((maxy-miny)*sc))
        canvas=PILImage.new('RGB',(W,H),'white')
        for p in ps:
            try:
                with PILImage.open(io.BytesIO(p.image.blob)) as im:
                    im=im.convert('RGB');tw=max(1,int(p.width*sc));th=max(1,int(p.height*sc));im.thumbnail((tw,th),PILImage.Resampling.LANCZOS)
                    canvas.paste(im,(int((p.left-minx)*sc),int((p.top-miny)*sc)))
            except Exception:pass
        b=io.BytesIO();canvas.save(b,'PNG');return b.getvalue()
    for sh in sl.shapes:
        if sh.shape_type==MSO_SHAPE_TYPE.GROUP:
            ps=pics(sh)
            if len(ps)>=2:
                b=comp(sh)
                if b:visuals.append((sh.left,sh.top,sh.width,sh.height,b,True,len(ps)))
        elif sh.shape_type==MSO_SHAPE_TYPE.PICTURE:
            try:
                with PILImage.open(io.BytesIO(sh.image.blob)) as im:
                    if im.width*im.height>=10000:visuals.append((sh.left,sh.top,sh.width,sh.height,sh.image.blob,False,1))
            except Exception:pass
    anchors=[]
    for sh in sl.shapes:
        q=base.compact(getattr(sh,'text',''))
        if any(k in q for k in ('2d','문제현황','문제현상','불량현상')):anchors.append((sh.left+sh.width/2,sh.top+sh.height/2))
    if not visuals:return None
    if anchors:
        hit=[]
        for v in visuals:
            dist=min(math.hypot(float(v[0]+v[2]/2-a[0]),float(v[1]+v[3]/2-a[1])) for a in anchors)
            hit.append((0 if v[5] and v[6]>=2 else 1,dist,v))
        hit.sort(key=lambda z:(z[0],z[1]));return hit[0][2][4]
    groups=[v for v in visuals if v[5]]
    return max(groups,key=lambda v:v[6])[4] if groups else None


def extract_v296(path):
    temp=None
    try:
        # Critical: extract all semantic fields from a PPT that contains ONLY page 1.
        temp=_clone_first_slide_only(path)
        d=v293.extract_v293(temp)
        b=_rep_page1(path)
        d['_images']=[(1,1,1,b)] if b else []
        return d
    finally:
        if temp:
            try:os.unlink(temp)
            except Exception:pass


def weekly_v296(src,out,d,g,mode):
    prs=Presentation(src)
    st=base.status(d); title=base.task(d); issue=v29.v29_one(d.get('issue_name')); summary=None;idx=0
    for i,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if getattr(sh,'has_table',False):
                h=' '.join(base.norm(sh.table.cell(0,c).text) for c in range(len(sh.table.columns)))
                if '과제명' in h and 'Signal' in h:summary=sh.table;idx=i;break
        if summary is not None:break
    if summary is not None:
        r=1
        for rr in range(1,len(summary.rows)):
            if not any(base.norm(summary.cell(rr,c).text) for c in range(min(5,len(summary.columns))) if c!=4):r=rr;break
        vals=[title,issue,norm(d.get('problem')),v29.v29_progress(d),'●']
        for c,v in enumerate(vals):v29.v29_set_cell_text(summary.cell(r,c),v,8)
        base.signal(summary.cell(r,4),st)
    for sl in prs.slides[idx+1:]:
        _layout(sl,d)
        for sh in sl.shapes:
            t=_norm(getattr(sh,'text',''))
            if t.startswith('과제명_이슈 제목'): _set_text(sh,title,8)
            elif t.startswith('이슈명 :'): _set_text(sh,'이슈명 : '+issue,8)
            elif '00팀 담당자' in t:_set_text(sh,f"{g.get('team','')} 담당자 : {g.get('owner','')}",8)
            if getattr(sh,'has_table',False):
                tb=sh.table
                for rr in range(len(tb.rows)):
                    for c in range(len(tb.columns)):
                        if base.compact(tb.cell(rr,c).text)=='signal' and c+1<len(tb.columns):base.signal(tb.cell(rr,c+1),st)
    Path(out).parent.mkdir(exist_ok=True)
    try:prs.save(out);saved=out
    except PermissionError:
        p=Path(out);saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix);prs.save(saved)
    return '주간회의 PPT 업데이트: '+st,saved

base.extract=extract_v296
base.weekly=weekly_v296

if __name__=='__main__':base.App().mainloop()
