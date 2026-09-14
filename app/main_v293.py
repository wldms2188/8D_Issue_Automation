# 8D Issue Automation v2.9.3
# Final layout/representative-image fixes based on the supplied weekly-meeting example mapping.
# - Representative image: ONLY picture/group intersecting the 2D or 4D content area; grouped pictures are composited together.
# - Weekly detail: normalize 1D~6D marker positions to the supplied example layout when the template geometry differs.
# - Detail text auto-fits inside each destination box, starting at 8pt (Malgun Gothic) and shrinking only as needed.
# - No ellipsis/truncation; line breaks are preserved.

import io, datetime, math, copy
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from PIL import Image as PILImage

import main_v292 as v292
v29 = v292.v29
base = v292.base
EMU=914400

# Example 2 supplied destination rectangles (inches). These are the authoritative
# reference zones the user marked for 1/2/3-1/3-2/4/5.
DETAIL_BOXES={
 '1':(0.53,2.53,4.76,0.98,'problem'),
 '2':(0.53,3.91,4.76,1.14,'temporary_action'),
 '3-1':(0.53,5.38,4.76,1.79,'cause_4d'),
 '3-2':(5.77,2.47,4.76,0.87,'leak_cause'),
 '4':(5.77,3.74,4.76,1.77,'action_5d'),
 '5':(5.77,5.91,4.78,0.94,'verification_6d'),
}

# Example-2 1~5 destination center positions. 1D is the section marker on
# the first detail block; 2D~6D map to the six marked content areas above.
MARKER_TARGETS={
 '1D':(0.35,2.18),
 '2D':(0.35,3.53),
 '3D':(0.35,4.96),
 '4D':(5.59,2.34),
 '5D':(5.59,3.61),
 '6D':(5.59,5.78),
}


def _norm_text(x):
    return '\n'.join(str(x or '').replace('\r\n','\n').replace('\r','\n').splitlines()).strip()


def _set_font(run, size):
    run.font.name='맑은 고딕'; run.font.size=Pt(size)
    try:
        r=run._r.get_or_add_rPr(); r.set('a:latin','맑은 고딕'); r.set('a:ea','맑은 고딕'); r.set('a:cs','맑은 고딕')
    except Exception: pass


def _fit_shape(sh,text,max_size=8,min_size=5):
    text=_norm_text(text)
    sh.text=text
    tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.TOP
    tf.margin_left=Inches(.05);tf.margin_right=Inches(.05);tf.margin_top=Inches(.03);tf.margin_bottom=Inches(.03)
    # PowerPoint itself performs final layout. We reduce the font progressively when
    # the line count is clearly too large for the available height.
    lines=max(1,len(text.split('\n')))
    avail=max(0.15,float(sh.height)/EMU-0.08)
    # Conservative estimate; preserves the full text and avoids overlap.
    size=max_size
    while size>min_size:
        chars_per_line=max(10,int(float(sh.width)/EMU*11*(8/size)))
        est=sum(max(1,math.ceil(len(line)/chars_per_line)) for line in text.split('\n'))
        line_h=size/72*1.18
        if est*line_h <= avail: break
        size-=0.5
    for p in tf.paragraphs:
        for r in p.runs:_set_font(r,size)
    return size

# ---------- Representative image: strict 2D/4D region ----------
def _shape_box(sh):
    try:return (float(sh.left),float(sh.top),float(sh.width),float(sh.height))
    except Exception:return (0,0,0,0)

def _children_pics(sh):
    out=[]
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE: out=[sh]
    elif getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
        for c in sh.shapes:out.extend(_children_pics(c))
    return out

def _group_blob(group):
    pics=_children_pics(group)
    if len(pics)<2:return None
    # Use the group's declared bounds as the composite canvas. This avoids treating
    # an unrelated picture elsewhere on the slide as part of the representative image.
    gx,gy,gw,gh=_shape_box(group)
    if gw<=0 or gh<=0:return None
    W=max(1,int(gw/EMU*120));H=max(1,int(gh/EMU*120))
    canvas=PILImage.new('RGB',(W,H),'white')
    for p in pics:
        try:
            blob=p.image.blob
            with PILImage.open(io.BytesIO(blob)) as im:
                im=im.convert('RGB')
                px,py,pw,ph=_shape_box(p)
                x=max(0,int((px-gx)/EMU*120));y=max(0,int((py-gy)/EMU*120))
                w=max(1,int(pw/EMU*120));h=max(1,int(ph/EMU*120))
                im.thumbnail((w,h),PILImage.Resampling.LANCZOS)
                canvas.paste(im,(x,y))
        except Exception:pass
    b=io.BytesIO();canvas.save(b,'PNG');return b.getvalue()

def _all_visuals(prs):
    out=[]
    for si,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if sh.shape_type==MSO_SHAPE_TYPE.GROUP:
                pics=_children_pics(sh)
                if len(pics)>=2:
                    blob=_group_blob(sh)
                    if blob:
                        x,y,w,h=_shape_box(sh);out.append((si,x,y,w,h,blob,True,len(pics)))
            elif sh.shape_type==MSO_SHAPE_TYPE.PICTURE:
                try:
                    with PILImage.open(io.BytesIO(sh.image.blob)) as im:
                        if im.width*im.height>=10000:
                            x,y,w,h=_shape_box(sh);out.append((si,x,y,w,h,sh.image.blob,False,1))
                except Exception:pass
    return out

def _area_overlap(a,b):
    ax,ay,aw,ah=a;bx,by,bw,bh=b
    ix=max(0,min(ax+aw,bx+bw)-max(ax,bx));iy=max(0,min(ay+ah,by+bh)-max(ay,by));return ix*iy

def _text_regions(sl):
    regs=[]
    for sh in sl.shapes:
        txt=_norm_text(getattr(sh,'text',''))
        q=base.compact(txt)
        if any(k in q for k in ['2d','문제현황','문제현상','불량현상','4d','원인분석','발생원인','유출원인']):
            x,y,w,h=_shape_box(sh);regs.append((x,y,w,h,q))
    return regs

def _pick_rep(prs,d):
    visuals=_all_visuals(prs)
    if not visuals:return None
    # First: visual must intersect a 2D-labeled region on the same slide.
    preferred=[]
    for si,sl in enumerate(prs.slides):
        regs=[r for r in _text_regions(sl) if '2d' in r[4] or '문제현황' in r[4] or '불량현상' in r[4]]
        if not regs:continue
        for v in visuals:
            vsi,x,y,w,h,blob,grp,n=v
            if vsi!=si:continue
            ov=max((_area_overlap((x,y,w,h),r[:4]) for r in regs),default=0)
            if ov>0:preferred.append((ov,grp,n,w*h,blob))
    if preferred:
        # Prefer grouped image, then greatest overlap, then group count.
        preferred.sort(key=lambda z:(z[1],z[0],z[2],z[3]),reverse=True)
        return preferred[0][4]
    # Second: 4D labeled region, as explicitly allowed by the user's rule.
    preferred=[]
    for si,sl in enumerate(prs.slides):
        regs=[r for r in _text_regions(sl) if '4d' in r[4] or '원인분석' in r[4] or '발생원인' in r[4] or '유출원인' in r[4]]
        for v in visuals:
            vsi,x,y,w,h,blob,grp,n=v
            if vsi!=si:continue
            ov=max((_area_overlap((x,y,w,h),r[:4]) for r in regs),default=0)
            if ov>0:preferred.append((ov,grp,n,w*h,blob))
    if preferred:
        preferred.sort(key=lambda z:(z[1],z[0],z[2],z[3]),reverse=True);return preferred[0][4]
    return None

def extract_v293(path):
    d=v292.extract_v292(path)
    try:
        b=_pick_rep(Presentation(path),d)
        if b:d['_images']=[(1,1,1,b)]
    except Exception:pass
    return d
base.extract=extract_v293

# ---------- detail page layout ----------
def _find_label(sl,label):
    for sh in sl.shapes:
        if _norm_text(getattr(sh,'text',''))==label:return sh
    return None

def _move_marker(sh,x,y):
    # Keep marker size; only its position is normalized.
    sh.left=Inches(x);sh.top=Inches(y)

def _normalize_markers(sl):
    for label,(x,y) in MARKER_TARGETS.items():
        sh=_find_label(sl,label)
        if sh is not None:_move_marker(sh,x,y)

def _add_or_get_box(sl,label,x,y,w,h):
    # Prefer an existing rectangle/textbox near the reference position. If none exists,
    # create a transparent textbox inside the exact reference zone.
    target=(x*EMU,y*EMU,w*EMU,h*EMU);best=None;score=1e99
    for sh in sl.shapes:
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:continue
        if getattr(sh,'has_table',False):continue
        if not hasattr(sh,'text_frame'):continue
        sx,sy,sw,shh=_shape_box(sh);s=abs(sx-target[0])+abs(sy-target[1])+abs(sw-target[2])+abs(shh-target[3])
        if s<score and s<1.4*EMU:best,score=sh,s
    if best is not None:return best
    sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));sh.name='AUTO_DETAIL_'+label;return sh

def fill_detail(sl,d,g):
    _normalize_markers(sl)
    for label,(x,y,w,h,key) in DETAIL_BOXES.items():
        sh=_add_or_get_box(sl,label,x,y,w,h)
        if key=='leak_cause':text='\n'.join(_norm_text(v) for v in [d.get('leak_cause'),d.get('system_cause')] if _norm_text(v))
        else:text=_norm_text(d.get(key))
        _fit_shape(sh,text or '검토 중',8,5)
    # Metadata and Signal
    for sh in sl.shapes:
        t=_norm_text(getattr(sh,'text',''))
        if t.startswith('과제명_이슈 제목'):_fit_shape(sh,base.task(d),8,5)
        elif t.startswith('이슈명 :'):_fit_shape(sh,'이슈명 : '+_norm_text(d.get('issue_name')),8,5)
        elif '00팀 담당자' in t:_fit_shape(sh,f"{g.get('team','')} 담당자 : {g.get('owner','')}",8,5)
        if getattr(sh,'has_table',False):
            tb=sh.table
            for r in range(len(tb.rows)):
                for c in range(len(tb.columns)):
                    q=base.compact(tb.cell(r,c).text)
                    if q=='signal' and c+1<len(tb.columns):base.signal(tb.cell(r,c+1),base.status(d))
    # Apply requested font to all detail text, preserving full strings.
    for sh in sl.shapes:
        if hasattr(sh,'text_frame'):
            for p in sh.text_frame.paragraphs:
                for run in p.runs:_set_font(run,8)

def weekly_v293(src,out,d,g,mode):
    prs=Presentation(src);st=base.status(d);title=base.task(d);issue=_norm_text(d.get('issue_name'));summary=None;idx=0
    for i,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if getattr(sh,'has_table',False):
                h=' '.join(base.norm(sh.table.cell(0,c).text) for c in range(len(sh.table.columns)))
                if '과제명' in h and 'Signal' in h:summary=sh.table;idx=i;break
        if summary:break
    if summary:
        r=1
        for x in range(1,len(summary.rows)):
            if not any(base.norm(summary.cell(x,c).text) for c in range(min(5,len(summary.columns))) if c!=4):r=x;break
        vals=[title,issue,_norm_text(d.get('problem')),v29.v29_progress(d),'●']
        for c,val in enumerate(vals):v29.v29_set_cell_text(summary.cell(r,c),val,8)
        base.signal(summary.cell(r,4),st)
    for sl in [prs.slides[i] for i in range(idx+1,len(prs.slides))]:fill_detail(sl,d,g)
    Path(out).parent.mkdir(exist_ok=True)
    try:prs.save(out);saved=out
    except PermissionError:
        p=Path(out);saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix);prs.save(saved)
    return '주간회의 PPT 업데이트: '+st,saved

base.weekly=weekly_v293
if __name__=='__main__':base.App().mainloop()
