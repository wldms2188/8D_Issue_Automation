# 8D Issue Automation v3.0.8
# Focused correction for weekly PPT: preserve template formatting, exact metadata mapping,
# stable 2D~6D marker handling, and controlled multi-image rendering.
import sys, re, io, copy, datetime
from pathlib import Path
APP_DIR=Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path: sys.path.insert(0,str(APP_DIR))
import main_v306 as impl
base=impl.base
renderer=impl.renderer
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from PIL import Image as PILImage, ImageOps, ImageDraw
EMU=914400

# ---------- text ----------
def N(x): return str(x or '').replace('\r\n','\n').replace('\r','\n').strip()
def C(x):
    try:return base.compact(x)
    except:return re.sub(r'[^0-9A-Za-z가-힣]','',N(x)).lower()

def set_run(run,size=None):
    if size is not None: run.font.size=Pt(size)
    run.font.name='맑은 고딕'
    try:
        r=run._r.get_or_add_rPr(); r.set('a:latin','맑은 고딕'); r.set('a:ea','맑은 고딕'); r.set('a:cs','맑은 고딕')
    except Exception: pass

def set_cell(cell,text,size=8):
    cell.text=N(text); cell.text_frame.word_wrap=True
    for p in cell.text_frame.paragraphs:
        for r in p.runs:set_run(r,size)

def set_shape(sh,text,size=8):
    if not hasattr(sh,'text_frame'): return
    tf=sh.text_frame; tf.clear(); tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.TOP
    p=tf.paragraphs[0]; p.text=N(text)
    for r in p.runs:set_run(r,size)

# ---------- exact task / issue extraction ----------
def task_from_customer_issue(path,d):
    prs=Presentation(path); customer=N(d.get('customer')); issue=N(d.get('issue_name'))
    # The requested rule: Ford_xxxx -> xxxx. Do not accept unrelated labels such as 상세시험조건.
    texts=[]
    for sl in prs.slides:
        for sh in sl.shapes:
            if getattr(sh,'has_table',False):
                tb=sh.table
                for r in range(len(tb.rows)):
                    for c in range(len(tb.columns)):
                        texts.append(N(tb.cell(r,c).text))
            t=N(getattr(sh,'text',''))
            if t:texts.append(t)
    if customer:
        pat=re.compile(r'(?<![\w가-힣])'+re.escape(customer)+r'\s*[_\-/／|:]\s*([^\n\r,;|]+)',re.I)
        for t in texts:
            m=pat.search(t)
            if m:
                v=N(m.group(1)).strip(' _-/／|:：')
                if v and C(v) not in {'상세시험조건','model','packer'}: return v
    # If issue_name itself contains Ford_xxxx, use only suffix.
    if customer and issue:
        m=re.search(re.escape(customer)+r'\s*[_\-/／|:]\s*([^\n\r,;|]+)',issue,re.I)
        if m:return N(m.group(1)).strip(' _-/／|:：')
    return N(d.get('task_name'))

def extract(path):
    d=impl.extract(path)
    try:
        t=task_from_customer_issue(path,d)
        if t:d['task_name']=t
    except Exception:pass
    try:d['_section_images']=collect_section_images(path)
    except Exception:d['_section_images']={k:[] for k in ('2D','3D','4D','5D','6D')}
    return d
base.extract=extract

# ---------- source image collection ----------
def walk(container):
    for sh in getattr(container,'shapes',[]):
        yield sh
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP: yield from walk(sh)

def box(sh): return (float(sh.left),float(sh.top),float(sh.width),float(sh.height))
def area_overlap(a,b):
    ix=max(0,min(a[0]+a[2],b[0]+b[2])-max(a[0],b[0])); iy=max(0,min(a[1]+a[3],b[1]+b[3])-max(a[1],b[1])); return ix*iy

def sec_of(t):
    q=C(t)
    if re.search(r'2d',q) or any(k in q for k in ('문제현상','문제현황','불량현상')):return '2D'
    if re.search(r'3d',q) or any(k in q for k in ('임시조치','임시대책','고객대응')):return '3D'
    if re.search(r'4d',q) or any(k in q for k in ('원인분석','발생원인','유출원인','시스템원인')):return '4D'
    if re.search(r'5d',q) or any(k in q for k in ('개선대책','개선사항','영구개선')):return '5D'
    if re.search(r'6d',q) or any(k in q for k in ('효과검증','유효성검증','효과성검증')):return '6D'
    return None

def image_blob(sh):
    try:
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:return sh.image.blob
        pics=[p for p in walk(sh) if getattr(p,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE]
        if not pics:return None
        minx=min(box(p)[0] for p in pics); miny=min(box(p)[1] for p in pics)
        maxx=max(box(p)[0]+box(p)[2] for p in pics); maxy=max(box(p)[1]+box(p)[3] for p in pics)
        W=max(100,int((maxx-minx)/EMU*160)); H=max(80,int((maxy-miny)/EMU*160))
        canvas=PILImage.new('RGB',(W,H),'white')
        for p in pics:
            with PILImage.open(io.BytesIO(p.image.blob)) as im:
                im=im.convert('RGB'); px,py,pw,ph=box(p); w=max(1,int(pw/EMU*160)); h=max(1,int(ph/EMU*160)); im.thumbnail((w,h),PILImage.Resampling.LANCZOS); canvas.paste(im,(max(0,int((px-minx)/EMU*160)),max(0,int((py-miny)/EMU*160))))
        b=io.BytesIO();canvas.save(b,'PNG');return b.getvalue()
    except Exception:return None

def collect_section_images(path):
    prs=Presentation(path); out={k:[] for k in ('2D','3D','4D','5D','6D')}
    for si,sl in enumerate(prs.slides):
        anchors=[]
        for sh in walk(sl):
            t=N(getattr(sh,'text','')); sec=sec_of(t)
            if sec: anchors.append((sec,box(sh)))
        if not anchors: continue
        for sh in sl.shapes:
            if getattr(sh,'shape_type',None) not in (MSO_SHAPE_TYPE.PICTURE,MSO_SHAPE_TYPE.GROUP):continue
            blob=image_blob(sh)
            if not blob:continue
            ib=box(sh); candidates=[]
            for sec,ab in anchors:
                ov=area_overlap(ib,ab); cx=ib[0]+ib[2]/2; cy=ib[1]+ib[3]/2; ax=ab[0]+ab[2]/2; ay=ab[1]+ab[3]/2; dist=((cx-ax)**2+(cy-ay)**2)**0.5
                candidates.append((1 if ov>0 else 0,ov,-dist,sec))
            candidates.sort(reverse=True); sec=candidates[0][3]
            out[sec].append((blob,ib,si))
    # unique + largest first; retain several distinct source pictures for page 2.
    for sec,vals in out.items():
        uniq=[]; seen=set()
        for v in sorted(vals,key=lambda z:z[1][2]*z[1][3],reverse=True):
            h=hash(v[0])
            if h in seen:continue
            seen.add(h);uniq.append(v)
        out[sec]=uniq[:4]
    return out

def composite(items,cols=2):
    if not items:return None
    ims=[]
    for blob,_,_ in items:
        try:
            with PILImage.open(io.BytesIO(blob)) as im:
                im=im.convert('RGB'); im.thumbnail((700,420),PILImage.Resampling.LANCZOS); ims.append(im.copy())
        except Exception:pass
    if not ims:return None
    cellw=700; cellh=450; rows=(len(ims)+cols-1)//cols
    canvas=PILImage.new('RGB',(cols*cellw,rows*cellh),'white'); draw=ImageDraw.Draw(canvas)
    for i,im in enumerate(ims):
        x=(i%cols)*cellw; y=(i//cols)*cellh
        fitted=ImageOps.contain(im,(cellw-20,cellh-30)); canvas.paste(fitted,(x+(cellw-fitted.width)//2,y+10)); draw.rectangle((x,y,x+cellw-1,y+cellh-1),outline=(190,190,190),width=2)
    b=io.BytesIO();canvas.save(b,'PNG');return b.getvalue()

# ---------- weekly page 1: preserve title formatting, write only table cells ----------
def header_map(tb):
    m={}
    for c in range(len(tb.columns)):
        q=C(tb.cell(0,c).text)
        if '과제명' in q:m['task']=c
        if q=='이슈' or '이슈명' in q:m['issue']=c
        if '문제' in q or '현상' in q:m['problem']=c
        if '진행현황' in q or '진행' in q:m['progress']=c
        if 'signal' in q:m['signal']=c
    return m

def write_summary(prs,d):
    vals={'task':N(d.get('task_name')),'issue':N(d.get('issue_name')),'problem':N(d.get('problem'))}
    try: vals['progress']=impl.v305.v303.v301.impl.v29.v29_progress(d)
    except Exception: vals['progress']=''
    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False):continue
            tb=sh.table; hm=header_map(tb)
            if 'issue' not in hm or 'task' not in hm:continue
            row=1
            for r in range(1,len(tb.rows)):
                # choose first blank data row; otherwise first data row
                if not any(N(tb.cell(r,c).text) for c in range(len(tb.columns))):row=r;break
            for k,c in hm.items():
                if k in vals:set_cell(tb.cell(row,c),vals[k],8)
            if 'signal' in hm:base.signal(tb.cell(row,hm['signal']),base.status(d))
            return

def replace_title_preserve(sh,team):
    if not team or not hasattr(sh,'text_frame'):return
    # Preserve run-level formatting: replace text in-place when possible.
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if '팀 주요 논의 사항' in r.text:
                r.text=re.sub(r'[^\s]*팀\s*주요\s*논의\s*사항',team+'팀 주요 논의 사항',r.text)
                return
    t=N(getattr(sh,'text',''))
    if '팀 주요 논의 사항' in t:
        # only fallback when the title is a single run-free text box
        old=sh.text_frame.text; new=re.sub(r'[^\n]*팀\s*주요\s*논의\s*사항',team+'팀 주요 논의 사항',old)
        if new!=old:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:r.text=r.text.replace(old,new)

# ---------- page 2 ----------
def marker_info(sl,label):
    # Return marker and its immediate parent group, if any.
    for sh in sl.shapes:
        if N(getattr(sh,'text',''))==label:return sh,None
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
            for ch in walk(sh):
                if ch is sh:continue
                if N(getattr(ch,'text',''))==label:return ch,sh
    return None,None

def marker_bounds(sh,parent):
    if parent is not None:return box(parent)
    return box(sh)

def move_marker_unit(sl,label,x,y):
    sh,parent=marker_info(sl,label)
    if sh is None:return
    target_left=Inches(x); target_top=Inches(y)
    bx,by,_,_=marker_bounds(sh,parent); dx=target_left-Inches(bx/EMU); dy=target_top-Inches(by/EMU)
    try:
        if parent is not None:parent.left+=int(dx); parent.top+=int(dy)
        else: sh.left+=int(dx); sh.top+=int(dy)
    except Exception:pass

def remove_auto(sl):
    for sh in list(sl.shapes):
        if str(getattr(sh,'name','')).startswith('AUTO_8D_'):
            try:sh._element.getparent().remove(sh._element)
            except Exception:pass

def remove_all_pictures(sl):
    # Source/template pictures on detail page are removed so they cannot remain as
    # tiny scattered thumbnails. New controlled composites are added below.
    for sh in list(sl.shapes):
        if getattr(sh,'shape_type',None) in (MSO_SHAPE_TYPE.PICTURE,MSO_SHAPE_TYPE.GROUP):
            if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP and not any(getattr(x,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE for x in walk(sh)):continue
            try:sh._element.getparent().remove(sh._element)
            except Exception:pass

def add_image(sl,blob,x,y,w,h,name):
    try:
        with PILImage.open(io.BytesIO(blob)) as im: iw,ih=im.size
        scale=min(w*EMU/iw,h*EMU/ih); ww=iw*scale/EMU; hh=ih*scale/EMU
        p=sl.shapes.add_picture(io.BytesIO(blob),Inches(x+(w-ww)/2),Inches(y+(h-hh)/2),width=Inches(ww),height=Inches(hh));p.name='AUTO_8D_IMG_'+name
    except Exception:pass

def add_text(sl,name,text,x,y,w,h,size=8):
    sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));sh.name='AUTO_8D_'+name;set_shape(sh,text,size);return sh

def fill_zone(sl,zkey,z,text,items):
    x,y,w,h=z['x'],z['y'],z['w'],z['h']
    # Reserve right side for ONE composite containing several source pictures.
    blob=composite(items)
    if blob:
        iw=min(1.75,w*.38); txw=w-iw-.08; add_text(sl,'TEXT_'+zkey,text,x,y,txw,h,8); add_image(sl,blob,x+txw+.08,y,iw,h,'COLLAGE_'+zkey)
    else:add_text(sl,'TEXT_'+zkey,text,x,y,w,h,8)

def detail_layout(): return copy.deepcopy(getattr(renderer,'DEFAULT_LAYOUT',{
    '2D':{'x':.53,'y':2.53,'w':4.76,'h':.98},'3D':{'x':.53,'y':3.91,'w':4.76,'h':1.14},'4D_CAUSE':{'x':.53,'y':5.38,'w':4.76,'h':1.79},'4D_LEAK':{'x':5.77,'y':2.47,'w':4.76,'h':.87},'5D':{'x':5.77,'y':3.74,'w':4.76,'h':1.77},'6D':{'x':5.77,'y':5.91,'w':4.78,'h':.94}}))

def is_new(mode):
    q=C(mode);return any(k in q for k in ('신규','new','신규이슈'))

def fill_detail(sl,d,g,mode):
    remove_auto(sl)
    layout=detail_layout()
    new=is_new(mode)
    # For a new issue, normalize the existing marker+title unit to the example positions.
    # For an existing issue, do not move markers at all: preserve the supplied template.
    if new:
        targets={'2D':(.35,2.18),'3D':(.35,3.53),'4D':(5.59,2.18),'5D':(5.59,3.45),'6D':(5.59,5.60)}
        for label,(x,y) in targets.items():move_marker_unit(sl,label,x,y)
    # Delete all old pictures/groups that contain pictures, preventing scatter from source/template.
    remove_all_pictures(sl)
    imgs=d.get('_section_images',{}) or {}
    fill_zone(sl,'2D',layout['2D'],N(d.get('problem')) or '검토 중',imgs.get('2D',[]))
    fill_zone(sl,'3D',layout['3D'],'\n'.join(x for x in (d.get('temporary_action'),d.get('customer_response')) if N(x)) or '검토 중',imgs.get('3D',[]))
    fill_zone(sl,'4D_CAUSE',layout['4D_CAUSE'],N(d.get('cause_4d')) or '검토 중',imgs.get('4D',[]))
    fill_zone(sl,'4D_LEAK',layout['4D_LEAK'],'\n'.join(x for x in (d.get('leak_cause'),d.get('system_cause')) if N(x)) or '검토 중',[])
    fill_zone(sl,'5D',layout['5D'],N(d.get('action_5d')) or '검토 중',imgs.get('5D',[]))
    fill_zone(sl,'6D',layout['6D'],N(d.get('verification_6d')) or '검토 중',imgs.get('6D',[]))
    # Page-2 metadata: write only target placeholders; do not globally reformat the template.
    team=N(g.get('team')); owner=N(g.get('owner')); issue=N(d.get('issue_name')); task=N(d.get('task_name'))
    for sh in list(sl.shapes):
        if not hasattr(sh,'text_frame'):continue
        t=N(getattr(sh,'text',''))
        if t.startswith('이슈명'):
            prefix=t.split(':',1)[0] if ':' in t else '이슈명'
            set_shape(sh,prefix+' : '+issue,8)
        elif '과제명_이슈 제목' in t:
            set_shape(sh,t.replace('과제명_이슈 제목',task),8)
        elif '00팀 담당자' in t:
            set_shape(sh,f'{team} 담당자 : {owner}',8)
        elif t.startswith('이슈명 :'):
            set_shape(sh,'이슈명 : '+issue,8)
        if getattr(sh,'has_table',False):
            tb=sh.table
            for r in range(len(tb.rows)):
                for c in range(len(tb.columns)):
                    q=C(tb.cell(r,c).text)
                    if 'signal'==q and c+1<len(tb.columns):
                        cell=tb.cell(r,c+1);base.signal(cell,base.status(d));
                        for p in cell.text_frame.paragraphs:
                            for rr in p.runs:set_run(rr,8)
                    if '발생단계' in q and c+1<len(tb.columns):set_cell(tb.cell(r,c+1),(N(g.get('stage'))+' ('+N(d.get('occurrence_date'))+')').strip(' ()'),9)

def weekly(src,out,d,g,mode):
    prs=Presentation(src);write_summary(prs,d)
    team=N(g.get('team'))
    # page-1 title: replace only the title run, never apply a global font change.
    for sl in prs.slides[:1]:
        for sh in walk(sl):replace_title_preserve(sh,team)
    # Page 2: exactly one detail slide. If the weekly template has more pages, process only
    # the first detail page so source-page images cannot create multiple scattered pages.
    if len(prs.slides)>1: fill_detail(prs.slides[1],d,g,mode)
    Path(out).parent.mkdir(exist_ok=True)
    try:prs.save(out);saved=out
    except PermissionError:
        p=Path(out);saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix);prs.save(saved)
    return '주간회의 PPT 업데이트: '+base.status(d),saved
base.weekly=weekly
if __name__=='__main__':base.App().mainloop()
