# 8D Issue Automation v2.9.4
# Hard-fix for v2.9.3 layout/image failures.
# - Representative image: prefer the COMPLETE 2D picture group; never choose an unrelated page/image.
# - Weekly detail: use exact reference zones from supplied example; do not reuse a misplaced old textbox.
# - Marker labels are found recursively, even when grouped/embedded in tables.
# - Existing generated detail boxes are removed/replaced so 2D cannot overlap 3D.
# - Font fitting is measured from actual text-frame geometry; final 8pt pass no longer overwrites the fitted size.

import io, datetime, math
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from PIL import Image as PILImage
import main_v293 as v293
v292=v293.v292; v29=v293.v29; base=v293.base
EMU=914400

DETAIL_BOXES={
 '1':(0.53,2.53,4.76,0.98,'problem'),
 '2':(0.53,3.91,4.76,1.14,'temporary_action'),
 '3-1':(0.53,5.38,4.76,1.79,'cause_4d'),
 '3-2':(5.77,2.47,4.76,0.87,'leak_cause'),
 '4':(5.77,3.74,4.76,1.77,'action_5d'),
 '5':(5.77,5.91,4.78,0.94,'verification_6d'),
}
MARKER_TARGETS={
 '1D':(0.35,2.18),'2D':(0.35,3.53),'3D':(0.35,4.96),
 '4D':(5.59,2.34),'5D':(5.59,3.61),'6D':(5.59,5.78),
}

def _norm(s):
 return str(s or '').replace('\r\n','\n').replace('\r','\n').strip()

def _compact(s):
 try:return base.compact(_norm(s))
 except Exception:return ''.join(_norm(s).split()).lower()

def _font(run,size):
 run.font.name='맑은 고딕'; run.font.size=Pt(size)
 try:
  r=run._r.get_or_add_rPr(); r.set('a:latin','맑은 고딕'); r.set('a:ea','맑은 고딕'); r.set('a:cs','맑은 고딕')
 except Exception:pass

def _iter_shapes(container):
 for sh in getattr(container,'shapes',[]):
  yield sh
  if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
   yield from _iter_shapes(sh)

def _box(sh):
 return float(sh.left),float(sh.top),float(sh.width),float(sh.height)

def _set_box(sh,x,y,w,h):
 sh.left=Inches(x);sh.top=Inches(y);sh.width=Inches(w);sh.height=Inches(h)

def _find_marker(sl,label):
 target=_compact(label)
 candidates=[]
 for sh in _iter_shapes(sl):
  txt=_compact(getattr(sh,'text',''))
  if not txt:continue
  # Exact first; also accept marker embedded in a short label such as '1D\n문제 현황'.
  if txt==target or (len(txt)<24 and target in txt):
   x,y,w,h=_box(sh); candidates.append((len(txt),abs(x-Inches(MARKER_TARGETS[label][0]))+abs(y-Inches(MARKER_TARGETS[label][1])),sh))
 if not candidates:return None
 candidates.sort(key=lambda z:(z[0],z[1]))
 return candidates[0][2]

def _normalize_markers(sl):
 for lab,(x,y) in MARKER_TARGETS.items():
  sh=_find_marker(sl,lab)
  if sh is not None:_set_box(sh,x,y,float(sh.width)/EMU,float(sh.height)/EMU)

def _blank_generated(sl):
 # v2.9.x generated boxes are explicitly named; remove them so old/misplaced text cannot remain.
 doomed=[]
 for sh in list(sl.shapes):
  if str(getattr(sh,'name','')).startswith('AUTO_DETAIL_'):
   doomed.append(sh)
 for sh in doomed:
  sp=sh._element
  sp.getparent().remove(sp)

def _add_exact_box(sl,label,x,y,w,h):
 sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
 sh.name='AUTO_DETAIL_'+label+'_V294'
 tf=sh.text_frame; tf.clear(); tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.TOP
 tf.margin_left=Inches(.05);tf.margin_right=Inches(.05);tf.margin_top=Inches(.03);tf.margin_bottom=Inches(.03)
 return sh

def _estimate_height(sh,text,size):
 text=_norm(text)
 if not text:return 0.0
 # Conservative width estimate: Korean/Latin mixed text is treated as ~1 character per 0.095in at 8pt.
 width=max(0.2,float(sh.width)/EMU-0.10)
 chars=max(5,int(width/(0.095*(size/8.0))))
 lines=0
 for line in text.split('\n'):
  lines+=max(1,math.ceil(len(line)/chars))
 return lines*(size/72.0*1.22)+0.08

def _fit(sh,text,max_size=8,min_size=4):
 text=_norm(text) or '검토 중'
 sh.text_frame.clear(); sh.text_frame.word_wrap=True; sh.text_frame.vertical_anchor=MSO_ANCHOR.TOP
 size=float(max_size)
 avail=max(0.05,float(sh.height)/EMU-0.06)
 while size>min_size and _estimate_height(sh,text,size)>avail:
  size-=0.5
 # Write once at the final calculated size. Never run a later blanket 8pt pass.
 p=sh.text_frame.paragraphs[0]; p.text=text
 for pp in sh.text_frame.paragraphs:
  for r in pp.runs:_font(r,size)
 return size

def _overlap(a,b):
 ax,ay,aw,ah=a;bx,by,bw,bh=b
 return max(0,min(ax+aw,bx+bw)-max(ax,bx))*max(0,min(ay+ah,by+bh)-max(ay,by))

def _text_regions(sl):
 out=[]
 for sh in _iter_shapes(sl):
  txt=_compact(getattr(sh,'text',''))
  if any(k in txt for k in ['2d','문제현황','문제현상','불량현상','4d','원인분석','발생원인','유출원인']):
   out.append((*_box(sh),txt))
 return out

def _children_pics(sh):
 out=[]
 if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:return [sh]
 if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
  for c in sh.shapes:out.extend(_children_pics(c))
 return out

def _group_blob(group):
 pics=_children_pics(group)
 if len(pics)<2:return None
 gx,gy,gw,gh=_box(group)
 if gw<=0 or gh<=0:return None
 W=max(1,int(gw/EMU*160));H=max(1,int(gh/EMU*160));canvas=PILImage.new('RGB',(W,H),'white')
 for p in pics:
  try:
   with PILImage.open(io.BytesIO(p.image.blob)) as im:
    im=im.convert('RGB');px,py,pw,ph=_box(p)
    x=max(0,int((px-gx)/EMU*160));y=max(0,int((py-gy)/EMU*160));w=max(1,int(pw/EMU*160));h=max(1,int(ph/EMU*160))
    im.thumbnail((w,h),PILImage.Resampling.LANCZOS);canvas.paste(im,(x,y))
  except Exception:pass
 b=io.BytesIO();canvas.save(b,'PNG');return b.getvalue()

def _visuals(prs):
 out=[]
 for si,sl in enumerate(prs.slides):
  for sh in sl.shapes:
   if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
    pics=_children_pics(sh)
    if len(pics)>=2:
     blob=_group_blob(sh)
     if blob:out.append((si,*_box(sh),blob,True,len(pics)))
   elif getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:
    try:
     with PILImage.open(io.BytesIO(sh.image.blob)) as im:
      if im.width*im.height>=10000:out.append((si,*_box(sh),sh.image.blob,False,1))
    except Exception:pass
 return out

def _pick_rep(prs,d):
 vs=_visuals(prs)
 if not vs:return None
 def choose(keys):
  cand=[]
  for si,sl in enumerate(prs.slides):
   regs=[r for r in _text_regions(sl) if any(k in r[4] for k in keys)]
   for vsi,x,y,w,h,blob,grp,n in vs:
    if vsi!=si:continue
    ov=max((_overlap((x,y,w,h),r[:4]) for r in regs),default=0)
    if ov>0:cand.append((1 if grp else 0,ov,n,w*h,blob))
  if cand:
   cand.sort(key=lambda z:(z[0],z[1],z[2],z[3]),reverse=True);return cand[0][4]
  return None
 # 2D is mandatory preference. Only if no 2D visual exists, try 4D.
 return choose(['2d','문제현황','문제현상','불량현상']) or choose(['4d','원인분석','발생원인','유출원인'])

def extract_v294(path):
 d=v292.extract_v292(path)
 try:
  b=_pick_rep(Presentation(path),d)
  if b:d['_images']=[(1,1,1,b)]
 except Exception:pass
 return d
base.extract=extract_v294

def fill_detail(sl,d,g):
 _blank_generated(sl)
 _normalize_markers(sl)
 for label,(x,y,w,h,key) in DETAIL_BOXES.items():
  sh=_add_exact_box(sl,label,x,y,w,h)
  if key=='leak_cause':text='\n'.join(_norm(v) for v in [d.get('leak_cause'),d.get('system_cause')] if _norm(v))
  else:text=_norm(d.get(key))
  _fit(sh,text,8,4)
 for sh in _iter_shapes(sl):
  t=_norm(getattr(sh,'text',''))
  if t.startswith('과제명_이슈 제목'):_fit(sh,base.task(d),8,4)
  elif t.startswith('이슈명 :'):_fit(sh,'이슈명 : '+_norm(d.get('issue_name')),8,4)
  elif '00팀 담당자' in t:_fit(sh,f"{g.get('team','')} 담당자 : {g.get('owner','')}",8,4)
  if getattr(sh,'has_table',False):
   tb=sh.table
   for r in range(len(tb.rows)):
    for c in range(len(tb.columns)):
     q=base.norm(tb.cell(r,c).text)
     if q=='signal' and c+1<len(tb.columns):base.signal(tb.cell(r,c+1),base.status(d))

def weekly_v294(src,out,d,g,mode):
 prs=Presentation(src);st=base.status(d);title=base.task(d);issue=_norm(d.get('issue_name'));summary=None;idx=0
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
  vals=[title,issue,_norm(d.get('problem')),v29.v29_progress(d),'●']
  for c,val in enumerate(vals):v29.v29_set_cell_text(summary.cell(r,c),val,8)
  base.signal(summary.cell(r,4),st)
 for sl in [prs.slides[i] for i in range(idx+1,len(prs.slides))]:fill_detail(sl,d,g)
 Path(out).parent.mkdir(exist_ok=True)
 try:prs.save(out);saved=out
 except PermissionError:
  p=Path(out);saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix);prs.save(saved)
 return '주간회의 PPT 업데이트: '+st,saved
base.weekly=weekly_v294
if __name__=='__main__':base.App().mainloop()
