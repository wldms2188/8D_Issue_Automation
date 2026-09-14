# 8D Issue Automation v2.9.5 hotfix
# Removes the broken main_v293 import dependency. Loads the known-good v2.9.3 module
# explicitly from the same app directory, then applies page-1-only image selection.
import sys, io, datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# v2.9.3 is the known-good implementation in this repository.
# Import it by module name after explicitly putting app/ on sys.path.
import main_v293

base = main_v293.base
v29 = main_v293.v29

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image as PILImage

EMU = 914400


def _pics(sh):
    out=[]
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:
        out.append(sh)
    elif getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
        for c in sh.shapes:
            out.extend(_pics(c))
    return out


def _box(sh):
    return float(sh.left),float(sh.top),float(sh.width),float(sh.height)


def _overlap(a,b):
    ax,ay,aw,ah=a; bx,by,bw,bh=b
    return max(0,min(ax+aw,bx+bw)-max(ax,bx))*max(0,min(ay+ah,by+bh)-max(ay,by))


def _group_image(group):
    pics=_pics(group)
    if len(pics)<2:
        return None
    gx,gy,gw,gh=_box(group)
    W=max(1,int(gw/EMU*120)); H=max(1,int(gh/EMU*120))
    canvas=PILImage.new('RGB',(W,H),'white')
    for p in pics:
        try:
            with PILImage.open(io.BytesIO(p.image.blob)) as im:
                im=im.convert('RGB')
                px,py,pw,ph=_box(p)
                x=max(0,int((px-gx)/EMU*120)); y=max(0,int((py-gy)/EMU*120))
                w=max(1,int(pw/EMU*120)); h=max(1,int(ph/EMU*120))
                im.thumbnail((w,h),PILImage.Resampling.LANCZOS)
                canvas.paste(im,(x,y))
        except Exception:
            pass
    b=io.BytesIO(); canvas.save(b,'PNG'); return b.getvalue()


def _page1_rep_image(ppt_path):
    """Representative image may ONLY come from page 1 of the 8D source."""
    prs=Presentation(ppt_path)
    if not prs.slides:
        return None
    sl=prs.slides[0]  # absolute page-1 boundary

    # Find semantic 2D/problem/defect regions on page 1.
    regions=[]
    for sh in sl.shapes:
        text=str(getattr(sh,'text','') or '').lower().replace(' ','')
        if any(k in text for k in ('2d','문제현황','문제현상','불량현상')):
            regions.append(_box(sh))

    candidates=[]
    for sh in sl.shapes:
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
            pics=_pics(sh)
            if len(pics)>=2:
                blob=_group_image(sh)
                if blob:
                    candidates.append((_box(sh),blob,True,len(pics)))
        elif getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:
            try:
                with PILImage.open(io.BytesIO(sh.image.blob)) as im:
                    if im.width*im.height>=10000:
                        candidates.append((_box(sh),sh.image.blob,False,1))
            except Exception:
                pass

    # 2D grouped pictures are absolute priority.
    hits=[]
    for box,blob,group,n in candidates:
        ov=max((_overlap(box,r) for r in regions),default=0)
        if ov>0:
            hits.append((group,n,ov,blob))
    if hits:
        hits.sort(key=lambda x:(x[0],x[1],x[2]),reverse=True)
        return hits[0][3]

    # If page 1 has no detectable 2D label, use a grouped image on page 1 only.
    groups=[x for x in candidates if x[2]]
    if groups:
        return max(groups,key=lambda x:x[3])[1]

    # Never fall through to page 2+.
    return None


# Preserve all existing v2.9.3 extraction/Excel/weekly behavior except the image source.
_original_extract=base.extract


def extract_v295(ppt_path):
    d=_original_extract(ppt_path)
    try:
        blob=_page1_rep_image(ppt_path)
        d['_images']=[(1,1,1,blob)] if blob else []
    except Exception:
        d['_images']=[]
    return d

base.extract=extract_v295

if __name__=='__main__':
    base.App().mainloop()
