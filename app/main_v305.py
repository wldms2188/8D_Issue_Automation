# 8D Issue Automation v3.0.5
# Fixes weekly mapping based on the user's latest feedback.
# - Summary page maps values by the EXISTING column headers (Issue gets issue_name,
#   Task gets task_name), never by hard-coded column positions.
# - Task name is taken from the issue-name/customer area; when issue_name is in the
#   common "고객사/과제명" form, the part after 고객사 is the task.
# - Weekly page 2 keeps the template's existing circular D markers and moves them
#   with their content; no new D labels are created.
# - Each D section is treated as one movable unit. Existing marker + generated content
#   are grouped in PowerPoint XML after rendering.
# - Source images are filtered to one meaningful, largest image per D section to stop
#   many tiny images from different source pages being scattered over page 2.
# - 8 pt is retained unless the final zone truly cannot contain the text; minimum 7 pt.

import sys, re, copy
from pathlib import Path
APP_DIR=Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path: sys.path.insert(0,str(APP_DIR))
import main_v303 as v303
base=v303.base
renderer=v303.renderer
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

# ---------------------------------------------------------------------------
# Task-name correction
# ---------------------------------------------------------------------------
def _n(x): return str(x or '').replace('\r\n','\n').replace('\r','\n').strip()
def _c(x):
    try: return base.compact(x)
    except Exception: return re.sub(r'[^0-9A-Za-z가-힣]','',_n(x)).lower()

def _clean_task(x):
    s=_n(x)
    # Remove a leading metadata label when present.
    s=re.sub(r'^과제명\s*[:：]\s*','',s,flags=re.I)
    return s.strip(' /／|:：')

def _task_from_issue_name(d):
    issue=_n(d.get('issue_name'))
    if not issue: return ''
    # The agreed source is the issue-name area: 고객사 is before the task.
    # Handle both "고객사/과제명" and "고객사 / 과제명" forms.
    m=re.search(r'고객사\s*[/／|>→:-]\s*(.+)$',issue)
    if m:
        return _clean_task(m.group(1))
    parts=re.split(r'[/／|]',issue,maxsplit=1)
    if len(parts)==2 and _n(parts[1]):
        return _clean_task(parts[1])
    return ''

def _task_from_ppt(path):
    # First try the issue-name value already extracted by the stable extractor.
    try:
        d=base._STABLE_EXTRACT(path) if hasattr(base,'_STABLE_EXTRACT') else None
        t=_task_from_issue_name(d or {})
        if t: return t
    except Exception: pass
    # Direct shape/table search for the explicit 고객사 -> 과제명 relationship.
    prs=Presentation(path)
    for sl in prs.slides:
        for sh in sl.shapes:
            if getattr(sh,'has_table',False):
                tb=sh.table
                rows=[[_n(tb.cell(r,c).text) for c in range(len(tb.columns))] for r in range(len(tb.rows))]
                for r,row in enumerate(rows):
                    for c,val in enumerate(row):
                        if '고객사' not in _c(val): continue
                        # Same-row cells after 고객사: skip customer value, then use task.
                        vals=[_n(x) for x in row[c+1:] if _n(x)]
                        for v in vals:
                            if '과제명' in _c(v):
                                continue
                            if len(vals)>=2 and v==vals[0]:
                                continue
                            if not any(k in _c(v) for k in ('이슈명','model','packer','발생site','발생일자','lotno','발생line','담당자')):
                                return _clean_task(v)
                        # Explicit label/value below.
                        for rr in range(r+1,min(r+4,len(rows))):
                            for cc,x in enumerate(rows[rr]):
                                if '과제명' in _c(x) and cc+1<len(rows[rr]):
                                    return _clean_task(rows[rr][cc+1])
    return ''

_prev_extract=base.extract
def extract_v305(path):
    d=_prev_extract(path)
    try:
        t=_task_from_issue_name(d)
        if not t: t=_task_from_ppt(path)
        if t: d['task_name']=t
    except Exception: pass
    return d
base.extract=extract_v305

# ---------------------------------------------------------------------------
# Keep only one useful image per D section. This prevents a page full of tiny
# thumbnails when the source 8D has many pages/images.
# ---------------------------------------------------------------------------
def _filter_section_images(d):
    imgs=d.get('_section_images',{}) or {}
    out={k:[] for k in ('2D','3D','4D','5D','6D')}
    for sec,vals in imgs.items():
        if not vals: continue
        # Prefer the largest source image; it is the most useful visual for the
        # compact weekly section. Keep only one per D section.
        vals=list(vals)
        vals.sort(key=lambda item: (item[1][2]*item[1][3], item[2]), reverse=True)
        out[sec]=[vals[0]]
    d['_section_images']=out
    return d

_old_extract=base.extract
def extract_images_v305(path):
    d=_old_extract(path)
    try: _filter_section_images(d)
    except Exception: pass
    return d
base.extract=extract_images_v305

# ---------------------------------------------------------------------------
# Existing marker movement
# ---------------------------------------------------------------------------
def _walk(container):
    for sh in getattr(container,'shapes',[]):
        yield sh
        if hasattr(sh,'shapes'):
            yield from _walk(sh)

def _move_markers(sl,layout):
    # 4D marker belongs to the right-side 4D leak/system block in the supplied
    # example; the other markers belong to their corresponding zones.
    mapping=[('2D','2D'),('3D','3D'),('4D','4D_LEAK'),('5D','5D'),('6D','6D')]
    used=set()
    for label,zkey in mapping:
        z=layout.get(zkey)
        if not z: continue
        candidates=[]
        for sh in _walk(sl):
            if _n(getattr(sh,'text',''))==label and not str(getattr(sh,'name','')).startswith('AUTO_8D_'):
                candidates.append(sh)
        target=next((x for x in candidates if id(x) not in used),None)
        if target is None: continue
        used.add(id(target))
        try:
            tw=target.width/914400.0; th=target.height/914400.0
            target.left=Inches(max(.03,z['x']-tw-.08))
            target.top=Inches(z['y']+max(0,(z['h']-th)/2))
        except Exception: pass

# ---------------------------------------------------------------------------
# Remove generated D labels and create real PowerPoint groups.
# ---------------------------------------------------------------------------
def _remove_generated_labels(sl):
    for sh in list(sl.shapes):
        if str(getattr(sh,'name','')).startswith('AUTO_8D_LABEL_'):
            el=sh._element; el.getparent().remove(el)

def _direct_shapes(sl):
    return list(sl.shapes)

def _make_group(sl, shapes, name):
    shapes=[s for s in shapes if s is not None and getattr(s,'_element',None) is not None]
    if len(shapes)<2: return None
    # python-pptx has no public grouping API. Build a standard PowerPoint group
    # shape using the existing shape XML, preserving every child shape.
    from pptx.oxml.xmlchemy import OxmlElement
    spTree=sl.shapes._spTree
    grp=OxmlElement('p:grpSp')
    nv=OxmlElement('p:nvGrpSpPr'); cNvPr=OxmlElement('p:cNvPr'); cNvPr.set('id',str(max([int(x.get('id')) for x in sl.shapes if x.get('id')] or [1])+1000)); cNvPr.set('name',name)
    cNvGrpSpPr=OxmlElement('p:cNvGrpSpPr'); nvPr=OxmlElement('p:nvPr')
    nv.append(cNvPr); nv.append(cNvGrpSpPr); nv.append(nvPr); grp.append(nv)
    xfrm=OxmlElement('p:grpSpPr'); a_xfrm=OxmlElement('a:xfrm')
    off=OxmlElement('a:off'); off.set('x','0'); off.set('y','0'); ext=OxmlElement('a:ext'); ext.set('cx','0'); ext.set('cy','0')
    chOff=OxmlElement('a:chOff'); chOff.set('x','0'); chOff.set('y','0'); chExt=OxmlElement('a:chExt'); chExt.set('cx','0'); chExt.set('cy','0')
    a_xfrm.extend([off,ext,chOff,chExt]); xfrm.append(a_xfrm); grp.append(xfrm)
    # Move selected shape XML nodes into the group, preserving order.
    first=min([list(spTree).index(s._element) for s in shapes])
    spTree.insert(first,grp)
    for s in shapes:
        if s._element.getparent() is spTree:
            spTree.remove(s._element)
        grp.append(s._element)
    return grp

def _group_sections(sl):
    # Group each marker with the generated text/image shapes belonging to it.
    # 4D uses both cause and leak/system content as one D-level unit.
    mapping={
        '2D':['TEXT_2D','IMG_2D'],
        '3D':['TEXT_3D','IMG_3D'],
        '4D':['TEXT_4D_CAUSE','IMG_4D_CAUSE','TEXT_4D_LEAK','IMG_4D_LEAK'],
        '5D':['TEXT_5D','IMG_5D'],
        '6D':['TEXT_6D','IMG_6D'],
    }
    marker_by_label={}
    for sh in _walk(sl):
        txt=_n(getattr(sh,'text',''))
        if txt in ('2D','3D','4D','5D','6D') and not str(getattr(sh,'name','')).startswith('AUTO_8D_'):
            marker_by_label.setdefault(txt,sh)
    for label,prefixes in mapping.items():
        shapes=[]
        m=marker_by_label.get(label)
        if m: shapes.append(m)
        for sh in _direct_shapes(sl):
            nm=str(getattr(sh,'name',''))
            if any(p in nm for p in prefixes) or (label=='4D' and ('TEXT_4D_' in nm or 'IMG_4D_' in nm)):
                shapes.append(sh)
        _make_group(sl,shapes,'8D_GROUP_'+label)

# Wrap renderer. v303 supplies the collision-aware final layout.
_orig_fill=renderer._fill_page2
def _fill_page2_v305(sl,d,g,layout):
    _orig_fill(sl,d,g,layout)
    _remove_generated_labels(sl)
    _move_markers(sl,layout)
    _group_sections(sl)
renderer._fill_page2=_fill_page2_v305

# ---------------------------------------------------------------------------
# Summary table: map by header names, not by column order.
# ---------------------------------------------------------------------------
def _header_map(table):
    m={}
    for c in range(len(table.columns)):
        q=_c(table.cell(0,c).text)
        if '과제명' in q: m['task']=c
        elif '이슈' in q: m['issue']=c
        elif '문제' in q or '현상' in q: m['problem']=c
        elif '진행' in q: m['progress']=c
        elif 'signal' in q: m['signal']=c
    return m

def _set_summary(prs,d):
    title=_n(d.get('task_name')); issue=_n(d.get('issue_name')); problem=_n(d.get('problem')); progress=v303.v301.v301.v29.v29_progress(d) if hasattr(v303,'v301') else ''
    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False): continue
            tb=sh.table
            hm=_header_map(tb)
            if not hm.get('issue') is None and ('task' in hm or 'issue' in hm):
                row=None
                for r in range(1,len(tb.rows)):
                    vals=[_n(tb.cell(r,c).text) for c in range(len(tb.columns))]
                    if not any(vals): row=r; break
                if row is None: row=1 if len(tb.rows)>1 else 0
                data={'task':title,'issue':issue,'problem':problem,'progress':progress}
                for key,val in data.items():
                    if key in hm:
                        tb.cell(row,hm[key]).text=val
                if 'signal' in hm:
                    base.signal(tb.cell(row,hm['signal']),base.status(d))
                return

_orig_weekly=v303.renderer.weekly_v300 if hasattr(v303.renderer,'weekly_v300') else None

def weekly_v305(src,out,d,g,mode):
    prs=Presentation(src)
    _set_summary(prs,d)
    layout=renderer._load_layout()
    st=base.status(d)
    # Find the summary slide, then render all following detail slides once.
    summary_idx=0
    for i,sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if getattr(sh,'has_table',False) and _header_map(sh.table).get('issue') is not None:
                summary_idx=i; break
        if summary_idx==i and i>0: break
    for sl in [prs.slides[i] for i in range(summary_idx+1,len(prs.slides))]:
        renderer._fill_page2(sl,d,g,layout)
    Path(out).parent.mkdir(exist_ok=True)
    try: prs.save(out); saved=out
    except PermissionError:
        import datetime
        p=Path(out); saved=p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix); prs.save(saved)
    return '주간회의 PPT 업데이트: '+st,saved

base.weekly=weekly_v305
if __name__=='__main__': base.App().mainloop()
