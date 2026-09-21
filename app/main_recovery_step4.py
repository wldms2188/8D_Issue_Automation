# Recovery STEP4: restore full page1 problem and force exact page2 title by geometry.
# Preserve v3.2.0 image/layout/font/Excel behavior.
import datetime
import re
from pathlib import Path
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

import main_recovery_step3 as step3
import main_recovery_step1 as step1
import main_v320 as v320
import main_v319 as v319
import main_v310 as v310

base=step3.base
N=v310.N
EMU=v310.EMU


def _full_problem(d):
    return N(d.get('problem'))


def _title_shape_by_geometry(sl):
    """Find page2 title even after placeholder text was already overwritten.
    Prefer a text box in the upper-left header region and exclude the issue-name line.
    """
    candidates=[]
    for sh in v310.walk(sl):
        if not hasattr(sh,'text_frame'):
            continue
        t=N(getattr(sh,'text',''))
        try:
            x=float(sh.left)/EMU; y=float(sh.top)/EMU; w=float(sh.width)/EMU; h=float(sh.height)/EMU
        except Exception:
            continue
        if y>0.55 or x>4.0:
            continue
        q=t.replace(' ','').lower()
        if q.startswith('이슈명') or '담당자' in t or 'signal' in q:
            continue
        # actual template title is around x=.17 y=.11 w=2.11 h=.44
        score=abs(x-.17)+abs(y-.11)+.15*abs(w-2.11)+.15*abs(h-.44)
        candidates.append((score,sh))
    if not candidates:
        return None
    candidates.sort(key=lambda z:z[0])
    return candidates[0][1]


def _owner_shape(sl):
    candidates=[]
    for sh in v310.walk(sl):
        if not hasattr(sh,'text_frame'):
            continue
        t=N(getattr(sh,'text',''))
        if '담당자' not in t:
            continue
        try:
            x=float(sh.left)/EMU; y=float(sh.top)/EMU
        except Exception:
            continue
        if y<=0.80:
            candidates.append((-x,y,sh))
    if not candidates:
        return None
    candidates.sort(key=lambda z:(z[1],z[0]))
    return candidates[0][2]

def _estimated_text_width_in(sh,text):
    """Conservative one-line width estimate for the visible detail title."""
    pt=18.0
    try:
        runs=[r for p in sh.text_frame.paragraphs for r in p.runs]
        sizes=[r.font.size.pt for r in runs if r.font.size]
        if sizes:
            # Older generated pages may already have a shrunken font. Do not let
            # that hide a collision that would occur with the normal template title.
            pt=max(14.0,max(sizes))
    except Exception:
        pass
    em=pt/72.0
    units=0.0
    for ch in N(text):
        if ord(ch)>=0x2E80:
            units+=1.0
        elif ch.isspace():
            units+=0.35
        elif ch in '._-()/[]':
            units+=0.45
        else:
            units+=0.58
    return units*em

def _title_owner_overlap_risk(sl,title_sh,title_text):
    owner=_owner_shape(sl)
    if title_sh is None:
        return False
    try:
        tx=float(title_sh.left)/EMU; ty=float(title_sh.top)/EMU
        tw=float(title_sh.width)/EMU; th=float(title_sh.height)/EMU
    except Exception:
        return False

    estimated=_estimated_text_width_in(title_sh,title_text)
    if owner is None:
        # Even without a detected owner box, do not let a one-line title visibly
        # run out of its own template title area.
        return estimated>max(.25,tw*.98)

    try:
        ox=float(owner.left)/EMU; oy=float(owner.top)/EMU
        ow=float(owner.width)/EMU; oh=float(owner.height)/EMU
    except Exception:
        return estimated>max(.25,tw*.98)

    # Header boxes are sometimes a few pixels vertically offset even though the
    # rendered text is on the same visual line. Treat near-aligned boxes as colliding.
    tc=ty+th/2; oc=oy+oh/2
    same_header_line=abs(tc-oc)<=max(.35,(th+oh)*.75)
    if not same_header_line:
        return estimated>max(.25,tw*.98)

    available=max(.25,ox-tx-.12)
    rendered_right=tx+min(max(estimated,.01),max(tw,estimated))
    return estimated>available or rendered_right>ox-.08

def _strip_selected_project_prefix(issue,d):
    """Remove routing markers and visible customer/project prefix from detail title."""
    s=v319._trim_before_customer(issue,d.get('customer'))
    s=v319._clean_issue_label(s)
    if not s:
        return ''

    prefix=N(v319._weekly_task(d))
    if prefix:
        m=re.match(r'^\s*'+re.escape(prefix)+r'\s*[_\-/／|:：]*\s*',s,re.I)
        if m:
            rest=s[m.end():].strip(' _-/／|:：')
            if rest:
                return rest

    customer=N(d.get('customer'))
    if customer:
        # Old generated labels can contain a project name different from the GUI
        # selection. Remove the visible CUSTOMER_<project> prefix rather than
        # leaving it in a collision fallback title.
        m=re.match(
            r'^\s*'+re.escape(customer)+r'\s*[_\-/／|:：]\s*'
            r'([^_／|:：]+)\s*[_／|:：]\s*(.+)$',
            s,re.I
        )
        if m and N(m.group(2)):
            return N(m.group(2)).strip(' _-/／|:：')

    return s


def _detail_title_suffix(d,g):
    """V1-style title name + '이슈 발생', extended to current origin/site inputs.

    Priority:
      1) specific test/build name found in the 8D,
      2) explicit 시험 occurrence site,
      3) confirmed issue-origin category (부품/설계/공정/기타),
      4) occurrence-site fallback,
      5) cleaned issue name.
    """
    kind,event=step3._event_name(d)
    if kind=='시험' and event:
        return f'{event} 이슈 발생'
    if kind=='빌드' and event:
        return f'{event} 이슈 발생'

    site=N((g or {}).get('occurrence_site') or d.get('occurrence_site'))
    if '시험' in site:
        return '시험 이슈 발생'

    origin=N((g or {}).get('_issue_origin_selected') or d.get('_issue_origin_selected'))
    if origin in ('부품','설계','공정','기타'):
        return f'{origin} 이슈 발생'

    if '부품' in site:
        return '부품 이슈 발생'
    if '생산' in site or '공정' in site:
        return '공정 이슈 발생'

    issue=_strip_selected_project_prefix(d.get('issue_name'),d)
    q=N(issue)
    for label in ('시험','공정','부품','설계'):
        if label in q:
            return f'{label} 이슈 발생'
    if not q:
        return '이슈 발생'
    if '이슈 발생' in q:
        return q
    if q.endswith('발생'):
        return q
    return q+' 이슈 발생'


def _title_without_selected_project(d,g):
    """Collision fallback: category/event title only, never 고객사_과제명."""
    return _detail_title_suffix(d,g)


def _set_title_exact(sh,text):
    if sh is None or not hasattr(sh,'text_frame'):
        return False
    tf=sh.text_frame
    runs=[r for p in tf.paragraphs for r in p.runs]
    if runs:
        runs[0].text=N(text)
        for r in runs[1:]:
            r.text=''
    else:
        tf.text=N(text)
    try:
        tf.word_wrap=False
        tf.auto_size=MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass
    return True


def _force_page2_header(sl,d,g):
    customer=N(d.get('customer'))
    prefix=v319._weekly_task(d)
    suffix=_detail_title_suffix(d,g)
    title=f'{prefix}_{suffix}'.strip('_') if prefix else suffix

    # First try the placeholder-aware finder, then always fall back to geometry.
    title_sh,issue_sh=v319._find_page2_header_shapes(sl)
    if title_sh is None:
        title_sh=_title_shape_by_geometry(sl)

    # Finalize owner first. If the full title can collide, remove 고객사_과제명
    # completely and keep only '<시험명/공정/부품/...> 이슈 발생'.
    step1._update_team_owner(sl,g)
    if _title_owner_overlap_risk(sl,title_sh,title):
        title=_title_without_selected_project(d,g)
    _set_title_exact(title_sh,title)

    if issue_sh is not None:
        issue=v319._trim_before_customer(d.get('issue_name'),customer)
        v319._set_issue_line(issue_sh,issue)


def weekly(src,out,d,g,mode):
    dd=v319._weekly_display(d)
    p1=dict(dd)
    p1['issue_name']=step1._page1_issue(d)
    p1['problem']=_full_problem(d)

    prs=Presentation(src)
    v319._update_page1(prs,p1,g)

    if len(prs.slides)>1:
        v320._update_page2(prs.slides[1],dd,g,mode)
        _force_page2_header(prs.slides[1],d,g)

    Path(out).parent.mkdir(parents=True,exist_ok=True)
    try:
        prs.save(out); saved=out
    except PermissionError:
        p=Path(out)
        saved=str(p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix))
        prs.save(saved)
    return '주간회의 PPT 업데이트: '+base.status(d),saved

base.weekly=weekly

class RecoveryStep4App(base.App):
    def __init__(self):
        super().__init__(); self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP4')

if __name__=='__main__':
    RecoveryStep4App().mainloop()
