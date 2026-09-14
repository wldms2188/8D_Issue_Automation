# Recovery STEP4: restore full page1 problem and force exact page2 title by geometry.
# Preserve v3.2.0 image/layout/font/Excel behavior.
import datetime
from pathlib import Path
from pptx import Presentation

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
    except Exception:
        pass
    return True


def _force_page2_header(sl,d,g):
    customer=N(d.get('customer'))
    task=N(d.get('task_name'))
    kind,event=step3._event_name(d)
    prefix=f'{customer}_{task}'.strip('_')
    if kind=='시험' and event:
        title=f'{prefix}_{event} 이슈 발생'.strip('_')
    elif kind=='빌드' and event:
        title=f'{prefix}_{event} 빌드 이슈 발생'.strip('_')
    else:
        issue=v319._trim_before_customer(d.get('issue_name'),customer)
        title=f'{prefix}_{issue}'.strip('_')

    # First try the placeholder-aware finder, then always fall back to geometry.
    title_sh,issue_sh=v319._find_page2_header_shapes(sl)
    if title_sh is None:
        title_sh=_title_shape_by_geometry(sl)
    _set_title_exact(title_sh,title)

    if issue_sh is not None:
        issue=v319._trim_before_customer(d.get('issue_name'),customer)
        v319._set_issue_line(issue_sh,issue)
    step1._update_team_owner(sl,g)


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
