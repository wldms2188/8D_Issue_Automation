import datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

import main_recovery_step11 as step11
import main_recovery_step10 as step10
import main_recovery_step8 as step8
import main_recovery_step4 as step4
import main_v319 as v319
import main_v310 as v310

base=step11.base
N=v310.N
C=v310.C


def _signal_status_from_g(d,g):
    """Use the independently confirmed weekly-meeting Signal when available."""
    weekly=N(g.get('_weekly_status_selected'))
    if weekly in ('원인/개선 미확인','개선 검증중','개선 완료'):
        return weekly

    # Backward-compatible fallback for older callers that only supply Issue DB status.
    selected=N(g.get('_issue_status_selected')).lower()
    if selected=='close':
        return '개선 완료'
    if selected=='open':
        return '개선 검증중' if N(d.get('action_5d')) or N(d.get('verification_6d')) else '원인/개선 미확인'
    return base.status(d)


def _sync_page2_signal_only(prs,d,g):
    """Update ONLY page-2 Signal.

    Page 1 is already updated by v319._update_page1(), which knows the correct
    issue row. Rewriting page 1 here is unsafe for existing issues because the
    active issue may not be the row immediately below the header; doing so can
    overwrite the 유첨/비고 cell. Page 2 must simply mirror the same status.
    """
    st=_signal_status_from_g(d,g)
    if len(prs.slides)<2:
        return st

    sl=prs.slides[1]
    for sh in v310.walk(sl):
        if not getattr(sh,'has_table',False):
            continue
        tb=sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                if C(tb.cell(r,c).text)=='signal' and c+1<len(tb.columns):
                    base.signal(tb.cell(r,c+1),st)
    return st


def _labeled_4d_text(d,key):
    if key=='4D_CAUSE':
        body=N(d.get('cause_4d')) or '검토 중'
        return '- 발생원인\n'+body
    body='\n'.join(x for x in (N(d.get('leak_cause')),N(d.get('system_cause'))) if x) or '검토 중'
    return '- 유출원인\n'+body


def _update_page2_step12(sl,d,g,mode):
    v310.remove_previous_auto(sl)
    imgs=d.get('_section_images',{}) or {}
    english_mode=bool(d.get('_english_mode'))
    if english_mode:
        zones,texts,fonts=step11._english_template_anchored_layout(d,imgs)
    else:
        zones,texts,fonts=step11._layout_with_4d_placeholders(d,imgs)

    texts['4D_CAUSE']=_labeled_4d_text(d,'4D_CAUSE')
    texts['4D_LEAK']=_labeled_4d_text(d,'4D_LEAK')

    if v310.is_new(mode):
        for label,key in [('2D','2D'),('3D','3D'),('5D','5D'),('6D','6D')]:
            z=zones[key]
            v310.move_marker_unit(sl,label,max(.10,z['x']-.18),max(.10,z['y']-.35))

    for key in ('2D','3D','4D_CAUSE','4D_LEAK','5D','6D'):
        v319._render(sl,key,zones[key],texts[key],imgs.get(key,[]),fonts[key])

    if english_mode:
        # Prevent visual text overflow beyond the geometry even when the source
        # contains exceptionally long English sentences.
        for sh in sl.shapes:
            name=str(getattr(sh,'name',''))
            if name.startswith('AUTO_8D_TEXT_') and hasattr(sh,'text_frame'):
                try:
                    sh.text_frame.word_wrap=True
                    sh.text_frame.auto_size=MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                except Exception:
                    pass

    v319._page2_meta(sl,d,g)
    step11._restore_missing_4d_unit(sl,zones)


def weekly_step12(src,out,d,g,mode):
    dd=v319._weekly_display(d)
    p1=dict(dd)
    p1['issue_name']=step4.step1._page1_issue(d)
    p1['problem']=step4._full_problem(d)

    prs=Presentation(src)

    # Page 1 owns its own correct row-selection logic. Do not touch it again later.
    v319._update_page1(prs,p1,g)

    if len(prs.slides)>1:
        _update_page2_step12(prs.slides[1],dd,g,mode)
        step4._force_page2_header(prs.slides[1],d,g)

    # Synchronize page 2 only. This keeps page 1 유첨/비고 completely untouched.
    _sync_page2_signal_only(prs,d,g)

    Path(out).parent.mkdir(parents=True,exist_ok=True)
    try:
        prs.save(out); saved=out
    except PermissionError:
        p=Path(out)
        saved=str(p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix))
        prs.save(saved)

    origin=N(g.get('_issue_origin_selected'))
    if origin:
        step8._set_origin_in_ppt(saved,origin)

    return '주간회의 PPT 업데이트: '+_signal_status_from_g(d,g),saved


base.weekly=weekly_step12


class RecoveryStep12App(step10.RecoveryStep10App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP12 FIX2')


if __name__=='__main__':
    RecoveryStep12App().mainloop()
