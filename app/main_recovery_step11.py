import copy
import datetime
from pathlib import Path

from pptx import Presentation

import main_recovery_step10 as step10
import main_recovery_step4 as step4
import main_recovery_step8 as step8
import main_v320 as v320
import main_v319 as v319
import main_v315 as v315
import main_v313 as v313
import main_v310 as v310

base=step10.base
N=v310.N


def _layout_with_4d_placeholders(d, imgs):
    """Keep BOTH 4D slots in the normal two-column chain even when 4D is not written yet."""
    texts={k:v313._section_text(d,k) for k in v310.ZONES}
    # Explicit requested placeholder. Do not let missing 4D collapse the layout.
    if not N(d.get('cause_4d')):
        texts['4D_CAUSE']='검토 중'
    if not (N(d.get('leak_cause')) or N(d.get('system_cause'))):
        texts['4D_LEAK']='검토 중'

    # v319._fit_chain normally removes an empty 4D. Feed placeholders only to the
    # LAYOUT copy so 2D-3D-4D and 4D-5D-6D always keep their normal vertical order.
    dl=dict(d)
    if not N(dl.get('cause_4d')):
        dl['cause_4d']='검토 중'
    if not (N(dl.get('leak_cause')) or N(dl.get('system_cause'))):
        dl['leak_cause']='검토 중'

    li={'2D':imgs.get('2D',[]),'3D':imgs.get('3D',[]),'4D_CAUSE':imgs.get('4D_CAUSE',[])}
    ri={'4D_LEAK':imgs.get('4D_LEAK',[]),'5D':imgs.get('5D',[]),'6D':imgs.get('6D',[])}
    lz,lf=v319._fit_chain(('2D','3D','4D_CAUSE'),dl,texts,li)
    rz,rf=v319._fit_chain(('4D_LEAK','5D','6D'),dl,texts,ri)
    return {**lz,**rz},texts,{**lf,**rf}


def _restore_missing_4d_unit(sl, zones):
    """Never delete a 4D marker/title. If one native unit remains, clone it only when needed.
    No fallback rectangle is created.
    """
    try:
        units=v315._top_4d_units(sl)
    except Exception:
        return
    if len(units)>=2 or not units:
        return

    unit=units[0]
    try:
        ux=float(unit.left)/v310.EMU
    except Exception:
        ux=0

    zl=zones['4D_CAUSE']; zr=zones['4D_LEAK']
    left_xy=(max(.10,zl['x']-.18),max(.10,zl['y']-.35))
    right_xy=(max(.10,zr['x']-.18),max(.10,zr['y']-.35))
    target=left_xy if ux>4.0 else right_xy

    # Clone the existing native marker/title unit. This preserves its appearance and
    # does not construct a new PowerPoint XML group.
    try:
        newel=copy.deepcopy(unit._element)
        sl.shapes._spTree.insert_element_before(newel,'p:extLst')
        dup=None
        for sh in sl.shapes:
            if sh._element is newel:
                dup=sh
                break
        if dup is not None:
            dup.name='AUTO_8D_4D_RESTORED'
            bx,by,_,_=v310.box(dup)
            from pptx.util import Inches
            dup.left += Inches(target[0])-int(bx)
            dup.top += Inches(target[1])-int(by)
    except Exception:
        pass


def _update_page2_step11(sl,d,g,mode):
    # Remove only prior AUTO text/images; keep the template/native 2D~6D graphics.
    v310.remove_previous_auto(sl)
    imgs=d.get('_section_images',{}) or {}
    zones,texts,fonts=_layout_with_4d_placeholders(d,imgs)

    # Existing issue: preserve the current marker/title positions.
    # New issue: retain the same established movement behavior for 2D/3D/5D/6D.
    if v310.is_new(mode):
        for label,key in [('2D','2D'),('3D','3D'),('5D','5D'),('6D','6D')]:
            z=zones[key]
            v310.move_marker_unit(sl,label,max(.10,z['x']-.18),max(.10,z['y']-.35))

    # 4D is ALWAYS rendered. Missing source content becomes '검토 중'.
    for key in ('2D','3D','4D_CAUSE','4D_LEAK','5D','6D'):
        v319._render(sl,key,zones[key],texts[key],imgs.get(key,[]),fonts[key])

    v319._page2_meta(sl,d,g)

    # Previous versions deleted 4D units when the cause was empty. Never do that now.
    # If the input is already a prior generated file with only one 4D unit, restore
    # the missing unit by cloning the surviving native unit rather than drawing a box.
    _restore_missing_4d_unit(sl,zones)


def weekly_step11(src,out,d,g,mode):
    dd=v319._weekly_display(d)
    p1=dict(dd)
    p1['issue_name']=step4.step1._page1_issue(d)
    p1['problem']=step4._full_problem(d)

    prs=Presentation(src)
    v319._update_page1(prs,p1,g)

    if len(prs.slides)>1:
        _update_page2_step11(prs.slides[1],dd,g,mode)
        step4._force_page2_header(prs.slides[1],d,g)

    Path(out).parent.mkdir(parents=True,exist_ok=True)
    try:
        prs.save(out); saved=out
    except PermissionError:
        p=Path(out)
        saved=str(p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix))
        prs.save(saved)

    # Keep STEP8 issue-origin formatting/update after the PPT has been saved.
    origin=N(g.get('_issue_origin_selected'))
    if origin:
        step8._set_origin_in_ppt(saved,origin)

    return '주간회의 PPT 업데이트: '+base.status(d),saved


base.weekly=weekly_step11


class RecoveryStep11App(step10.RecoveryStep10App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP11')


if __name__=='__main__':
    RecoveryStep11App().mainloop()
