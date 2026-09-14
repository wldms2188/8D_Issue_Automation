import datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_recovery_step12 as s12
import main_recovery_step10 as s10
import main_v310 as v310

base=s14.base
N=v310.N


def _project_parts(d):
    return (
        s13._k(s13._customer_task(d)),
        s13._k(d.get('task_name')),
        s13._k(d.get('customer')),
    )


def _project_score(sl,d):
    q=s13._k(s13._slide_text(sl))
    full,task,customer=_project_parts(d)
    if full and full in q:
        return 140
    if task and customer and task in q and customer in q:
        return 130
    if task and task in q:
        return 100
    return 0


def _project_detail_indices(prs,d):
    summary={si for si,_,_ in s14._summary_pages(prs)}
    return [i for i,sl in enumerate(prs.slides)
            if i not in summary and _project_score(sl,d)>=100]


def _find_existing_detail(prs,d):
    pages=_project_detail_indices(prs,d)
    if not pages:
        return None
    issue=s13._k(s13._issue_display(d))
    best=None
    for i in pages:
        text=s13._slide_text(prs.slides[i])
        q=s13._k(text)
        score=_project_score(prs.slides[i],d)
        score+=s13._detail_structure_score(prs.slides[i])*4
        if issue and issue in q:
            score+=140
        else:
            from difflib import SequenceMatcher
            ratios=[SequenceMatcher(None,issue,s13._k(x)).ratio()
                    for x in text.splitlines() if s13._k(x)] if issue else []
            if ratios:
                score+=int(70*max(ratios))
        if best is None or score>best[0]:
            best=(score,i)
    return best[1] if best and best[0]>=230 else None


def _new_detail_position(prs,d):
    pages=_project_detail_indices(prs,d)
    return max(pages)+1 if pages else len(prs.slides)


def _font(shape,name,size,bold=None,align=None,color=None):
    tf=shape.text_frame
    if align is not None:
        for p in tf.paragraphs:
            p.alignment=align
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name=name
            r.font.size=Pt(size)
            if bold is not None:
                r.font.bold=bold
            if color is not None:
                r.font.color.rgb=RGBColor(*color)


def _textbox(sl,x,y,w,h,text,name,size,bold=None,align=None):
    sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    sh.text=N(text)
    sh.text_frame.word_wrap=True
    sh.text_frame.vertical_anchor=MSO_ANCHOR.TOP
    _font(sh,name,size,bold,align)
    return sh


def _marker(sl,label,title,x,y,title_w):
    grp=sl.shapes.add_group_shape()
    c=grp.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(.23),Inches(.23))
    c.fill.solid(); c.fill.fore_color.rgb=RGBColor(255,192,0)
    c.line.fill.background(); c.text=label
    c.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    c.text_frame.margin_left=c.text_frame.margin_right=0
    c.text_frame.margin_top=c.text_frame.margin_bottom=0
    _font(c,'맑은 고딕',7.4,True,PP_ALIGN.CENTER,(0,0,0))
    t=grp.shapes.add_textbox(Inches(x+.15),Inches(y),Inches(title_w),Inches(.23))
    t.text=title
    t.text_frame.margin_left=t.text_frame.margin_right=0
    t.text_frame.margin_top=t.text_frame.margin_bottom=0
    _font(t,'LG스마트체 Regular',8)
    return grp


def _clean_detail_slide(prs):
    layout=None
    for ly in prs.slide_layouts:
        if len(ly.placeholders)==0:
            layout=ly; break
    if layout is None:
        layout=prs.slide_layouts[len(prs.slide_layouts)-1]
    sl=prs.slides.add_slide(layout)
    for sh in list(sl.shapes):
        if getattr(sh,'is_placeholder',False):
            try: sh._element.getparent().remove(sh._element)
            except Exception: pass

    _textbox(sl,.17,.11,2.11,.44,'과제명_이슈 제목','Arial Narrow',20)
    _textbox(sl,.20,.65,6.63,.27,'이슈명 : 8_불량명','LG스마트체 Regular',10,True)
    line=sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(.25),Inches(.54),Inches(10.55),Inches(.54))
    try: line.line.width=Pt(.75)
    except Exception: pass
    _textbox(sl,9.21,.23,1.42,.28,'00팀 담당자 : 000','LG스마트체2.0 Regular',11,None,PP_ALIGN.RIGHT)

    tb=sl.shapes.add_table(3,2,Inches(8.48),Inches(.62),Inches(1.96),Inches(.95)).table
    tb.columns[0].width=Inches(.73); tb.columns[1].width=Inches(1.23)
    vals=[('Signal','●'),('이슈기인',''),('발생단계','개발 단계 수동 기입 항목 기재 (’AA. BB. CC.)')]
    for r,(a,b) in enumerate(vals):
        base.set_cell_text(tb.cell(r,0),a,8)
        base.set_cell_text(tb.cell(r,1),b,8)

    _marker(sl,'2D','현상',.48,2.22,.42)
    _marker(sl,'3D','임시대책(필요시)',.46,3.62,1.00)
    _marker(sl,'4D','원인분석',.48,5.10,.65)
    _marker(sl,'4D','원인분석',5.64,2.17,.65)
    _marker(sl,'5D','개선대책',5.64,3.38,.65)
    _marker(sl,'6D','유효성점검',5.64,5.58,.76)
    _marker(sl,'7D','수평전개',5.68,6.94,.65)
    return sl


def _move_last(prs,index):
    ids=prs.slides._sldIdLst
    last=ids[len(ids)-1]
    ids.remove(last); ids.insert(index,last)


def weekly_fix2(src,out,d,g,mode):
    prs=Presentation(src)
    _,_,summary_action=s14._update_summary_by_task(prs,d,g,mode)

    target=_find_existing_detail(prs,d) if mode=='existing' else None
    detail_action='업데이트'
    if target is None:
        insert_at=_new_detail_position(prs,d)
        _clean_detail_slide(prs)
        _move_last(prs,insert_at)
        target=insert_at
        detail_action='공용 상세 양식으로 신규 페이지 추가'

    s13._update_detail_slide(prs.slides[target],d,g,mode)

    Path(out).parent.mkdir(parents=True,exist_ok=True)
    try:
        prs.save(out); saved=out
    except PermissionError:
        p=Path(out)
        saved=str(p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix))
        prs.save(saved)

    return ('주간회의 PPT: 요약 '+summary_action+' / 상세 '+detail_action+
            f' (과제명={N(d.get("task_name"))})',saved)


base.weekly=weekly_fix2


class RecoveryStep14Fix2App(s10.RecoveryStep10App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP14 FIX2')


if __name__=='__main__':
    RecoveryStep14Fix2App().mainloop()
