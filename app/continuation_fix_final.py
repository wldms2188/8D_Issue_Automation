"""Continuation fixes for weekly summary placement and detail title construction.

Applied last in main_enterprise_final so these user-confirmed rules win:
1) For a task that appears on multiple summary pages, continue from the LAST matching page.
   If one more row fits, insert there. If it does not fit, insert the continuation
   summary page IMMEDIATELY AFTER that last matching page.
2) When all five standard GUI inputs are present, detail title is:
   고객사_과제명_발생샘플_개발단계_발생처 이슈 발생
   The customer/project portion is the canonical combined GUI value (task_name).
   There is intentionally NO underscore before "이슈 발생".
   If any required GUI input is missing, preserve the existing title logic.
"""

import main_recovery_step14 as s14
import main_recovery_step4 as s4
import main_recovery_step13 as s13
import main_v310 as v310
import project_autocomplete_final as catalog

N=v310.N

_original_prepare_new_summary_page=s14._prepare_new_summary_page
_original_detail_title_text=s4._detail_title_text


def _prepare_new_summary_page_after_last_match(prs,pages,g,template_index=None):
    """Keep legacy new-task placement, but overflow continuation goes right after the matched page."""
    if template_index is None:
        return _original_prepare_new_summary_page(prs,pages,g,template_index=None)

    # The original helper already clones the matched page, clears old rows, keeps
    # its local styling, and returns the new table/row. Only its insertion index
    # needs changing from "end of leading summary block" to "right after matched page".
    original_index_picker=s14._first_summary_section_insert_index
    try:
        s14._first_summary_section_insert_index=lambda _prs,_pages: min(len(_prs.slides),int(template_index)+1)
        return _original_prepare_new_summary_page(prs,pages,g,template_index=template_index)
    finally:
        s14._first_summary_section_insert_index=original_index_picker


def _detail_title_from_user_inputs(d,g):
    g=g or {}
    d=d or {}

    # Per-field priority confirmed by the user:
    # GUI input > existing extraction/legacy value.
    customer_project=(
        N(g.get('task_name'))
        or N(s13._customer_task(d))
    )
    sample=(
        N(g.get('sample'))
        or N(d.get('sample'))
        or N(d.get('occurrence_sample'))
    )
    stage=(
        N(g.get('stage'))
        or N(d.get('development_stage'))
        or N(d.get('occurrence_stage'))
        or N(d.get('stage'))
    )
    occurrence_site=(
        N(g.get('occurrence_site'))
        or N(d.get('_origin_occurrence_site'))
        or N(d.get('occurrence_site'))
    )

    # Build the requested full title whenever every component can be resolved.
    if all((customer_project,sample,stage,occurrence_site)):
        return f'{customer_project}_{sample}_{stage}_{occurrence_site} 이슈 발생'

    # If a component cannot be resolved at all, preserve the previous safe title logic.
    return _original_detail_title_text(d,g)



def _weekly_task_display(selected, task_hits):
    """Use the exact customer/project text style already present in the matched summary row."""
    if not task_hits:
        return N(selected)
    si,tb,hr,hm,rows=sorted(task_hits,key=lambda x:x[0])[-1]
    raw=N(s13._row_text(tb,max(rows),hm['task']))
    return raw or N(selected)


def _update_summary_customer_project_then_project(prs,d,g,mode):
    pages=s14._summary_pages(prs)
    if not pages:
        raise ValueError('주간회의 PPT에서 과제명/Signal 요약 양식 페이지를 찾지 못했습니다.')

    selected=N((g or {}).get('task_name')) or N(s13._customer_task(d))
    full_key=catalog._norm(selected)
    project_key=catalog.project_key(selected,False)
    issue=s13._issue_display(d)

    full_hits=[]; project_hits=[]; similar_hits=[]
    for si,tb,hr in pages:
        hm=s13._summary_map(tb,hr)
        if 'task' not in hm: continue
        fr=[]; pr=[]; sr=[]
        for r in range(hr+1,len(tb.rows)):
            raw=s13._row_text(tb,r,hm['task'])
            raw_full=catalog._norm(raw)
            raw_project=catalog.project_key(raw,False)
            if full_key and raw_full==full_key:
                fr.append(r)
            if project_key and raw_project==project_key:
                pr.append(r)
            # Last fallback: project prefix/containment similarity.
            # Example: selected MBAG can match an existing "MBAG E~".
            # Require at least 4 normalized characters to avoid weak accidental hits.
            if project_key and raw_project and min(len(project_key),len(raw_project))>=4:
                if project_key in raw_project or raw_project in project_key:
                    sr.append(r)
        if fr: full_hits.append((si,tb,hr,hm,fr))
        if pr: project_hits.append((si,tb,hr,hm,pr))
        if sr: similar_hits.append((si,tb,hr,hm,sr))

    # Global priority across ALL summary pages:
    # customer+project exact > project exact > project similar.
    task_hits=full_hits if full_hits else (project_hits if project_hits else similar_hits)
    display=_weekly_task_display(selected,task_hits)

    original_customer_task=s13._customer_task
    if task_hits:
        s13._customer_task=lambda _d: display
    try:
        if mode=='existing' and task_hits:
            best=None
            for si,tb,hr,hm,rows in task_hits:
                for r in rows:
                    sc=s13._row_match_score(tb,r,hm,display,issue)
                    if best is None or sc>best[0]: best=(sc,si,tb,hr,r)
            if best and best[0]>=140:
                _,si,tb,hr,row=best
                s14._write_summary_row(tb,row,hr,d,g)
                return si,row,'기존 행 업데이트'

        if task_hits:
            si,tb,hr,hm,rows=sorted(task_hits,key=lambda x:x[0])[-1]
            after=max(rows)
            if s14._summary_row_insert_fits(prs,si,tb,after):
                row=s14._insert_row_after(tb,after)
                s14._write_summary_row(tb,row,hr,d,g)
                return si,row,'확정 과제의 마지막 요약 행 바로 아래 삽입'
            nsi,ntb,nhr,nrow=s14._prepare_new_summary_page(prs,pages,g,template_index=si)
            s14._write_summary_row(ntb,nrow,nhr,d,g)
            return nsi,nrow,'확정 과제 요약 공간 초과로 바로 다음 페이지에 추가'

        si,tb,hr,row=s14._prepare_new_summary_page(prs,pages,g)
        s14._write_summary_row(tb,row,hr,d,g)
        return si,row,'고객사/과제명 모두 없음: 신규 요약 페이지 생성'
    finally:
        s13._customer_task=original_customer_task

s14._prepare_new_summary_page=_prepare_new_summary_page_after_last_match
s14._update_summary_by_task=_update_summary_customer_project_then_project
s4._detail_title_text=_detail_title_from_user_inputs
