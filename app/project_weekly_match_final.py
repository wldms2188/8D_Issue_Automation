"""Project-name-centered weekly matching patch.
Canonical project identity preserves parenthesized variants such as EB-L(EU)/EB-L(US).
"""
import main_recovery_step13 as s13
import main_recovery_step14 as s14
import project_autocomplete_final as catalog

def _project_key(value):
    return catalog.project_key(value,False)

def _full_key(value):
    """Normalized customer+project identity, preserving the customer when present."""
    return catalog._norm(value)

def _task_rows_by_project(tb,hr,hm,task):
    """Weekly-summary matching priority:
    1) exact customer+project match across the whole summary;
    2) only when none exists, exact project-name match regardless of customer.

    This makes the project confirmed in the GUI/candidate popup the same identity
    used to locate the weekly-summary row.
    """
    if 'task' not in hm:return []

    target_full=_full_key(task)
    target_project=_project_key(task)
    if not target_full and not target_project:return []

    full_rows=[]
    project_rows=[]
    for r in range(hr+1,len(tb.rows)):
        raw=s13._row_text(tb,r,hm['task'])
        if target_full and _full_key(raw)==target_full:
            full_rows.append(r)
        if target_project and _project_key(raw)==target_project:
            project_rows.append(r)

    return full_rows if full_rows else project_rows

def _row_match_score_by_project(tb,r,hm,task,issue):
    pk=_project_key(task)
    rk=_project_key(s13._row_text(tb,r,hm.get('task',0))) if 'task' in hm else ''
    ik=s13._k(issue)
    ri=s13._k(s13._row_text(tb,r,hm.get('issue',0))) if 'issue' in hm else ''
    score=0
    if pk and rk==pk:score+=100
    elif pk and rk and (pk in rk or rk in pk):score+=60
    if ik and ri==ik:score+=100
    elif ik and ri and (ik in ri or ri in ik):score+=70
    return score

_original_detail_score=s13._detail_match_score

def _detail_match_score_by_project(sl,d):
    structure=s13._detail_structure_score(sl)
    if structure<5:return -1
    text=s13._slide_text(sl)
    pk=_project_key(s13._customer_task(d))
    # A detail page must belong to the same canonical project before issue similarity can win.
    if pk and pk not in _project_key(text):
        return -1
    return _original_detail_score(sl,d)

def _task_slide_indices_by_project(prs,d,summary_index):
    pk=_project_key(s13._customer_task(d))
    out=[]
    if not pk:return out
    for i,sl in enumerate(prs.slides):
        if i==summary_index:continue
        if s13._detail_structure_score(sl)>=5 and pk in _project_key(s13._slide_text(sl)):
            out.append(i)
    return out

s14._task_rows=_task_rows_by_project
s13._row_match_score=_row_match_score_by_project
s13._detail_match_score=_detail_match_score_by_project
s13._task_slide_indices=_task_slide_indices_by_project
