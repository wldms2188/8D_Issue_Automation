"""Project-name-centered weekly matching patch.
The managed list uses A_B = customer A + project B. Weekly duplicate lookup prioritizes B.
"""
import main_recovery_step13 as s13
import main_recovery_step14 as s14
import project_autocomplete_final as catalog

def _project_key(value):
    return catalog.project_key(value,True)

def _task_rows_by_project(tb,hr,hm,task):
    if 'task' not in hm:return []
    target=_project_key(task)
    if not target:return []
    rows=[]
    for r in range(hr+1,len(tb.rows)):
        raw=s13._row_text(tb,r,hm['task'])
        if _project_key(raw)==target:
            rows.append(r)
    return rows

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

s14._task_rows=_task_rows_by_project
s13._row_match_score=_row_match_score_by_project
