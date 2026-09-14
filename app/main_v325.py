# 8D Issue Automation v3.2.5
# Generic symptom-only summary; optional test/build label; DB uses same summary and numbered causes.
import re
import main_v324 as v324
import main_v323 as v323
import main_v322 as v322
import main_v321 as v321

base=v324.base
N=v321.N
C=v321.C


def _event_name(d):
    """Recognize tests plus build IDs such as A1, B2, C10, D3."""
    kind,name=v321._event_name(d)
    if kind=='시험' and name:
        return kind,name
    sources=[N(d.get('issue_name')),N(d.get('problem'))]
    for src in sources:
        # Build format requested by user: A-D followed by one or more digits.
        m=re.search(r'(?<![A-Za-z0-9])([A-D]\d+)(?![A-Za-z0-9])',src,re.I)
        if m:
            return '빌드',m.group(1).upper()
    if kind=='빌드' and name:
        return kind,name
    return '',''


def _short_problem(d):
    """For every issue type keep actual symptom; add test/build name only when applicable."""
    kind,name=_event_name(d)
    symptoms=v323._symptom_lines(d.get('problem'),name if kind in ('시험','빌드') else '')
    if not symptoms:
        for raw in N(d.get('problem')).split('\n'):
            s=v323._clean_problem_line(raw)
            if not s or v323._looks_like_condition(s):
                continue
            if name and C(s)==C(name):
                continue
            symptoms.append(s); break
    lines=list(symptoms)
    if kind=='시험' and name:
        lines.append(f'시험명: {name}')
    elif kind=='빌드' and name:
        lines.append(f'빌드명: {name}')
    return '\n'.join(x for x in lines if N(x)).strip() or N(d.get('problem'))


def _status_reason(d):
    """Status decision remains 6D-driven; user-facing reason is simply the 6D text."""
    raw=N(d.get('verification_6d'))
    if not raw:
        return 'open','(미기재)'
    pats=(r'진행\s*중',r'검증\s*중',r'확인\s*중',r'모니터링\s*중',r'시험\s*중',r'적용\s*중',r'조치\s*중',r'분석\s*중',r'예정',r'계획',r'추후')
    if any(re.search(p,raw,re.I) for p in pats):
        return 'open',raw
    return 'close',raw


def _numbered(text):
    """Normalize source cause lines to 1), 2), ... without inventing content."""
    raw=N(text)
    if not raw:
        return []
    parts=[]
    for line in raw.split('\n'):
        s=line.strip()
        s=re.sub(r'^[-•·▪◦]\s*','',s)
        s=re.sub(r'^\(?\d+\)?[.)]\s*','',s)
        if s:
            parts.append(s)
    if not parts and raw.strip():
        parts=[raw.strip()]
    return [f'{i}) {s}' for i,s in enumerate(parts,1)]


def _db_cause(d):
    out=[]
    occ=_numbered(d.get('cause_4d'))
    leak=_numbered(d.get('leak_cause'))
    system=_numbered(d.get('system_cause'))
    if occ:
        out.append('1. 발생원인')
        out.extend(occ)
    if leak or system:
        out.append('2. 유출원인')
        # System cause is kept with leak-side cause because DB format requested two categories.
        combined=[]
        for seq in (leak,system):
            for item in seq:
                combined.append(re.sub(r'^\d+\)\s*','',item))
        out.extend(f'{i}) {s}' for i,s in enumerate(combined,1))
    return '\n'.join(out)


def _excel_status(d,g=None):
    if g and g.get('_status_override') in ('open','close'):
        return g['_status_override']
    return _status_reason(d)[0]


def _write_row(ws,r,d,g):
    y,m=base.parse_year_month(d.get('occurrence_date'))
    vals={
        2:v321.datetime.date.today(),3:g.get('plm_no',''),4:g.get('form_factor',''),5:g.get('product_type',''),
        6:y,7:v321._month_label(m),8:g.get('team',''),9:g.get('owner',''),10:base.task(d),11:g.get('sample',''),
        12:g.get('stage',''),13:g.get('occurrence_site') or d.get('occurrence_site',''),
        14:_short_problem(d),15:'',16:_db_cause(d),17:N(d.get('action_5d')),18:_excel_status(d,g)
    }
    for c,v in vals.items(): ws.cell(r,c).value=v
    v321._add_excel_photo(ws,r,base.representative_image(d))
    return vals[18]

# Patch inherited behavior.
v321._event_name=_event_name
v321._short_problem=_short_problem
v322._status_reason=_status_reason
v322._write_row=_write_row
base.write_row=_write_row


class App(v324.App):
    def __init__(self):
        super().__init__(); self.title('8D 이슈 자동화 v3.2.5')

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d'):
            return v321.messagebox.showwarning('확인','8D PPT를 선택하세요.')
        try:
            d=base.extract(g['ppt8d']); kind,event=_event_name(d); st,six=_status_reason(d)
            self.log.delete('1.0','end'); self.log.insert('end','=== 8D 추출 결과 ===\n')
            for k in ['issue_name','task_name','customer','occurrence_date','problem','cause_4d','leak_cause','system_cause','action_5d','verification_6d','spread_7d']:
                self.log.insert('end',f'{k}: {d.get(k)}\n')
            self.log.insert('end',f'\n간략 현상:\n{_short_problem(d)}\n\n시험/빌드 판정: {kind} / {event}\nDB 원인:\n{_db_cause(d)}\n\nDB 상태: {st}\n판단 이유(6D 내용): {six}')
        except Exception as e:
            v321.messagebox.showerror('오류',repr(e))

    def _prepare_choices(self,d,g):
        # Keep v3.2.1 origin/problem choices, but use the corrected 6D-only display for close confirmation.
        if v321._cause_known(d):
            dlg=v321.ChoiceDialog(self,'이슈기인 선택','원인이 확인되어 있습니다. 이슈기인을 선택하세요.',v321.ISSUE_ORIGINS)
            if dlg.result is None: return None
            g['issue_origin']=dlg.result
        else:
            g['issue_origin']='TBD'
        if v321._problem_is_long(d):
            g['_short_problem']=bool(v321.messagebox.askyesno('현상 내용 확인','현상 내용이 깁니다.\n주요 논의 사항에는 실제 현상만 간략하게 기재할까요?\n시험/빌드 이슈이면 시험명/빌드명도 추가됩니다.'))
        else:
            g['_short_problem']=False
        st,six=_status_reason(d)
        if st=='close':
            msg=('6D까지 작성된 내용을 기준으로 이 이슈를 close로 판단했습니다.\n\n'
                 f'[판단 이유 / 6D 내용]\n{six}\n\n이슈 상태를 close로 처리할까요?\n\n예: close로 처리\n아니오: open으로 처리')
            g['_status_override']='close' if v321.messagebox.askyesno('이슈 상태 최종 확인',msg) else 'open'
        else:
            g['_status_override']='open'
        return g


if __name__=='__main__':
    App().mainloop()
