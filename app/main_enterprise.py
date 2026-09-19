"""Enterprise presentation layer for 8D Issue Automation.
Keeps the validated processing core and replaces presentation/UI only.
"""
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from openpyxl import load_workbook
from pptx import Presentation

import main_final as legacy
import ui_enterprise as ui

try:
    from tkinterdnd2 import TkinterDnD
    _RootBase=TkinterDnD.Tk
except Exception:
    _RootBase=tk.Tk

base=legacy.base; N=legacy.N

def target_ready_status(ppt8d,current):
    current=str(current or '').strip()
    if ppt8d and ('8D 원본을 선택' in current or not current):
        return 'READY  ·  8D 원본 선택 완료 · 미리보기 가능'
    if not ppt8d and current.startswith('READY'):
        return 'READY  ·  8D 원본을 선택해 주세요.'
    return current

def split_customer_project(value):
    """Split canonical A_B into customer A and full customer_project A_B.

    Underscores inside parentheses are part of the project name and are ignored.
    """
    s=N(value)
    depth=0
    for i,ch in enumerate(s):
        if ch=='(':
            depth+=1
        elif ch==')' and depth:
            depth-=1
        elif ch=='_' and depth==0:
            return s[:i].strip(),s
    return '',s

def apply_selected_customer_project(d,selected):
    d=dict(d or {})
    selected=N(selected)
    if not selected:
        return d
    customer,full=split_customer_project(selected)
    d['task_name']=full
    if customer:
        d['customer']=customer
    return d

def canonical_customer_project_from_8d(d):
    """Return one canonical customer_project label from extracted 8D metadata."""
    d=dict(d or {})
    task=N(d.get('task_name'))
    customer=N(d.get('customer'))
    if task:
        prefix,full=split_customer_project(task)
        if prefix:
            return full
        if customer:
            return customer+'_'+task
        return task
    return customer

def _customer_project_key(value):
    return re.sub(r'[^0-9A-Za-z가-힣]+','',N(value)).casefold()

def customer_project_mismatch(extracted_d,selected_value):
    extracted=canonical_customer_project_from_8d(extracted_d)
    selected=N(selected_value)
    if not extracted or not selected:
        return False,extracted,selected
    return _customer_project_key(extracted)!=_customer_project_key(selected),extracted,selected
WEEKLY_PENDING_TERMS=(
    '진행 중','진행중','검증 중','검증중','확인 중','확인중','검토 중','검토중',
    '예정','추정','계획','계획 중','계획중','미완료','완료 예정','추가 검토','모니터링 중','모니터링중',
    '추적 중','추적중','시험 예정','검증 예정','확인 예정','추가 검증','추가 확인',
    'in progress','ongoing','pending','planned','scheduled','tbd','tbc','wip',
    'under verification','under validation','under review','in review',
    'under evaluation','in evaluation','under monitoring','monitoring ongoing',
    'to be verified','to be validated','to be completed','to be confirmed',
    'not completed','not complete','not verified','not validated','not yet complete',
    'not yet completed','not yet verified','not yet validated',
    'awaiting verification','awaiting validation','awaiting result','awaiting results',
    'follow-up ongoing','follow up ongoing','verification pending','validation pending',
    'planned completion','scheduled completion','target completion','will be verified',
    'will be validated','will be completed','expected to be completed',
    'verification required','validation required','further verification required',
    'further validation required','additional verification required',
    'additional validation required','remaining verification','remaining validation',
    'verification to follow','validation to follow','follow-up required','follow up required'
)
WEEKLY_COMPLETE_TERMS=(
    '검증 완료','검증완료','개선 완료','개선완료','확인 완료','확인완료',
    '시험 완료','시험완료','평가 완료','평가완료','적용 완료','적용완료',
    '정상','이상 없음','이상없음','양호','문제 없음','문제없음','재발 없음','재발없음',
    '추가 이상 없음','추가이상없음','추가 불량 없음','추가불량없음',
    'verification complete','verification completed','verification is complete','verification finished',
    'validation complete','validation completed','validation is complete','validation finished',
    'evaluation complete','evaluation completed','evaluation finished',
    'test complete','test completed','testing completed','test finished',
    'effectiveness confirmed','effectiveness verified','effectiveness validated','effective',
    'completed','complete','successfully completed','finished','done',
    'validated','verified','passed','pass','all tests passed','test passed',
    'validation passed','verification passed','evaluation passed',
    'acceptable','satisfactory','normal result','normal condition','result normal','stable result',
    'no abnormality','no abnormalities','no abnomality','no abnomalities',
    'no additional abnormality','no additional abnormalities','no additional abnomality','no additional abnomalities',
    'no further abnormality','no further abnormalities','without abnormality','without abnormalities',
    'no issue','no issues','no additional issue','no additional issues',
    'no defect','no defects','no additional defect','no additional defects',
    'no recurrence','no recurrence observed','no repeat issue','no repeated issue',
    'no abnormality observed','no abnormality detected','no abnormality found','no abnormality occurred',
    'no abnormalities observed','no abnormalities detected','no abnormalities found','no abnormalities occurred',
    'no abnomality observed','no abnomality detected','no abnomality found',
    'zero defect','zero defects','within spec','within specification',
    'meets spec','meets specification','met spec','met specification',
    'criteria met','acceptance criteria met','all criteria met',
    'requirement met','requirements met','target met','all requirements met'
)
WEEKLY_ABNORMAL_TERMS=(
    '불량','이상 발생','이상발생','미흡','재발','부적합','실패','ng','nok','oos',
    'abnormal','abnomal','abnormality','abnomality','abnormalities','abnomalities',
    'fail','failed','failure','not ok','out of spec','out-of-spec',
    'defect remains','issue remains','recurred','recurrence','not acceptable',
    'criteria not met','acceptance criteria not met','target not met',
    'requirement not met','requirements not met','still abnormal','abnormality observed',
    'abnormality detected','abnormality found','defect observed','defect detected','defect found'
)

ACTION_PENDING_TERMS=(
    '진행 중','진행중','예정','계획','미완료','적용 예정','반영 예정','조치 예정',
    'in progress','ongoing','pending','planned','scheduled','not completed','not complete',
    'to be implemented','to be applied','to be released','to be updated',
    'implementation pending','implementation in progress','application pending',
    'will be implemented','will be applied','expected to be completed'
)
ACTION_COMPLETE_TERMS=(
    '개선 완료','개선완료','조치 완료','조치완료','적용 완료','적용완료','반영 완료','반영완료',
    '변경 완료','변경완료','시행 완료','시행완료','대책 완료','대책완료',
    'completed','complete','implemented','implementation completed','successfully implemented',
    'applied','application completed','released','deployed','updated','revised',
    'corrective action completed','corrective action implemented','countermeasure implemented',
    'countermeasure applied','action completed','action implemented','done','finished'
)

def _normalize_weekly_status_text(text):
    q=re.sub(r'\s+',' ',N(text)).casefold().strip()
    # Common spelling errors in customer 8D files.
    q=q.replace('abnomalities','abnormalities').replace('abnomality','abnormality').replace('abnomal','abnormal')
    return q

def _result_tail_state(q):
    """Interpret an explicit result/value after ':'/'→' as higher-value evidence.

    Examples:
      진행 중 : 이상없음        -> complete
      Progress: No abnormalities -> complete
      진행 중 : 이상 발생       -> abnormal
      진행 : 추가 검증 예정     -> pending

    'so far/to date/현재까지' keeps the result provisional rather than complete.
    """
    tails=[]
    for raw in re.split(r'[\n\r;]+',q):
        line=raw.strip()
        if not line:
            continue
        parts=re.split(r'\s*(?::|：|→|=>|⇒)\s*',line,maxsplit=1)
        if len(parts)==2 and parts[1].strip():
            tails.append(parts[1].strip())

    if not tails:
        return None,None

    provisional_terms=(
        '현재까지','현재 까지','지금까지','현 시점','현재 시점',
        'so far','to date','as of now','currently','at this time'
    )
    pending_terms=tuple(_normalize_weekly_status_text(x) for x in WEEKLY_PENDING_TERMS)
    complete_terms=tuple(_normalize_weekly_status_text(x) for x in WEEKLY_COMPLETE_TERMS)
    abnormal_terms=tuple(_normalize_weekly_status_text(x) for x in WEEKLY_ABNORMAL_TERMS)

    strong_normal_patterns=(
        r'이상\s*없(?:음|다|었습니다|었다)?',
        r'추가\s*이상\s*없(?:음|다|었습니다|었다)?',
        r'문제\s*없(?:음|다|었습니다|었다)?',
        r'재발\s*없(?:음|다|었습니다|었다)?',
        r'\bno\s+(?:further\s+|additional\s+|new\s+|repeated\s+)?abnormalit(?:y|ies)\b',
        r'\bwithout\s+(?:any\s+)?abnormalit(?:y|ies)\b',
        r'\bno\s+(?:further\s+|additional\s+|new\s+)?defects?\b',
        r'\bno\s+(?:further\s+|additional\s+|new\s+)?issues?\b',
        r'\bno\s+(?:further\s+|additional\s+)?recurrence\b',
        r'\bno\s+repeat(?:ed)?\s+issues?\b',
        r'\bzero\s+defects?\b',
    )

    # The final explicit result is the strongest; scan from the end.
    for tail in reversed(tails):
        t=_normalize_weekly_status_text(tail)
        if not t:
            continue

        provisional=any(x in t for x in provisional_terms)

        normal_mask=t
        strong_normal=any(re.search(p,t) for p in strong_normal_patterns)
        for p in strong_normal_patterns:
            normal_mask=re.sub(p,' ',normal_mask)

        abnormal=(
            any(x in normal_mask for x in abnormal_terms)
            or bool(re.search(r'\b(?:abnormalit(?:y|ies)|abnormal|ng|nok|fail(?:ed|ure)?|oos|recur(?:red|rence))\b',normal_mask))
        )
        if abnormal:
            return 'abnormal','결과 항목의 이상/실패/재발 표현 감지'

        tail_pending=any(x in t for x in pending_terms)
        if tail_pending:
            return 'pending','결과 항목의 진행 중/예정 표현 감지'

        tail_complete=strong_normal or any(x in t for x in complete_terms)
        if tail_complete:
            if provisional:
                return 'pending','결과는 이상 없으나 현재까지/so far 표현으로 검증 진행 중'
            return 'complete','결과 항목의 완료/정상/이상 없음 표현 감지'

    return None,None

def weekly_verification_state(text):
    """Return complete/pending/abnormal/unknown from Korean or English 6D wording."""
    q=_normalize_weekly_status_text(text)
    if not q:
        return 'unknown','6D 내용 없음'

    tail_state,tail_reason=_result_tail_state(q)
    if tail_state:
        return tail_state,tail_reason

    strong_normal_patterns=(
        r'\bno\s+(?:further\s+|additional\s+|new\s+|repeated\s+)?abnormalit(?:y|ies)\b',
        r'\bwithout\s+(?:any\s+)?abnormalit(?:y|ies)\b',
        r'\bno\s+(?:further\s+|additional\s+|new\s+)?defects?\b',
        r'\bno\s+(?:further\s+|additional\s+|new\s+)?issues?\b',
        r'\bno\s+(?:further\s+|additional\s+)?failures?\b',
        r'\bno\s+(?:further\s+|additional\s+)?recurrence\b',
        r'\bno\s+repeat(?:ed)?\s+issues?\b',
        r'\bzero\s+defects?\b',
    )
    strong_normal=any(re.search(p,q) for p in strong_normal_patterns)

    # Mask confirmed "no abnormality/no defect" phrases before looking for
    # abnormal tokens, so "No additional abnormalities" can never be read as abnormal.
    abnormal_scan=q
    for p in strong_normal_patterns:
        abnormal_scan=re.sub(p,' ',abnormal_scan)

    abnormal=(
        any(_normalize_weekly_status_text(x) in abnormal_scan for x in WEEKLY_ABNORMAL_TERMS)
        or bool(re.search(r'\b(?:abnormalit(?:y|ies)|abnormal|ng|nok|fail(?:ed|ure)?|oos|recur(?:red|rence))\b',abnormal_scan))
    )
    if abnormal:
        return 'abnormal','이상/실패/재발 표현 감지'

    pending=any(_normalize_weekly_status_text(x) in q for x in WEEKLY_PENDING_TERMS)

    # A definitive normal/no-additional-abnormality result is considered complete
    # unless the text explicitly says verification/validation itself is still pending.
    explicit_verification_pending=any(x in q for x in (
        'under verification','under validation','verification pending','validation pending',
        'verification in progress','validation in progress','verification scheduled',
        'validation scheduled','to be verified','to be validated',
        'not yet verified','not yet validated','verification required','validation required',
        'additional verification required','additional validation required',
        'remaining verification','remaining validation'
    ))
    if strong_normal and not explicit_verification_pending:
        return 'complete','이상 없음/추가 이상 없음/재발 없음 표현 감지'

    if pending:
        return 'pending','진행 중/검증 중/예정 표현 감지'

    if any(_normalize_weekly_status_text(x) in q for x in WEEKLY_COMPLETE_TERMS):
        return 'complete','완료/정상/검증 완료 표현 감지'
    if re.search(r'\b(?:complet(?:e|ed)|finish(?:ed)?|verif(?:ied|ication complete(?:d)?)|validat(?:ed|ion complete(?:d)?)|pass(?:ed)?|normal|acceptable|satisfactory|effective)\b',q):
        return 'complete','완료/정상/검증 완료 표현 감지'
    return 'unknown','6D 내용은 있으나 완료/진행 상태를 확정할 표현이 명확하지 않음'

def action_5d_state(text):
    """Classify 5D implementation state; used as supporting evidence for Issue DB."""
    q=_normalize_weekly_status_text(text)
    if not q:
        return 'unknown'
    if any(_normalize_weekly_status_text(x) in q for x in ACTION_PENDING_TERMS):
        return 'pending'
    if any(_normalize_weekly_status_text(x) in q for x in ACTION_COMPLETE_TERMS):
        return 'complete'
    if re.search(r'\b(?:implement(?:ed|ation completed)|appl(?:ied|ication completed)|deploy(?:ed|ment completed)|releas(?:ed|e completed)|revis(?:ed|ion completed)|updat(?:ed|e completed)|complet(?:e|ed)|finished|done)\b',q):
        return 'complete'
    return 'unknown'

def issue_db_recommended_status(d):
    """Unified bilingual Issue DB recommendation using both 5D and 6D evidence."""
    action=N((d or {}).get('action_5d'))
    verify=N((d or {}).get('verification_6d'))
    verify_state,verify_reason=weekly_verification_state(verify)
    action_state=action_5d_state(action)

    if verify_state=='abnormal':
        return 'open','6D에서 이상/실패/재발 결과가 확인되었습니다.'
    if verify_state=='pending':
        return 'open','6D 검증/확인이 진행 중이거나 예정 상태입니다.'
    if verify_state=='complete':
        if action_state=='complete':
            return 'close','5D 개선대책 완료와 6D 완료/정상/추가 이상 없음 결과가 함께 확인되었습니다.'
        return 'close','6D에서 완료/정상/추가 이상 없음 결과가 확인되었습니다.'

    # Keep the earlier V1 behavior for a non-empty 6D with no pending/abnormal
    # evidence: propose close and let the user confirm in the status window.
    if verify:
        if action_state=='complete':
            return 'close','5D 개선대책 완료가 확인되고 6D에 진행 중/이상 표현이 없습니다.'
        return 'close','6D 내용이 있으며 진행 중/예정/이상 표현이 명확히 확인되지 않았습니다.'

    # 5D alone is not enough to close without 6D effectiveness evidence.
    if action_state=='complete':
        return 'open','5D 개선대책은 완료되었으나 6D 효과검증 내용이 없어 open으로 판단합니다.'
    if action:
        return 'open','5D 개선대책은 있으나 6D 효과검증 완료가 확인되지 않았습니다.'
    return 'open','5D/6D 완료 근거가 없습니다.'

def weekly_recommended_status(d):
    """Recommend weekly Signal using the same Korean/English 6D state rules."""
    cause=N(d.get('cause_4d'))
    action=N(d.get('action_5d'))
    verify=N(d.get('verification_6d'))
    state,_reason=weekly_verification_state(verify)

    if state=='complete':
        return '개선 완료'
    if state in ('pending','abnormal'):
        return '개선 검증중'
    if cause or action or verify:
        return '개선 검증중'
    return '원인/개선 미확인'

def weekly_status_from_choice(d,choice):
    # Kept for compatibility with older callers. Close still forces green;
    # open uses the agreed 8D-content recommendation logic.
    if str(choice or '').lower()=='close':
        return '개선 완료'
    return weekly_recommended_status(d)

def weekly_status_confirmation_required(issue_db_status):
    """Only a final open Issue DB status needs a separate weekly Signal choice."""
    return str(issue_db_status or '').strip().lower()=='open'

def issue_db_status_confirmation_required(d):
    """Show one Issue DB status chooser whenever 5D or 6D has content."""
    return bool(N(d.get('action_5d')) or N(d.get('verification_6d')))

def issue_db_status_reason(d,judged):
    """Human-readable reason for the Issue DB open/close proposal."""
    judged='close' if str(judged or '').lower()=='close' else 'open'
    sixd=N(d.get('verification_6d'))
    state,reason=weekly_verification_state(sixd)
    if not sixd:
        return '6D 효과검증 내용이 없어 완료 여부를 확인하기 어렵습니다.'
    if judged=='close' and state=='complete':
        return '6D에서 완료·정상·이상 없음·검증 완료 의미의 표현이 확인되었습니다.'
    if state=='pending':
        return '6D에서 진행 중·검증 중·예정·계획 의미의 표현이 확인되었습니다.'
    if state=='abnormal':
        return '6D에서 이상·실패·재발 의미의 표현이 확인되었습니다.'
    return '6D에 완료를 확정할 표현이 명확하지 않습니다.'

def weekly_status_reason(d,weekly_status):
    st=N(weekly_status)
    cause=N(d.get('cause_4d')); action=N(d.get('action_5d')); verify=N(d.get('verification_6d'))
    state,_reason=weekly_verification_state(verify)

    if st=='개선 완료':
        if state=='complete':
            return '6D 효과검증에서 완료·정상·이상 없음·검증 완료 의미의 표현이 확인되었습니다.'
        return '6D 효과검증 결과가 완료 상태로 선택되었습니다.'

    if st=='개선 검증중':
        if state=='pending':
            return '6D에서 진행 중·검증 중·예정·계획 또는 미완료 상태를 의미하는 표현이 확인되었습니다.'
        if state=='abnormal':
            return '6D에서 이상·실패·재발 의미의 표현이 확인되어 개선 검증 중으로 추천합니다.'
        if verify:
            return '6D 내용은 있으나 검증 완료가 명확히 확인되지 않아 개선 검증 중으로 추천합니다.'
        if action:
            return '5D 개선대책은 작성되어 있으나 6D 완료 검증이 아직 없어 개선 검증 중으로 추천합니다.'
        if cause:
            return '4D 원인 내용은 확인되지만 개선 완료 단계는 아니므로 개선 검증 중으로 추천합니다.'
    return '4D 원인, 5D 개선대책, 6D 효과검증 내용이 모두 확인되지 않아 원인/개선 미확인으로 추천합니다.'






class EnterpriseOriginDialog(tk.Toplevel):
    def __init__(self,parent,d):
        super().__init__(parent); self.withdraw(); self.result=None; self.title('이슈기인 확인'); self.configure(bg=ui.WHITE); self.resizable(False,False); self.transient(parent)
        cause=legacy.step8._cause_text(d); rec,reason=legacy.step8._recommend_origin(d)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='이슈기인 확인',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
        body=tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=18)
        tk.Label(body,text='8D 원인 내용을 기준으로 이슈기인을 확인합니다.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold')).pack(anchor='w')
        tk.Label(body,text='8D 원인 내용',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w',pady=(14,5))
        box=tk.Text(body,height=10,wrap='word',bg='#F7F9FB',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=10,pady=8)
        box.pack(fill='x'); box.insert('1.0',cause or '(4D 원인 내용 없음)'); box.configure(state='disabled')
        card=tk.Frame(body,bg='#EEF5FA',highlightbackground='#D5E3EE',highlightthickness=1); card.pack(fill='x',pady=12)
        tk.Label(card,text=f'추천  {rec}',bg='#EEF5FA',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(anchor='w',padx=12,pady=(9,2))
        tk.Label(card,text=reason,bg='#EEF5FA',fg='#4E6375',font=('Malgun Gothic',8),wraplength=690,justify='left').pack(anchor='w',padx=12,pady=(0,9))
        self.var=tk.StringVar(value=rec)
        if rec=='TBD':
            tk.Label(body,text='원인 미확정 상태이므로 TBD로 반영됩니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',9)).pack(anchor='w')
        else:
            row=tk.Frame(body,bg='white'); row.pack(fill='x',pady=(2,5))
            for value in legacy.step8.ORIGINS:
                ttk.Radiobutton(row,text=value+('  · 추천' if value==rec else ''),variable=self.var,value=value,style='Mode.TRadiobutton').pack(side='left',padx=(0,14))
        foot=tk.Frame(self,bg='#F6F8FA',height=62); foot.pack(fill='x'); foot.pack_propagate(False)
        b=tk.Frame(foot,bg='#F6F8FA'); b.pack(side='right',padx=20,pady=12)
        self._button(b,'취소',self.cancel,False).pack(side='left',padx=4); self._button(b,'확인',self.ok,True).pack(side='left',padx=4)
        self.protocol('WM_DELETE_WINDOW',self.cancel); ui.center_window(self,parent,760,610); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary): return tk.Button(p,text=text,command=cmd,width=12,bd=0,font=('Malgun Gothic',9,'bold'),bg=ui.BLUE if primary else '#E5EBF0',fg='white' if primary else ui.TEXT,pady=7,cursor='hand2')
    def ok(self): self.result=self.var.get(); self.destroy()
    def cancel(self): self.result=None; self.destroy()


WEEKLY_SIGNAL_COLORS={
    '원인/개선 미확인':'#FF0000',
    '개선 검증중':'#FFC000',
    '개선 완료':'#00B050',
}
WEEKLY_SIGNAL_DISPLAY={
    '원인/개선 미확인':'원인/개선 미확인',
    '개선 검증중':'개선 검증 중',
    '개선 완료':'개선 완료',
}

class EnterpriseWeeklyStatusDialog(tk.Toplevel):
    """Weekly Signal confirmation using the same visual pattern as 이슈기인 확인."""
    def __init__(self,parent,d,recommended):
        super().__init__(parent); self.withdraw(); self.result=None
        self.title('주간회의 상태 확인'); self.configure(bg=ui.WHITE); self.resizable(False,False); self.transient(parent)
        recommended=N(recommended) if N(recommended) in WEEKLY_SIGNAL_COLORS else '원인/개선 미확인'
        five=N(d.get('action_5d')) or '(5D 개선대책 내용 없음)'
        six=N(d.get('verification_6d')) or '(6D 효과검증 내용 없음)'
        reason=weekly_status_reason(d,recommended)

        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='주간회의 상태 확인',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)

        body=tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=18)
        tk.Label(body,text='8D 진행 내용을 기준으로 주간회의 Signal을 확인합니다.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold')).pack(anchor='w')
        tk.Label(body,text='판단 기준이 된 5D / 6D 내용',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w',pady=(14,5))
        box=tk.Text(body,height=8,wrap='word',bg='#F7F9FB',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=10,pady=8)
        box.pack(fill='x')
        box.insert('1.0',f'[5D 개선대책]\n{five}\n\n[6D 효과검증]\n{six}')
        box.configure(state='disabled')

        card=tk.Frame(body,bg='#EEF5FA',highlightbackground='#D5E3EE',highlightthickness=1); card.pack(fill='x',pady=12)
        top=tk.Frame(card,bg='#EEF5FA'); top.pack(fill='x',padx=12,pady=(9,2))
        tk.Label(top,text='추천',bg='#EEF5FA',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left')
        tk.Label(top,text='●',bg='#EEF5FA',fg=WEEKLY_SIGNAL_COLORS[recommended],font=('Malgun Gothic',12,'bold')).pack(side='left',padx=(9,4))
        tk.Label(top,text=WEEKLY_SIGNAL_DISPLAY[recommended],bg='#EEF5FA',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left')
        tk.Label(card,text=reason,bg='#EEF5FA',fg='#4E6375',font=('Malgun Gothic',8),wraplength=690,justify='left').pack(anchor='w',padx=12,pady=(0,9))

        self.var=tk.StringVar(value=recommended)
        row=tk.Frame(body,bg='white'); row.pack(fill='x',pady=(2,5))
        for value in ('원인/개선 미확인','개선 검증중','개선 완료'):
            opt=tk.Frame(row,bg='white'); opt.pack(side='left',padx=(0,18))
            tk.Label(opt,text='●',bg='white',fg=WEEKLY_SIGNAL_COLORS[value],font=('Malgun Gothic',11,'bold')).pack(side='left',padx=(0,3))
            ttk.Radiobutton(opt,text=WEEKLY_SIGNAL_DISPLAY[value]+('  · 추천' if value==recommended else ''),variable=self.var,value=value,style='Mode.TRadiobutton').pack(side='left')

        foot=tk.Frame(self,bg='#F6F8FA',height=62); foot.pack(fill='x'); foot.pack_propagate(False)
        b=tk.Frame(foot,bg='#F6F8FA'); b.pack(side='right',padx=20,pady=12)
        self._button(b,'취소',self.cancel,False).pack(side='left',padx=4)
        self._button(b,'확인',self.ok,True).pack(side='left',padx=4)

        self.protocol('WM_DELETE_WINDOW',self.cancel)
        ui.center_window(self,parent,760,575); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)

    def _button(self,p,text,cmd,primary):
        return tk.Button(p,text=text,command=cmd,width=12,bd=0,font=('Malgun Gothic',9,'bold'),bg=ui.BLUE if primary else '#E5EBF0',fg='white' if primary else ui.TEXT,pady=7,cursor='hand2')
    def ok(self): self.result=self.var.get(); self.destroy()
    def cancel(self): self.result=None; self.destroy()


class EnterpriseProblemDialog(tk.Toplevel):
    def __init__(self,parent,summary):
        super().__init__(parent); self.withdraw(); self.result=None; self.title('Issue DB 현상 입력'); self.configure(bg='white'); self.resizable(False,False); self.transient(parent)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False); tk.Label(head,text='Issue DB 현상 입력',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
        body=tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=18)
        tk.Label(body,text='Issue DB의 현상 항목에 입력할 범위를 선택해 주세요.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold')).pack(anchor='w')
        tk.Label(body,text='주간회의 현상은 선택과 관계없이 원문 전체를 유지합니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',pady=(3,13))
        tk.Label(body,text='현상 요약 미리보기',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w',pady=(0,5))
        box=tk.Text(body,height=12,wrap='word',bg='#F7F9FB',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=10,pady=8); box.pack(fill='both',expand=True); box.insert('1.0',summary or '(요약 내용 없음)'); box.configure(state='disabled')
        foot=tk.Frame(self,bg='#F6F8FA',height=66); foot.pack(fill='x'); foot.pack_propagate(False); b=tk.Frame(foot,bg='#F6F8FA'); b.pack(side='right',padx=20,pady=13)
        self._button(b,'취소',lambda:self.choose(None),False).pack(side='left',padx=4); self._button(b,'전체 내용',lambda:self.choose('full'),False).pack(side='left',padx=4); self._button(b,'현상 요약',lambda:self.choose('summary'),True).pack(side='left',padx=4)
        self.protocol('WM_DELETE_WINDOW',lambda:self.choose(None)); ui.center_window(self,parent,720,500); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary): return tk.Button(p,text=text,command=cmd,width=12,bd=0,font=('Malgun Gothic',9,'bold'),bg=ui.BLUE if primary else '#E5EBF0',fg='white' if primary else ui.TEXT,pady=7,cursor='hand2')
    def choose(self,v): self.result=v; self.destroy()


class EnterpriseStatusDialog(tk.Toplevel):
    def __init__(self,parent,d,judged,for_excel=True,for_weekly=True):
        super().__init__(parent); self.withdraw(); self.result=None; self.title('상태 판단 최종 확인'); self.configure(bg='white'); self.resizable(False,False); self.transient(parent)
        judged='close' if str(judged).lower()=='close' else 'open'
        weekly=weekly_status_from_choice(d,judged)
        self.db_var=tk.StringVar(value=judged)
        self.weekly_var=tk.StringVar(value=weekly)
        self.db_confirm=tk.BooleanVar(value=False)
        self.weekly_confirm=tk.BooleanVar(value=False)
        self.for_excel=bool(for_excel); self.for_weekly=bool(for_weekly)

        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='상태 판단 최종 확인',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
        body=tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=16)
        tk.Label(body,text='주간회의와 Issue DB의 상태를 각각 확인해 주세요. 상태를 변경한 뒤에도 각 항목의 확인 체크가 필요합니다.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold'),wraplength=790,justify='left').pack(anchor='w',pady=(0,10))

        if self.for_weekly:
            self._weekly_card(body,d,weekly)
        if self.for_excel:
            self._db_card(body,d,judged)

        foot=tk.Frame(self,bg='#F6F8FA',height=70); foot.pack(fill='x'); foot.pack_propagate(False)
        b=tk.Frame(foot,bg='#F6F8FA'); b.pack(side='right',padx=20,pady=13)
        tk.Button(b,text='취소',command=self.cancel,width=12,bd=0,font=('Malgun Gothic',9,'bold'),bg='#E5EBF0',fg=ui.TEXT,pady=8,cursor='hand2').pack(side='left',padx=4)
        self.ok_btn=tk.Button(b,text='상태 확인 완료',command=self.ok,width=16,bd=0,font=('Malgun Gothic',9,'bold'),bg='#AAB7C2',fg='white',pady=8,cursor='arrow',state='disabled')
        self.ok_btn.pack(side='left',padx=4)

        if self.for_excel: self.db_confirm.trace_add('write',lambda *_:self._refresh_ok())
        if self.for_weekly: self.weekly_confirm.trace_add('write',lambda *_:self._refresh_ok())
        self.protocol('WM_DELETE_WINDOW',self.cancel)
        height=690 if self.for_excel and self.for_weekly else 480
        ui.center_window(self,parent,840,height); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)

    def _card_shell(self,parent,title,auto_status,reason):
        card=tk.Frame(parent,bg='#F8FAFC',highlightbackground='#D5E3EE',highlightthickness=1)
        card.pack(fill='x',pady=(0,10))
        tk.Label(card,text=title,bg='#F8FAFC',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(anchor='w',padx=14,pady=(10,2))
        tk.Label(card,text=f'추천  {auto_status}',bg='#EEF5FA',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(anchor='w',padx=14,pady=(7,2))
        tk.Label(card,text='추천 이유',bg='#F8FAFC',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w',padx=14,pady=(5,1))
        tk.Label(card,text=reason,bg='#F8FAFC',fg='#4E6375',font=('Malgun Gothic',8),wraplength=760,justify='left').pack(anchor='w',padx=14,pady=(0,8))
        return card

    def _weekly_card(self,parent,d,auto_status):
        card=self._card_shell(parent,'주간회의 Signal',auto_status,weekly_status_reason(d,auto_status))
        row=tk.Frame(card,bg='#F8FAFC'); row.pack(fill='x',padx=14,pady=(0,7))
        for val in ('원인/개선 미확인','개선 검증중','개선 완료'):
            ttk.Radiobutton(row,text=val,variable=self.weekly_var,value=val,style='Mode.TRadiobutton',command=lambda:self.weekly_confirm.set(False)).pack(side='left',padx=(0,18))
        tk.Checkbutton(card,text='주간회의 상태를 이 선택으로 확인',variable=self.weekly_confirm,command=self._refresh_ok,bg='#F8FAFC',fg=ui.TEXT,activebackground='#F8FAFC',font=('Malgun Gothic',9,'bold')).pack(anchor='w',padx=12,pady=(0,10))

    def _db_card(self,parent,d,auto_status):
        card=self._card_shell(parent,'Issue DB 상태',auto_status,issue_db_status_reason(d,auto_status))
        row=tk.Frame(card,bg='#F8FAFC'); row.pack(fill='x',padx=14,pady=(0,7))
        ttk.Radiobutton(row,text='open  ·  진행/검증 필요',variable=self.db_var,value='open',style='Mode.TRadiobutton',command=lambda:self.db_confirm.set(False)).pack(side='left',padx=(0,24))
        ttk.Radiobutton(row,text='close  ·  개선 완료',variable=self.db_var,value='close',style='Mode.TRadiobutton',command=lambda:self.db_confirm.set(False)).pack(side='left')
        tk.Checkbutton(card,text='Issue DB 상태를 이 선택으로 확인',variable=self.db_confirm,command=self._refresh_ok,bg='#F8FAFC',fg=ui.TEXT,activebackground='#F8FAFC',font=('Malgun Gothic',9,'bold')).pack(anchor='w',padx=12,pady=(0,10))

    def _refresh_ok(self):
        ready=(not self.for_excel or self.db_confirm.get()) and (not self.for_weekly or self.weekly_confirm.get())
        if ready:
            self.ok_btn.configure(state='normal',bg=ui.BLUE,cursor='hand2')
        else:
            self.ok_btn.configure(state='disabled',bg='#AAB7C2',cursor='arrow')

    def ok(self):
        if not ((not self.for_excel or self.db_confirm.get()) and (not self.for_weekly or self.weekly_confirm.get())):
            return
        self.result={
            'issue_db': self.db_var.get() if self.for_excel else None,
            'weekly': self.weekly_var.get() if self.for_weekly else None,
        }
        self.destroy()

    def cancel(self): self.result=None; self.destroy()


class EnterpriseApp(legacy.FinalApp, _RootBase):
    def __init__(self):
        _RootBase.__init__(self); self.title('8D Issue Automation 1.0 | Pack 개발품질'); self.geometry('1180x850'); self.minsize(1080,780); self.configure(bg=ui.BG)
        self.vars={}; self.mode=tk.StringVar(value='existing'); self.status_var=tk.StringVar(value='READY  ·  8D 원본을 선택해 주세요.')
        self._configure_styles(); self._build_enterprise_ui(); self._center_main()
    def _center_main(self):
        self.update_idletasks(); w=1180; h=850; self.geometry(f'{w}x{h}+{max(0,(self.winfo_screenwidth()-w)//2)}+{max(0,(self.winfo_screenheight()-h)//2)}')
    def _build_enterprise_ui(self):
        head=tk.Frame(self,bg=ui.NAVY,height=96); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='8D Issue Automation',bg=ui.NAVY,fg='white',font=('Malgun Gothic',20,'bold')).place(x=32,y=18); tk.Label(head,text='Issue Management Productivity Tool  ·  v1.0',bg=ui.NAVY,fg='#C9D8E6',font=('Malgun Gothic',9)).place(x=34,y=58); tk.Label(head,text='PACK DEVELOPMENT QUALITY',bg=ui.NAVY,fg='#D7E4EF',font=('Segoe UI',8,'bold')).place(relx=1,x=-32,y=38,anchor='e')
        body=tk.Frame(self,bg=ui.BG); body.pack(fill='both',expand=True,padx=26,pady=18)
        mode=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); mode.pack(fill='x',pady=(0,12)); tk.Label(mode,text='이슈 처리 방식',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left',padx=(18,24),pady=13)
        for text,value in [('기존 이슈 업데이트','existing'),('신규 이슈 등록','new')]: ttk.Radiobutton(mode,text=text,variable=self.mode,value=value,style='Mode.TRadiobutton').pack(side='left',padx=(0,20))
        tk.Label(mode,text='기존 이슈는 동일 이슈 행/페이지를 찾아 업데이트합니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='right',padx=18)
        content=tk.Frame(body,bg=ui.BG); content.pack(fill='x'); left=tk.Frame(content,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); left.pack(side='left',fill='both',expand=True,padx=(0,7)); right=tk.Frame(content,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); right.pack(side='left',fill='both',expand=True,padx=(7,0))
        self._section_title(left,'01','입력 파일','Drag & Drop 지원'); filebody=tk.Frame(left,bg='white'); filebody.pack(fill='x',padx=16,pady=(0,14))
        for key in ('ppt8d','pptweekly','xlsx'): self.vars[key]=tk.StringVar()
        self.dropzones={}; specs=[('8D 원본 PPT','ppt8d',('.pptx',),True),('주간회의 PPT','pptweekly',('.pptx',),False),('Issue DB Excel','xlsx',('.xlsx',),False)]
        for title,key,exts,req in specs:
            z=ui.DropZone(
                filebody,title,key,self.vars[key],exts,self.pick,
                lambda k=key:self._file_selection_changed(k),req
            )
            z.pack(fill='x',pady=5); self.dropzones[key]=z
        tk.Label(filebody,text=f"Drag & Drop: {'사용 가능' if ui.DND_AVAILABLE else '미사용 · [찾기] 버튼 사용'}  |  8D는 필수, 나머지는 선택 입력",bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',pady=(4,0))
        self._section_title(right,'02','담당 및 분류 정보','표준 입력'); fields=tk.Frame(right,bg='white'); fields.pack(fill='x',padx=18,pady=(0,10))
        for key in ('team','task_name','owner','sample','plm_no'): self.vars[key]=tk.StringVar()
        self._enterprise_entry(fields,'담당팀','team','예: Pack개발품질1팀'); self._enterprise_entry(fields,'고객사/과제명','task_name','담당팀 연계 과제 선택'); self._enterprise_entry(fields,'담당자','owner','예: 홍길동'); self._enterprise_entry(fields,'발생 샘플','sample','예: DUT3 / Sample No.'); self._enterprise_entry(fields,'PMS/PLM 이슈번호','plm_no','기존값이 있을 때 입력'); tk.Frame(fields,bg=ui.BORDER,height=1).pack(fill='x',pady=10)
        self.vars['form_factor']=tk.StringVar(value='파우치형'); self.vars['product_type']=tk.StringVar(value='EV Pack'); self.vars['occurrence_site']=tk.StringVar(value=legacy.OCCURRENCE_SITES[0]); self.vars['stage']=tk.StringVar(value='DV')
        self._enterprise_combo(fields,'폼팩터','form_factor',legacy.FORM_FACTORS); self._enterprise_combo(fields,'제품 타입','product_type',legacy.PRODUCT_TYPES); self._enterprise_combo(fields,'발생처','occurrence_site',legacy.OCCURRENCE_SITES); self._enterprise_combo(fields,'개발 단계','stage',legacy.STAGES)
        target=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); target.pack(fill='x',pady=12); tk.Label(target,text='UPDATE TARGET',bg='white',fg=ui.MUTED,font=('Segoe UI',8,'bold')).pack(side='left',padx=(16,12),pady=12); self.target_label=tk.Label(target,text='',bg='white',fg=ui.NAVY,font=('Malgun Gothic',9,'bold')); self.target_label.pack(side='left'); self.target_weekly=tk.Label(target,text='',bg='#EDF1F4',font=('Segoe UI',8,'bold'),padx=9,pady=4); self.target_weekly.pack(side='right',padx=(5,16)); self.target_excel=tk.Label(target,text='',bg='#EDF1F4',font=('Segoe UI',8,'bold'),padx=9,pady=4); self.target_excel.pack(side='right')
        actions=tk.Frame(body,bg=ui.BG); actions.pack(fill='x',pady=(0,10)); tk.Button(actions,text='8D 내용 미리보기',command=self.preview,bg='#E5EBF0',fg=ui.NAVY,bd=0,font=('Malgun Gothic',9,'bold'),padx=18,pady=9,cursor='hand2').pack(side='left'); tk.Button(actions,text='자동 업데이트 실행  →',command=self.run,bg=ui.BLUE,fg='white',bd=0,font=('Malgun Gothic',10,'bold'),padx=24,pady=10,cursor='hand2').pack(side='right')
        result=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); result.pack(fill='both',expand=True); rt=tk.Frame(result,bg='white'); rt.pack(fill='x',padx=16,pady=(12,7)); tk.Label(rt,text='실행 결과',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left'); tk.Label(rt,textvariable=self.status_var,bg='#E9F1F8',fg='#315A7D',font=('Malgun Gothic',8),padx=9,pady=4).pack(side='right'); self.log=tk.Text(result,height=7,wrap='word',font=('Consolas',9),bg='#FBFCFD',fg='#283A49',relief='flat',highlightthickness=1,highlightbackground='#E1E7EC',padx=10,pady=8); self.log.pack(fill='both',expand=True,padx=16,pady=(0,14)); self.log.insert('end','8D 원본을 Drag & Drop하거나 [찾기]로 선택해 주세요.\n')
        foot=tk.Frame(self,bg='#E8EEF3',height=30); foot.pack(fill='x'); foot.pack_propagate(False); tk.Label(foot,text='8D Issue Automation 1.0   ·   원본 파일은 직접 덮어쓰지 않습니다.',bg='#E8EEF3',fg='#607180',font=('Malgun Gothic',8)).pack(side='left',padx=26); self._refresh_target()
    def _section_title(self,parent,no,title,sub):
        f=tk.Frame(parent,bg='white'); f.pack(fill='x',padx=18,pady=(15,10)); tk.Label(f,text=no,bg=ui.NAVY,fg='white',font=('Segoe UI',8,'bold'),width=3,pady=3).pack(side='left'); tk.Label(f,text=title,bg='white',fg=ui.NAVY,font=('Malgun Gothic',11,'bold')).pack(side='left',padx=9); tk.Label(f,text=sub,bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='right')
    def _enterprise_entry(self,parent,label,key,hint):
        row=tk.Frame(parent,bg='white'); row.pack(fill='x',pady=4)
        tk.Label(row,text=label,bg='white',fg=ui.TEXT,font=('Malgun Gothic',9),width=17,anchor='w').pack(side='left')
        wrap=tk.Frame(row,bg='white'); wrap.pack(side='left',fill='x',expand=True)
        ttk.Entry(wrap,textvariable=self.vars[key]).pack(fill='x')
        tk.Label(wrap,text=hint,bg='white',fg='#98A3AD',font=('Malgun Gothic',7),anchor='e').pack(fill='x',anchor='e')
    def _enterprise_combo(self,parent,label,key,values):
        row=tk.Frame(parent,bg='white'); row.pack(fill='x',pady=5); tk.Label(row,text=label,bg='white',fg=ui.TEXT,font=('Malgun Gothic',9),width=17,anchor='w').pack(side='left'); ttk.Combobox(row,textvariable=self.vars[key],values=values,state='readonly').pack(side='left',fill='x',expand=True)
    def _file_selection_changed(self,key):
        """Synchronize target chips and the visible status immediately after file changes."""
        self._refresh_target()
        if key!='ppt8d':
            return
        path=(self.vars.get('ppt8d').get().strip() if self.vars.get('ppt8d') else '')
        if path:
            msg='READY  ·  8D 원본 선택 완료 · 미리보기 가능'
            self.status_var.set(msg)
            # V3 uses a display-only status variable; sync it immediately as well.
            try:self._sync_status_display()
            except Exception:pass
            # Keep the large result/message area consistent with the small status badge.
            try:
                self.log.delete('1.0','end')
                self.log.insert('end','[ 입력 준비 ]\n'+'─'*72+'\n')
                self.log.insert('end',f'8D 원본 선택 완료 : {Path(path).name}\n')
                self.log.insert('end','상태              : 미리보기 또는 자동 업데이트 실행 가능\n')
            except Exception:
                pass
        else:
            self.status_var.set('READY  ·  8D 원본을 선택해 주세요.')
            try:self._sync_status_display()
            except Exception:pass
            try:
                self.log.delete('1.0','end')
                self.log.insert('end','8D 원본을 Drag & Drop하거나 [찾기]로 선택해 주세요.\n')
            except Exception:
                pass

    def pick(self,k,desc=None,pattern=None):
        types=[('Excel (*.xlsx)','*.xlsx')] if k=='xlsx' else [('PowerPoint (*.pptx)','*.pptx')]
        f=filedialog.askopenfilename(title='파일 선택',filetypes=types+[('모든 파일','*.*')],parent=self)
        if f:self.vars[k].set(f)
    def _refresh_target(self):
        if not hasattr(self,'target_label'): return
        ppt8d=(self.vars.get('ppt8d').get().strip() if self.vars.get('ppt8d') else '')
        weekly=bool(self.vars.get('pptweekly') and self.vars['pptweekly'].get().strip()); excel=bool(self.vars.get('xlsx') and self.vars['xlsx'].get().strip()); text='주간회의 PPT + Issue DB Excel' if weekly and excel else '주간회의 PPT만 업데이트' if weekly else 'Issue DB Excel만 업데이트' if excel else '주간회의 또는 Issue DB를 선택해 주세요.'; self.target_label.configure(text=text); self.target_weekly.configure(text='WEEKLY  ON' if weekly else 'WEEKLY  OFF',bg='#E7F4ED' if weekly else '#EDF1F4',fg=ui.GREEN if weekly else ui.MUTED); self.target_excel.configure(text='ISSUE DB  ON' if excel else 'ISSUE DB  OFF',bg='#E7F4ED' if excel else '#EDF1F4',fg=ui.GREEN if excel else ui.MUTED)
        current=self.status_var.get().strip()
        refreshed=target_ready_status(ppt8d,current)
        if refreshed!=current:self.status_var.set(refreshed)

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d'): return ui.warning(self,'입력 확인','8D 원본 PPT를 선택해 주세요.')
        try:
            self.status_var.set('ANALYZING  ·  8D 내용을 추출하고 있습니다...'); self.update_idletasks(); d=base.extract(g['ppt8d']); labels=[('이슈명','issue_name'),('과제명','task_name'),('고객사','customer'),('발생일자','occurrence_date'),('2D 현상','problem'),('3D 임시조치','temporary_action'),('4D 발생원인','cause_4d'),('4D 유출원인','leak_cause'),('5D 개선대책','action_5d'),('6D 효과검증','verification_6d')]; self.log.delete('1.0','end'); self.log.insert('end','[ 8D 추출 결과 ]\n'+'─'*72+'\n')
            for label,key in labels: self.log.insert('end',f'{label:<12} : {N(d.get(key)) or "-"}\n')
            self.status_var.set('READY  ·  8D 추출 완료')
        except Exception as e: self.status_var.set('ERROR  ·  추출 실패'); ui.error(self,'미리보기 오류',repr(e))

    def _confirm_issue_db_status_v1(self,d):
        """V1-style custom confirmation window; shown once whenever 5D or 6D exists."""
        if not issue_db_status_confirmation_required(d):
            return 'open'

        recommended,_reason=issue_db_recommended_status(d)
        other='open' if recommended=='close' else 'close'
        five=N(d.get('action_5d')) or '(5D 내용 없음)'
        six=N(d.get('verification_6d')) or '(6D 내용 없음)'
        result={'value':None}

        dlg=tk.Toplevel(self)
        dlg.withdraw()
        dlg.title('이슈 상태 확인')
        dlg.configure(bg='white')
        dlg.resizable(False,False)
        dlg.transient(self)

        head=tk.Frame(dlg,bg=ui.NAVY,height=72)
        head.pack(fill='x')
        head.pack_propagate(False)
        tk.Label(
            head,text='이슈 상태 확인',
            bg=ui.NAVY,fg='white',
            font=('Malgun Gothic',16,'bold')
        ).pack(side='left',padx=30)

        body=tk.Frame(dlg,bg='white')
        body.pack(fill='both',expand=True,padx=34,pady=24)

        top=tk.Frame(body,bg='white')
        top.pack(fill='x',pady=(0,12))
        tk.Label(
            top,text='?',width=2,height=1,
            bg=ui.BLUE,fg='white',
            font=('Malgun Gothic',17,'bold')
        ).pack(side='left',anchor='n',padx=(0,18))
        tk.Label(
            top,
            text=f'5D/6D 내용 기준으로 {recommended}로 판단되었습니다.',
            bg='white',fg=ui.TEXT,
            font=('Malgun Gothic',12,'bold'),
            justify='left'
        ).pack(side='left',anchor='n',pady=4)

        content=tk.Frame(body,bg='white')
        content.pack(fill='both',expand=True,padx=(58,0))

        tk.Label(
            content,text='판단 이유 [5D/6D 원문]',
            bg='white',fg=ui.TEXT,
            font=('Malgun Gothic',10,'bold')
        ).pack(anchor='w',pady=(0,6))

        tk.Frame(content,bg='#8E9AA5',height=1).pack(fill='x',pady=(0,7))

        six_wrap=tk.Frame(content,bg='white')
        six_wrap.pack(fill='x')
        six_box=tk.Text(
            six_wrap,height=7,wrap='word',
            bg='white',fg=ui.TEXT,
            relief='flat',bd=0,
            font=('Malgun Gothic',10),
            padx=2,pady=2
        )
        six_scroll=tk.Scrollbar(six_wrap,command=six_box.yview)
        six_box.configure(yscrollcommand=six_scroll.set)
        six_scroll.pack(side='right',fill='y')
        six_box.pack(side='left',fill='both',expand=True)
        six_box.insert('1.0',f'[5D 개선대책]\n{five}\n\n[6D 효과검증]\n{six}')
        six_box.configure(state='disabled')

        tk.Frame(content,bg='#8E9AA5',height=1).pack(fill='x',pady=(7,12))

        def show_full_preview():
            preview=tk.Toplevel(dlg)
            preview.withdraw()
            preview.title('8D 전체내용 확인')
            preview.configure(bg='white')
            preview.resizable(True,True)
            preview.transient(dlg)

            ph=tk.Frame(preview,bg=ui.NAVY,height=60)
            ph.pack(fill='x')
            ph.pack_propagate(False)
            tk.Label(
                ph,text='8D 전체내용 확인',
                bg=ui.NAVY,fg='white',
                font=('Malgun Gothic',12,'bold')
            ).pack(side='left',padx=22)

            pb=tk.Frame(preview,bg='white')
            pb.pack(fill='both',expand=True,padx=22,pady=18)

            wrap=tk.Frame(pb,bg='white')
            wrap.pack(fill='both',expand=True)
            box=tk.Text(
                wrap,wrap='word',
                bg='#F7F9FB',fg=ui.TEXT,
                relief='flat',
                highlightthickness=1,
                highlightbackground=ui.BORDER,
                font=('Malgun Gothic',9),
                padx=12,pady=10
            )
            sb=tk.Scrollbar(wrap,command=box.yview)
            box.configure(yscrollcommand=sb.set)
            sb.pack(side='right',fill='y')
            box.pack(side='left',fill='both',expand=True)

            labels=[
                ('이슈명','issue_name'),
                ('고객사','customer'),
                ('고객사/과제명','task_name'),
                ('발생일자','occurrence_date'),
                ('2D 현상','problem'),
                ('3D 임시조치','temporary_action'),
                ('4D 발생원인','cause_4d'),
                ('4D 유출원인','leak_cause'),
                ('4D 시스템원인','system_cause'),
                ('5D 개선대책','action_5d'),
                ('6D 효과검증','verification_6d'),
            ]
            box.insert('end','[ 8D 추출 전체내용 ]\n'+'─'*70+'\n')
            for label,key in labels:
                box.insert('end',f'\n{label}\n')
                box.insert('end',f'{N(d.get(key)) or "-"}\n')
            box.configure(state='disabled')

            def close_preview():
                try:preview.grab_release()
                except Exception:pass
                preview.destroy()
                try:
                    dlg.grab_set()
                    dlg.focus_force()
                except Exception:
                    pass

            pf=tk.Frame(preview,bg='#F6F8FA',height=62)
            pf.pack(fill='x')
            pf.pack_propagate(False)
            tk.Button(
                pf,text='닫기',command=close_preview,
                width=12,bd=0,
                font=('Malgun Gothic',9,'bold'),
                bg=ui.BLUE,fg='white',
                pady=7,cursor='hand2'
            ).pack(side='right',padx=20,pady=13)

            preview.protocol('WM_DELETE_WINDOW',close_preview)
            ui.center_window(preview,dlg,760,620)
            preview.deiconify()
            preview.grab_set()
            preview.focus_force()

        tk.Button(
            content,
            text='5D 포함 전체내용 확인하기',
            command=show_full_preview,
            bd=0,
            bg='#E5EBF0',fg=ui.TEXT,
            activebackground='#DDE5EB',
            font=('Malgun Gothic',9,'bold'),
            padx=14,pady=7,
            cursor='hand2'
        ).pack(anchor='w',pady=(0,18))

        tk.Label(
            content,
            text=(
                f'이슈 상태를 {recommended}로 처리하시겠습니까?\n'
                f'아니오를 선택하면 {other}으로 처리합니다.'
            ),
            bg='white',fg=ui.TEXT,
            font=('Malgun Gothic',10),
            justify='left'
        ).pack(anchor='w')

        foot=tk.Frame(dlg,bg='#F6F8FA',height=72)
        foot.pack(fill='x')
        foot.pack_propagate(False)
        btns=tk.Frame(foot,bg='#F6F8FA')
        btns.pack(side='right',padx=22,pady=14)

        def choose(value):
            result['value']=value
            try:dlg.grab_release()
            except Exception:pass
            dlg.destroy()

        tk.Button(
            btns,text='아니오',
            command=lambda:choose(other),
            width=12,bd=0,
            font=('Malgun Gothic',9,'bold'),
            bg='#E5EBF0',fg=ui.TEXT,
            pady=8,cursor='hand2'
        ).pack(side='left',padx=5)

        tk.Button(
            btns,text='확인',
            command=lambda:choose(recommended),
            width=12,bd=0,
            font=('Malgun Gothic',9,'bold'),
            bg=ui.BLUE,fg='white',
            pady=8,cursor='hand2'
        ).pack(side='left',padx=5)

        dlg.protocol('WM_DELETE_WINDOW',lambda:choose(None))
        ui.center_window(dlg,self,720,570)
        dlg.deiconify()
        dlg.grab_set()
        dlg.focus_force()
        self.wait_window(dlg)
        return result['value']

    def _confirm_weekly_status_v1(self,d):
        """Show a separate weekly Signal chooser only when the Issue DB remains open."""
        recommended=weekly_recommended_status(d)
        dlg=EnterpriseWeeklyStatusDialog(self,d,recommended)
        return dlg.result


    def _confirm_weekly_section(self,weekly,d,g):
        """Resolve weekly section safely and ask before using any fallback area."""
        try:
            import main_recovery_step14_fix2 as weekly_core
            prs=Presentation(weekly)
            resolved=weekly_core.section_resolution(prs,d)
        except Exception:
            # If section inspection itself fails, prefer a new section over a guessed existing area.
            g['_weekly_create_new_section']='1'
            return True

        mode=resolved.get('mode')
        if mode=='exact':
            if resolved.get('name'):g['_weekly_section_override_name']=resolved.get('name')
            return True
        if mode=='new':
            g['_weekly_create_new_section']='1'
            return True

        name=resolved.get('name') or '확인된 구역'
        reason=resolved.get('reason') or '유사한 구역이 확인되었습니다.'
        msg=(
            f'추천  {name}\n\n'
            f'{reason}\n\n'
            '이 구역에 업데이트할지, 새 구역을 만들어 업데이트할지 선택해 주세요.'
        )
        choice=ui.dialog(
            self,'주간회의 구역 확인',msg,'question',
            (('새 구역 생성','new'),('해당 구역 업데이트','use')),
            width=620,height=340,button_width=14
        )
        if choice is None:
            return False
        if choice=='use':
            if resolved.get('section') is not None and resolved.get('name'):
                g['_weekly_section_override_name']=resolved.get('name')
            elif resolved.get('insert_after') is not None:
                g['_weekly_insert_after_override']=str(resolved.get('insert_after'))
            return True

        g['_weekly_create_new_section']='1'
        return True

    def run(self):
        g=self.gui(); ppt8d=g.get('ppt8d',''); weekly=g.get('pptweekly',''); xlsx=g.get('xlsx','')
        if not ppt8d or not os.path.exists(ppt8d): return ui.warning(self,'입력 확인','8D 원본 PPT는 반드시 선택해 주세요.')
        if not weekly and not xlsx: return ui.warning(self,'출력 확인','주간회의 PPT 또는 Issue DB Excel 중 하나 이상을 선택해 주세요.')
        if weekly and not os.path.exists(weekly): return ui.warning(self,'입력 확인','선택한 주간회의 PPT 파일을 찾을 수 없습니다.')
        if xlsx and not os.path.exists(xlsx): return ui.warning(self,'입력 확인','선택한 Issue DB Excel 파일을 찾을 수 없습니다.')
        if g.get('occurrence_site') not in legacy.OCCURRENCE_SITES: return ui.warning(self,'입력 확인','발생처를 선택해 주세요.')
        try:
            self.status_var.set('RUNNING  ·  8D 원본 분석 중...'); self.update_idletasks(); d=base.extract(ppt8d); selected_task=g.get('task_name','').strip()
            if selected_task:
                mismatch,extracted_task,entered_task=customer_project_mismatch(d,selected_task)
                if mismatch:
                    msg=(
                        '8D에서 확인된 고객사/과제명과 입력값이 다릅니다.\n\n'
                        f'8D 확인값     : {extracted_task}\n'
                        f'입력/선택값   : {entered_task}\n\n'
                        '입력/선택한 고객사/과제명으로 계속 진행할까요?'
                    )
                    choice=ui.dialog(
                        self,'고객사/과제명 확인',msg,'warning',
                        (('취소',False),('입력값 사용',True)),
                        width=620,height=330,button_width=13
                    )
                    if choice is not True:
                        self.status_var.set('READY  ·  고객사/과제명 확인이 취소되었습니다.'); return
                d=apply_selected_customer_project(d,selected_task)
            mode=self.mode.get(); do_weekly=bool(weekly); do_excel=bool(xlsx); excel_row=None
            if do_weekly:
                if not self._confirm_weekly_section(weekly,d,g):
                    self.status_var.set('READY  ·  주간회의 구역 선택이 취소되었습니다.'); return
            if do_excel:
                self.status_var.set('RUNNING  ·  Issue DB 대상 확인 중...'); self.update_idletasks(); wb=load_workbook(xlsx); ws=wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active; excel_row,_=base.find(ws,d,g)
                if mode=='existing' and not excel_row:
                    if do_weekly:
                        if ui.ask_yes_no(self,'기존 이슈 미확인','Issue DB에서 동일 이슈를 찾지 못했습니다.\n\nExcel은 변경하지 않고 주간회의 PPT만 업데이트할까요?'): do_excel=False
                        else: self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
                    else: self.status_var.set('READY  ·  기존 이슈 미확인'); return ui.warning(self,'기존 이슈 미확인','Issue DB에서 동일 이슈를 찾지 못했습니다.\n기존 이슈는 임의로 신규 행을 만들지 않습니다.')
                if mode=='new' and excel_row and not ui.ask_yes_no(self,'중복 가능성',f'유사 이슈(row {excel_row})가 있습니다.\n그래도 신규 이슈로 추가할까요?'): self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
            od=None
            if do_weekly:
                origin_d=dict(d)
                origin_d['_origin_occurrence_site']=g.get('occurrence_site') or d.get('occurrence_site')
                self.status_var.set('WAITING  ·  이슈기인 확인 필요'); self.update_idletasks(); od=EnterpriseOriginDialog(self,origin_d)
                if od.result is None: self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
                g['_issue_origin_selected']=od.result
            pd_result=None; db_status=None; weekly_status=None
            if do_excel:
                summary=legacy.step7._db_summary(d); self.status_var.set('WAITING  ·  Issue DB 현상 입력 확인 필요'); self.update_idletasks(); pd=EnterpriseProblemDialog(self,summary)
                if pd.result is None: self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
                pd_result=pd.result; g['_db_problem_selected']=summary if pd.result=='summary' else N(d.get('problem'))
            if do_excel:
                self.status_var.set('WAITING  ·  Issue DB 상태 확인 필요'); self.update_idletasks()
                db_status=self._confirm_issue_db_status_v1(d)
                if db_status is None:
                    self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
                g['_issue_status_selected']=db_status
            if do_weekly:
                # Show the weekly chooser ONLY when Issue DB was actually updated and
                # its final user-confirmed status is open.
                if do_excel and weekly_status_confirmation_required(db_status):
                    self.status_var.set('WAITING  ·  주간회의 상태 확인 필요'); self.update_idletasks()
                    weekly_status=self._confirm_weekly_status_v1(d)
                    if weekly_status is None:
                        self.status_var.set('READY  ·  사용자가 실행을 취소했습니다.'); return
                elif do_excel and str(db_status or '').lower()=='close':
                    weekly_status='개선 완료'
                else:
                    # Weekly-only execution has no "final Issue DB" state, so no extra
                    # confirmation popup is shown; use the conservative automatic Signal.
                    judged,_reason=legacy.step9._judge_issue_status(d)
                    weekly_status=weekly_recommended_status(d)
                g['_weekly_status_selected']=weekly_status
            anchor=xlsx if do_excel else weekly; out=Path(anchor).parent/'자동화_결과'; out.mkdir(exist_ok=True); results=[]
            self._last_saved_outputs={}
            if do_excel:
                self.status_var.set('RUNNING  ·  Issue DB 업데이트 중...'); self.update_idletasks()
                xo=out/(Path(xlsx).stem+'_업데이트.xlsx')
                a,xsaved=base.update_excel(xlsx,xo,d,g,new=(mode=='new'))
                self._last_saved_outputs['excel']=str(xsaved)
                mrow=re.search(r'row\s+(\d+)',str(a),re.I)
                self._last_excel_updated_row=int(mrow.group(1)) if mrow else None
                results.append(a)
            if do_weekly:
                self.status_var.set('RUNNING  ·  주간회의 PPT 업데이트 중...'); self.update_idletasks()
                po=out/(Path(weekly).stem+'_업데이트.pptx')
                b,psaved=base.weekly(weekly,po,d,g,mode)
                self._last_saved_outputs['weekly']=str(psaved); results.append(b)
            self.log.delete('1.0','end'); targets=[]
            if do_excel: targets.append('Issue DB Excel')
            if do_weekly: targets.append('주간회의 PPT')
            lines=['[ 업데이트 완료 ]','─'*72,f'이슈 구분       : {"신규 이슈" if mode=="new" else "기존 이슈"}',f'업데이트 대상   : {" + ".join(targets)}'];
            if do_weekly and od: lines.append(f'이슈기인        : {od.result}')
            if do_weekly: lines.append(f'주간회의 상태   : {weekly_status}')
            if do_excel: lines.extend([f'Issue DB 현상   : {"현상 요약" if pd_result=="summary" else "전체 내용"}',f'Issue DB 상태   : {db_status}'])
            lines.extend(['']+results+['',f'결과 폴더       : {out}']); self.log.insert('end','\n'.join(lines)); self.status_var.set('COMPLETE  ·  선택한 자료의 업데이트가 완료되었습니다.'); ui.info(self,'업데이트 완료','선택한 자료만 업데이트했습니다.\n\n'+'\n'.join('✓ '+x for x in targets)+f'\n\n결과 폴더\n{out}')
        except Exception as e: self.status_var.set('ERROR  ·  실행 중 오류가 발생했습니다.'); ui.error(self,'실행 오류',repr(e))


if __name__=='__main__': EnterpriseApp().mainloop()
