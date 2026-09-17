"""Editable Issue DB / weekly problem-summary workflow."""
import re
import tkinter as tk
import main_enterprise_v3 as v3
import main_enterprise as ent
import ui_enterprise as ui

legacy=ent.legacy; N=ent.N; MAX_LEN=150
_original_run=v3.EnterpriseAppV3.run
_original_weekly=ent.base.weekly
_original_update_excel=ent.base.update_excel
_original_summary_update=legacy.s14._update_summary_by_task
_active_summary=""; _active_db_choice=None; _active_weekly_choice=None; _active_weekly_summary=""

def _norm(s): return re.sub(r"\s+"," ",N(s)).strip()
def _db_clean(s):
    s=N(s)
    s=re.sub(r"(?im)^\s*[•·\-]?\s*\d+[.)]?\s*현상\s*[:：]\s*","",s)
    s=re.sub(r"(?im)^\s*현상\s*[:：]\s*","",s)
    return s.strip()

def _context_header(src):
    flat=_norm(src)
    rules=[
        ("시험명",r"([A-Za-z0-9#._/+\-가-힣 ]{1,45}?(?:시험|평가|검사))(?=\s*(?:진행\s*)?중|에서|시|,|\.|$)"),
        ("빌드명",r"([A-Za-z0-9#._/+\-가-힣 ]{1,45}?빌드)(?=\s*(?:진행\s*)?중|에서|시|,|\.|$)"),
        ("공정명",r"([A-Za-z0-9#._/+\-가-힣 ]{1,45}?공정)(?=\s*(?:진행\s*)?중|에서|시|,|\.|$)"),
        ("부품명",r"([A-Za-z0-9#._/+\-가-힣 ]{1,45}?부품)(?=\s*(?:이슈|불량|에서|,|\.|$))"),
    ]
    for label,pat in rules:
        m=re.search(pat,flat,re.I)
        if m:
            name=_norm(m.group(1)); name=re.sub(r"^(?:1[.)]?\s*)?현상\s*[:：]?\s*","",name)
            if name:return f"• {label} : {name}"
    return ""

def _compact_problem(d):
    src=_db_clean(d.get("problem"))
    if not src:return ""
    header=_context_header(src)
    try:symptom_lines=[N(x) for x in legacy.step3._actual_symptoms(d) if N(x)]
    except Exception:symptom_lines=[]
    chunks=[x.strip(" -•·\t") for x in re.split(r"[\n\r]+",src) if x.strip()]
    context_words=("시험","평가","검사","빌드","공정","부품")
    context_lines=[x for x in chunks if any(w in x.casefold() for w in context_words)]
    cleaned=[]; seen=[]
    for raw in context_lines+symptom_lines+chunks:
        s=_db_clean(_norm(raw))
        s=re.sub(r"\b진행\s*중\b","중",s)
        s=re.sub(r"(?:,?\s*)?(?:재)?확인\s*결과\s*(?:동일하게\s*)?"," ",s)
        s=re.sub(r"(?:하는\s*)?현상이\s*(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$","",s)
        s=re.sub(r"(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$","",s)
        s=re.sub(r"\s+"," ",s).strip(" ,.;·-")
        key=re.sub(r"[\s,.;:·/\\_-]+","",s).casefold()
        if not s or key in seen:continue
        if any(key and (key in old or old in key) for old in seen):continue
        seen.append(key); cleaned.append(s)
    # Never cut a sentence/line merely to hit the character limit. Add only whole
    # candidate statements that fit; if the next one would overflow, omit it.
    selected=[]
    prefix=(header+"\n") if header else ""
    used=len(prefix)
    for part in cleaned:
        sep=3 if selected else 0
        if used+sep+len(part)<=MAX_LEN:
            selected.append(part); used+=sep+len(part)
        else:
            continue
    if selected:return (prefix+" / ".join(selected)).strip()
    if header:return header[:MAX_LEN]
    # One source statement can itself exceed 150 chars. Do not return a broken
    # fragment; leave the automatic suggestion empty so the user can edit from
    # the full source shown in the dialog.
    return ""

class EditableProblemDialog(tk.Toplevel):
    def __init__(self,parent,summary,full_text="",title="이슈 DB 현상 요약 확인"):
        super().__init__(parent); self.withdraw(); self.result=None; self.summary=None; self.title(title); self.configure(bg="white"); self.transient(parent)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill="x"); head.pack_propagate(False); tk.Label(head,text=title,bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold")).pack(side="left",padx=22)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=16)
        tk.Label(body,text="이슈 DB의 현상 칸에 들어갈 요약입니다. 필요하면 직접 수정하세요.",bg="white",fg=ui.TEXT,font=("Malgun Gothic",9,"bold")).pack(anchor="w")
        tk.Label(body,text=f"※ {MAX_LEN}자 이내 · 문장을 중간에서 자르지 않습니다.",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8)).pack(anchor="w",pady=(3,10))
        tk.Label(body,text="8D 현상 원문",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(anchor="w")
        original=tk.Text(body,height=6,wrap="word",bg="#F7F9FB",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BORDER,font=("Malgun Gothic",9),padx=9,pady=7); original.pack(fill="x",pady=(4,10)); original.insert("1.0",full_text or "(원문 없음)"); original.configure(state="disabled")
        line=tk.Frame(body,bg="white"); line.pack(fill="x"); tk.Label(line,text="이슈 DB 자동 요약본 (수정 가능)",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(side="left"); self.count=tk.Label(line,text="",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")); self.count.pack(side="right")
        self.edit=tk.Text(body,height=5,wrap="word",bg="white",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BLUE,font=("Malgun Gothic",9),padx=9,pady=7); self.edit.pack(fill="x",pady=(4,4)); self.edit.insert("1.0",summary); self.edit.bind("<KeyRelease>",self._count); self._count()
        foot=tk.Frame(self,bg="#F6F8FA",height=66); foot.pack(fill="x"); foot.pack_propagate(False); b=tk.Frame(foot,bg="#F6F8FA"); b.pack(side="right",padx=20,pady=13)
        for text,val,primary in (("취소",None,False),("원문 전체 사용","full",False),("요약본 적용","summary",True)): self._button(b,text,lambda v=val:self.choose(v),primary).pack(side="left",padx=4)
        self.protocol("WM_DELETE_WINDOW",lambda:self.choose(None)); ui.center_window(self,parent,820,610); self.deiconify(); self.grab_set(); self.edit.focus_set(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary):return tk.Button(p,text=text,command=cmd,width=13,bd=0,font=("Malgun Gothic",9,"bold"),bg=ui.BLUE if primary else "#E5EBF0",fg="white" if primary else ui.TEXT,pady=7,cursor="hand2")
    def _count(self,_event=None):
        n=len(self.edit.get("1.0","end-1c").strip()); self.count.configure(text=f"{n} / {MAX_LEN}자",fg="#B42318" if n>MAX_LEN else ui.MUTED)
    def choose(self,value):
        if value=="summary":
            text=self.edit.get("1.0","end-1c").strip()
            if not text:return ui.warning(self,"현상 요약 확인","요약본이 비어 있습니다.")
            if len(text)>MAX_LEN:return ui.warning(self,"현상 요약 확인",f"요약본은 최대 {MAX_LEN}자입니다. 현재 {len(text)}자입니다.")
            self.summary=text
        self.result=value; self.destroy()

class WeeklyEditDialog(tk.Toplevel):
    def __init__(self,parent,text,full_text=""):
        super().__init__(parent); self.withdraw(); self.result=None; self.summary=None; self.title("주간회의 현상 수정"); self.configure(bg="white"); self.transient(parent)
        tk.Label(self,text="주간회의 요약 페이지 현상 수정",bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold"),anchor="w",padx=22).pack(fill="x",ipady=17)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=16)
        tk.Label(body,text="8D 현상 원문 (참고용)",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(anchor="w")
        original=tk.Text(body,height=7,wrap="word",bg="#F7F9FB",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BORDER,font=("Malgun Gothic",9),padx=9,pady=7); original.pack(fill="x",pady=(4,12)); original.insert("1.0",full_text or "(원문 없음)"); original.configure(state="disabled")
        tk.Label(body,text="주간회의 요약 페이지 문구 (수정 가능)",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(anchor="w")
        self.edit=tk.Text(body,height=7,wrap="word",font=("Malgun Gothic",9),padx=9,pady=7); self.edit.pack(fill="both",expand=True,pady=(4,0)); self.edit.insert("1.0",text)
        b=tk.Frame(self,bg="#F6F8FA"); b.pack(fill="x"); tk.Button(b,text="취소",command=self.destroy,width=12).pack(side="right",padx=4,pady=13); tk.Button(b,text="수정 적용",command=self.apply,width=12,bg=ui.BLUE,fg="white").pack(side="right",padx=4,pady=13)
        ui.center_window(self,parent,780,620); self.deiconify(); self.grab_set(); self.edit.focus_set(); parent.wait_window(self)
    def apply(self):
        text=self.edit.get("1.0","end-1c").strip()
        if not text:return ui.warning(self,"주간회의 현상 수정","내용이 비어 있습니다.")
        self.summary=text; self.result="summary"; self.destroy()

class WeeklyProblemDialog(tk.Toplevel):
    def __init__(self,parent,summary,full_text=""):
        super().__init__(parent); self.withdraw(); self.result=None; self.summary=summary; self.full_text=full_text; self.title("주간회의 요약 페이지 현상"); self.configure(bg="white"); self.transient(parent)
        tk.Label(self,text="주간회의 요약 페이지 현상",bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold"),anchor="w",padx=22).pack(fill="x",ipady=17)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=18); tk.Label(body,text="전체 내용/이슈 DB 요약본을 선택하거나, 주간회의용으로 별도 수정할 수 있습니다.\n상세 페이지 2D는 8D 원문 전체를 유지합니다.",bg="white",fg=ui.TEXT,justify="left").pack(anchor="w")
        box=tk.Text(body,height=8,wrap="word",bg="#F7F9FB",font=("Malgun Gothic",9)); box.pack(fill="both",expand=True,pady=(10,0)); box.insert("1.0",summary or "(요약 내용 없음)"); box.configure(state="disabled")
        b=tk.Frame(self,bg="#F6F8FA"); b.pack(fill="x")
        for text,cmd in (("취소",lambda:self.choose(None)),("전체 내용 기재",lambda:self.choose("full")),("수정",self.modify),("요약본 기재",lambda:self.choose("summary"))): tk.Button(b,text=text,command=cmd,width=14,pady=7).pack(side="right",padx=4,pady=13)
        ui.center_window(self,parent,760,470); self.deiconify(); self.grab_set(); parent.wait_window(self)
    def modify(self):
        ed=WeeklyEditDialog(self,self.summary or self.full_text,self.full_text)
        if ed.result=="summary":self.summary=ed.summary; self.choose("summary")
    def choose(self,v):self.result=v; self.destroy()

def _ask_weekly(parent):
    global _active_weekly_choice,_active_weekly_summary
    wd=WeeklyProblemDialog(parent,_active_summary,getattr(parent,"_problem_full_text",""))
    if wd.result is None:return False
    _active_weekly_choice=wd.result; _active_weekly_summary=wd.summary if wd.result=="summary" else ""; return True

def _problem_dialog_factory(parent,_legacy_summary):
    global _active_summary,_active_db_choice
    dlg=EditableProblemDialog(parent,getattr(parent,"_problem_summary_suggested",""),getattr(parent,"_problem_full_text",""))
    if dlg.result is None:return dlg
    _active_db_choice=dlg.result
    if dlg.result=="summary":_active_summary=dlg.summary
    if getattr(parent,"_problem_weekly_enabled",False) and not _ask_weekly(parent):dlg.result=None; _active_db_choice=None
    return dlg

def _update_excel_with_confirmed_summary(src,out,d,g,new=False):
    gg=dict(g)
    if _active_db_choice=="summary" and _active_summary:gg["_db_problem_selected"]=_db_clean(_active_summary)
    elif _active_db_choice=="full":gg["_db_problem_selected"]=_db_clean(d.get("problem"))
    return _original_update_excel(src,out,d,gg,new=new)

def _weekly_with_problem_choice(src,out,d,g,mode):
    if _active_weekly_choice!="summary":return _original_weekly(src,out,d,g,mode)
    chosen=_active_weekly_summary or _active_summary
    if not chosen:return _original_weekly(src,out,d,g,mode)
    def summary_only(prs,dd,gg,mm):
        dsum=dict(dd); dsum["problem"]=chosen; return _original_summary_update(prs,dsum,gg,mm)
    legacy.s14._update_summary_by_task=summary_only
    try:return _original_weekly(src,out,d,g,mode)
    finally:legacy.s14._update_summary_by_task=_original_summary_update

def _run_with_problem_summary(self):
    global _active_summary,_active_db_choice,_active_weekly_choice,_active_weekly_summary
    _active_summary=""; _active_db_choice=None; _active_weekly_choice=None; _active_weekly_summary=""
    g=self.gui(); ppt8d=g.get("ppt8d","")
    if ppt8d:
        try:
            d=ent.base.extract(ppt8d); self._problem_full_text=N(d.get("problem")); self._problem_summary_suggested=_compact_problem(d); _active_summary=self._problem_summary_suggested
        except Exception:self._problem_full_text=""; self._problem_summary_suggested=""
    weekly=bool(g.get("pptweekly","").strip()); excel=bool(g.get("xlsx","").strip()); self._problem_weekly_enabled=weekly
    if weekly and not excel and ppt8d:
        if not _ask_weekly(self):return
    return _original_run(self)

ent.EnterpriseProblemDialog=_problem_dialog_factory
ent.base.update_excel=_update_excel_with_confirmed_summary
ent.base.weekly=_weekly_with_problem_choice
v3.EnterpriseAppV3.run=_run_with_problem_summary
