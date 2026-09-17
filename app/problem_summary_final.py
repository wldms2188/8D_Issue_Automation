"""Editable problem-summary workflow for the final enterprise launcher."""
import re
import tkinter as tk

import main_enterprise_v3 as v3
import main_enterprise as ent
import ui_enterprise as ui

legacy = ent.legacy
N = ent.N
MAX_LEN = 120

_original_run = v3.EnterpriseAppV3.run
_original_weekly = ent.base.weekly
_original_update_excel = ent.base.update_excel
_original_summary_update = legacy.s14._update_summary_by_task

# One GUI run is synchronous, so these values safely carry the user's choices
# from modal dialogs into the existing Excel/weekly processing functions.
_active_summary = ""
_active_db_choice = None
_active_weekly_choice = None


def _norm(s):
    return re.sub(r"\s+", " ", N(s)).strip()


def _compact_problem(d):
    """Create a <=120-char suggestion using only text present in 2D."""
    src = N(d.get("problem"))
    if not src:
        return ""
    try:
        symptom_lines = [N(x) for x in legacy.step3._actual_symptoms(d) if N(x)]
    except Exception:
        symptom_lines = []

    chunks = [x.strip(" -•·\t") for x in re.split(r"[\n\r]+", src) if x.strip()]
    # Preserve lines carrying test/evaluation context before generic symptom lines.
    test_words = ("시험", "평가", "검사", "test", "eol", "dv", "cv", "pd")
    test_lines = [x for x in chunks if any(w in x.casefold() for w in test_words)]
    candidates = test_lines + symptom_lines + chunks

    cleaned, seen = [], []
    drop_heads = ("발생경위", "확인사항", "시험조건", "시험 조건", "시험방법", "시험 방법", "시험절차", "시험 절차", "시험환경", "시험 환경", "평가조건", "평가 조건")
    for raw in candidates:
        s = _norm(raw)
        if not s:
            continue
        has_test = any(w in s.casefold() for w in test_words)
        if not has_test and any(s.replace(" ", "").casefold().startswith(x.replace(" ", "").casefold()) for x in drop_heads):
            continue
        s = re.sub(r"(?:하는\s*)?현상이\s*(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$", "", s)
        s = re.sub(r"(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$", "", s)
        s = re.sub(r"확인\s*결과\s*", "", s)
        s = re.sub(r"\s+", " ", s).strip(" ,.;·-")
        key = re.sub(r"[\s,.;:·/\\_-]+", "", s).casefold()
        if not s or key in seen:
            continue
        if any(key and (key in old or old in key) for old in seen):
            continue
        seen.append(key); cleaned.append(s)

    text = " / ".join(cleaned) if cleaned else _norm(src)
    if len(text) <= MAX_LEN:
        return text

    out, total = [], 0
    for part in cleaned:
        sep = 3 if out else 0
        if total + sep + len(part) <= MAX_LEN:
            out.append(part); total += sep + len(part)
            continue
        remain = MAX_LEN - total - sep
        if remain >= 8:
            out.append(part[:remain].rstrip(" ,.;·-/"))
        break
    return (" / ".join(out) or text)[:MAX_LEN].rstrip(" ,.;·-/")


class EditableProblemDialog(tk.Toplevel):
    def __init__(self, parent, summary, full_text=""):
        super().__init__(parent); self.withdraw(); self.result=None; self.summary=None
        self.title("현상 요약 확인"); self.configure(bg="white"); self.transient(parent)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill="x"); head.pack_propagate(False)
        tk.Label(head,text="현상 요약 확인",bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold")).pack(side="left",padx=22)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=16)
        tk.Label(body,text="시험명·대상/위치·핵심 불량현상을 우선 유지해 축약했습니다. 필요하면 직접 수정하세요.",bg="white",fg=ui.TEXT,font=("Malgun Gothic",9,"bold"),wraplength=780,justify="left").pack(anchor="w")
        tk.Label(body,text="※ 120자 이내 · 더 짧아도 괜찮습니다.",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8)).pack(anchor="w",pady=(3,10))
        tk.Label(body,text="8D 현상 원문",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(anchor="w")
        original=tk.Text(body,height=6,wrap="word",bg="#F7F9FB",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BORDER,font=("Malgun Gothic",9),padx=9,pady=7)
        original.pack(fill="x",pady=(4,10)); original.insert("1.0",full_text or "(원문 없음)"); original.configure(state="disabled")
        line=tk.Frame(body,bg="white"); line.pack(fill="x")
        tk.Label(line,text="자동 축약본 (수정 가능)",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")).pack(side="left")
        self.count=tk.Label(line,text="",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8,"bold")); self.count.pack(side="right")
        self.edit=tk.Text(body,height=5,wrap="word",bg="white",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BLUE,font=("Malgun Gothic",9),padx=9,pady=7)
        self.edit.pack(fill="x",pady=(4,4)); self.edit.insert("1.0",summary); self.edit.bind("<KeyRelease>",self._count); self._count()
        foot=tk.Frame(self,bg="#F6F8FA",height=66); foot.pack(fill="x"); foot.pack_propagate(False); b=tk.Frame(foot,bg="#F6F8FA"); b.pack(side="right",padx=20,pady=13)
        self._button(b,"취소",lambda:self.choose(None),False).pack(side="left",padx=4); self._button(b,"원문 전체 사용",lambda:self.choose("full"),False).pack(side="left",padx=4); self._button(b,"요약본 적용",lambda:self.choose("summary"),True).pack(side="left",padx=4)
        self.protocol("WM_DELETE_WINDOW",lambda:self.choose(None)); ui.center_window(self,parent,820,610); self.deiconify(); self.grab_set(); self.edit.focus_set(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary): return tk.Button(p,text=text,command=cmd,width=13,bd=0,font=("Malgun Gothic",9,"bold"),bg=ui.BLUE if primary else "#E5EBF0",fg="white" if primary else ui.TEXT,pady=7,cursor="hand2")
    def _count(self,_event=None):
        n=len(self.edit.get("1.0","end-1c").strip()); self.count.configure(text=f"{n} / {MAX_LEN}자",fg="#B42318" if n>MAX_LEN else ui.MUTED)
    def choose(self,value):
        if value=="summary":
            text=self.edit.get("1.0","end-1c").strip()
            if not text: return ui.warning(self,"현상 요약 확인","요약본이 비어 있습니다.")
            if len(text)>MAX_LEN: return ui.warning(self,"현상 요약 확인",f"요약본은 최대 {MAX_LEN}자까지 입력할 수 있습니다.\n현재 {len(text)}자입니다.")
            self.summary=text
        self.result=value; self.destroy()


class WeeklyProblemDialog(tk.Toplevel):
    def __init__(self,parent,summary):
        super().__init__(parent); self.withdraw(); self.result=None; self.title("주간회의 요약 페이지 현상"); self.configure(bg="white"); self.transient(parent); self.resizable(False,False)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill="x"); head.pack_propagate(False); tk.Label(head,text="주간회의 요약 페이지 현상",bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold")).pack(side="left",padx=22)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=18)
        tk.Label(body,text="주간회의 요약 페이지에도 확정한 현상 요약본을 기재하시겠습니까?",bg="white",fg=ui.TEXT,font=("Malgun Gothic",10,"bold"),wraplength=690,justify="left").pack(anchor="w")
        tk.Label(body,text="상세 페이지의 2D 현상은 항상 8D 원문 전체를 유지합니다.",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8)).pack(anchor="w",pady=(4,12))
        box=tk.Text(body,height=8,wrap="word",bg="#F7F9FB",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BORDER,font=("Malgun Gothic",9),padx=9,pady=7); box.pack(fill="both",expand=True); box.insert("1.0",summary or "(요약 내용 없음)"); box.configure(state="disabled")
        foot=tk.Frame(self,bg="#F6F8FA",height=66); foot.pack(fill="x"); foot.pack_propagate(False); b=tk.Frame(foot,bg="#F6F8FA"); b.pack(side="right",padx=20,pady=13)
        self._button(b,"취소",lambda:self.choose(None),False).pack(side="left",padx=4); self._button(b,"전체 내용 기재",lambda:self.choose("full"),False).pack(side="left",padx=4); self._button(b,"요약본 기재",lambda:self.choose("summary"),True).pack(side="left",padx=4)
        self.protocol("WM_DELETE_WINDOW",lambda:self.choose(None)); ui.center_window(self,parent,740,440); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary): return tk.Button(p,text=text,command=cmd,width=14,bd=0,font=("Malgun Gothic",9,"bold"),bg=ui.BLUE if primary else "#E5EBF0",fg="white" if primary else ui.TEXT,pady=7,cursor="hand2")
    def choose(self,v): self.result=v; self.destroy()


def _problem_dialog_factory(parent, _legacy_summary):
    global _active_summary, _active_db_choice, _active_weekly_choice
    dlg=EditableProblemDialog(parent,getattr(parent,"_problem_summary_suggested",""),getattr(parent,"_problem_full_text",""))
    if dlg.result is None:
        return dlg
    _active_db_choice=dlg.result
    if dlg.result=="summary":
        _active_summary=dlg.summary
    # Excel + weekly: ask the weekly choice immediately after the editable confirmation.
    if getattr(parent,"_problem_weekly_enabled",False):
        wd=WeeklyProblemDialog(parent,_active_summary)
        if wd.result is None:
            dlg.result=None; _active_db_choice=None; return dlg
        _active_weekly_choice=wd.result
    return dlg


def _update_excel_with_confirmed_summary(src,out,d,g,new=False):
    gg=dict(g)
    if _active_db_choice=="summary" and _active_summary:
        gg["_db_problem_selected"]=_active_summary
    elif _active_db_choice=="full":
        gg["_db_problem_selected"]=N(d.get("problem"))
    return _original_update_excel(src,out,d,gg,new=new)


def _weekly_with_problem_choice(src,out,d,g,mode):
    if _active_weekly_choice!="summary" or not _active_summary:
        return _original_weekly(src,out,d,g,mode)

    # Patch ONLY the summary updater during this call. The detail updater receives
    # the untouched original d, therefore detailed 2D remains the complete 8D text.
    def summary_only(prs,dd,gg,mm):
        dsum=dict(dd); dsum["problem"]=_active_summary
        return _original_summary_update(prs,dsum,gg,mm)
    legacy.s14._update_summary_by_task=summary_only
    try:
        return _original_weekly(src,out,d,g,mode)
    finally:
        legacy.s14._update_summary_by_task=_original_summary_update


def _run_with_problem_summary(self):
    global _active_summary, _active_db_choice, _active_weekly_choice
    _active_summary=""; _active_db_choice=None; _active_weekly_choice=None
    g=self.gui(); ppt8d=g.get("ppt8d","")
    if ppt8d:
        try:
            d=ent.base.extract(ppt8d)
            self._problem_full_text=N(d.get("problem"))
            self._problem_summary_suggested=_compact_problem(d)
            _active_summary=self._problem_summary_suggested
        except Exception:
            self._problem_full_text=""; self._problem_summary_suggested=""; _active_summary=""

    weekly=bool(g.get("pptweekly","").strip()); excel=bool(g.get("xlsx","").strip())
    self._problem_weekly_enabled=weekly

    # Weekly-only has no legacy Issue DB dialog, so collect both decisions here.
    if weekly and not excel and ppt8d:
        dlg=EditableProblemDialog(self,self._problem_summary_suggested,self._problem_full_text)
        if dlg.result is None: return
        if dlg.result=="summary": _active_summary=dlg.summary
        wd=WeeklyProblemDialog(self,_active_summary)
        if wd.result is None: return
        _active_weekly_choice=wd.result

    return _original_run(self)


ent.EnterpriseProblemDialog=_problem_dialog_factory
ent.base.update_excel=_update_excel_with_confirmed_summary
ent.base.weekly=_weekly_with_problem_choice
v3.EnterpriseAppV3.run=_run_with_problem_summary
