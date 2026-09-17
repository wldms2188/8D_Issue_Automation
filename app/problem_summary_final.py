"""Final problem-summary workflow.

- Builds a compact Issue DB suggestion (target <=120 chars; aim around 100-120 when source supports it).
- Preserves test/evaluation names and core symptom wording before removing boilerplate/duplicates.
- Lets the user edit the suggestion before applying it.
- Issue DB and weekly summary page can independently use full text or the confirmed summary.
- Weekly detail-page 2D remains the original full 8D text.
"""
import re
import tkinter as tk

import main_enterprise_v3 as v3
import main_enterprise as ent
import ui_enterprise as ui

legacy = ent.legacy
N = ent.N
MAX_LEN = 120

_original_problem_dialog = ent.EnterpriseProblemDialog
_original_run = v3.EnterpriseAppV3.run
_original_weekly = ent.base.weekly


def _norm(s):
    return re.sub(r"\s+", " ", N(s)).strip()


def _compact_problem(d):
    """Rule-based, non-AI summary. Never invents content not present in 2D."""
    src = N(d.get("problem"))
    if not src:
        return ""

    # Start with lines already judged as actual symptoms by the validated parser.
    try:
        symptom_lines = [N(x) for x in legacy.step3._actual_symptoms(d) if N(x)]
    except Exception:
        symptom_lines = []

    # Test/evaluation names are important context and must be retained when present.
    chunks = [x.strip(" -•·\t") for x in re.split(r"[\n\r]+", src) if x.strip()]
    test_words = ("시험", "평가", "검사", "test", "eol", "dv", "cv", "pd")
    test_lines = [x for x in chunks if any(w in x.casefold() for w in test_words)]

    candidates = test_lines + symptom_lines + chunks
    cleaned = []
    seen = set()
    drop_line_heads = ("발생경위", "확인사항", "시험조건", "시험 조건", "시험방법", "시험 방법", "시험절차", "시험 절차", "시험환경", "시험 환경", "평가조건", "평가 조건")
    for raw in candidates:
        s = _norm(raw)
        if not s:
            continue
        # Do not discard a line containing a test name merely because it also contains conditions.
        has_test = any(w in s.casefold() for w in test_words)
        if not has_test and any(s.replace(" ", "").casefold().startswith(x.replace(" ", "").casefold()) for x in drop_line_heads):
            continue
        # Remove repetitive report-style endings, not technical nouns/values.
        s = re.sub(r"(?:하는\s*)?현상이\s*(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$", "", s)
        s = re.sub(r"(?:재)?확인(?:되었|됐)(?:습니다|음|다)?\.?$", "", s)
        s = re.sub(r"확인\s*결과\s*", "", s)
        s = re.sub(r"\s+", " ", s).strip(" ,.;·-")
        key = re.sub(r"[\s,.;:·/\\_-]+", "", s).casefold()
        if not s or key in seen:
            continue
        # Skip a near-duplicate if its normalized text is contained in an already kept line.
        if any(key and (key in old or old in key) for old in seen):
            continue
        seen.add(key)
        cleaned.append(s)

    text = " / ".join(cleaned) if cleaned else _norm(src)
    if len(text) <= MAX_LEN:
        return text

    # Keep complete chunks where possible; the final hard cap guarantees DB length.
    out = []
    total = 0
    for part in cleaned:
        extra = len(part) + (3 if out else 0)
        if total + extra <= MAX_LEN:
            out.append(part); total += extra
        else:
            remain = MAX_LEN - total - (3 if out else 0)
            if remain >= 8:
                out.append(part[:remain].rstrip(" ,.;·-/"))
            break
    result = " / ".join(out).strip()
    return (result or text)[:MAX_LEN].rstrip(" ,.;·-/")


class EditableProblemDialog(tk.Toplevel):
    def __init__(self, parent, summary, full_text=""):
        super().__init__(parent)
        self.withdraw(); self.result = None; self.summary = None
        self.title("Issue DB 현상 입력"); self.configure(bg="white"); self.transient(parent)
        head = tk.Frame(self, bg=ui.NAVY, height=58); head.pack(fill="x"); head.pack_propagate(False)
        tk.Label(head, text="현상 요약 확인", bg=ui.NAVY, fg="white", font=("Malgun Gothic",12,"bold")).pack(side="left", padx=22)
        body = tk.Frame(self, bg="white"); body.pack(fill="both", expand=True, padx=24, pady=16)
        tk.Label(body, text="시험명·대상/위치·핵심 불량현상을 우선 유지해 요약했습니다. 필요하면 아래 축약본을 직접 수정하세요.", bg="white", fg=ui.TEXT, font=("Malgun Gothic",9,"bold"), wraplength=780, justify="left").pack(anchor="w")
        tk.Label(body, text="※ 권장 100~120자 · 최대 120자", bg="white", fg=ui.MUTED, font=("Malgun Gothic",8)).pack(anchor="w", pady=(3,10))
        tk.Label(body, text="8D 현상 원문", bg="white", fg=ui.MUTED, font=("Malgun Gothic",8,"bold")).pack(anchor="w")
        original = tk.Text(body, height=6, wrap="word", bg="#F7F9FB", fg=ui.TEXT, relief="flat", highlightthickness=1, highlightbackground=ui.BORDER, font=("Malgun Gothic",9), padx=9, pady=7)
        original.pack(fill="x", pady=(4,10)); original.insert("1.0", full_text or "(원문 없음)"); original.configure(state="disabled")
        line = tk.Frame(body, bg="white"); line.pack(fill="x")
        tk.Label(line, text="자동 축약본 (수정 가능)", bg="white", fg=ui.MUTED, font=("Malgun Gothic",8,"bold")).pack(side="left")
        self.count = tk.Label(line, text="", bg="white", fg=ui.MUTED, font=("Malgun Gothic",8,"bold")); self.count.pack(side="right")
        self.edit = tk.Text(body, height=5, wrap="word", bg="white", fg=ui.TEXT, relief="flat", highlightthickness=1, highlightbackground=ui.BLUE, font=("Malgun Gothic",9), padx=9, pady=7)
        self.edit.pack(fill="x", pady=(4,4)); self.edit.insert("1.0", summary)
        self.edit.bind("<KeyRelease>", self._count); self._count()
        foot = tk.Frame(self, bg="#F6F8FA", height=66); foot.pack(fill="x"); foot.pack_propagate(False)
        b = tk.Frame(foot, bg="#F6F8FA"); b.pack(side="right", padx=20, pady=13)
        self._button(b,"취소",lambda:self.choose(None),False).pack(side="left",padx=4)
        self._button(b,"원문 전체 사용",lambda:self.choose("full"),False).pack(side="left",padx=4)
        self._button(b,"요약본 적용",lambda:self.choose("summary"),True).pack(side="left",padx=4)
        self.protocol("WM_DELETE_WINDOW",lambda:self.choose(None)); ui.center_window(self,parent,820,610); self.deiconify(); self.grab_set(); self.edit.focus_set(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary):
        return tk.Button(p,text=text,command=cmd,width=13,bd=0,font=("Malgun Gothic",9,"bold"),bg=ui.BLUE if primary else "#E5EBF0",fg="white" if primary else ui.TEXT,pady=7,cursor="hand2")
    def _count(self,_event=None):
        n=len(self.edit.get("1.0","end-1c").strip()); self.count.configure(text=f"{n} / {MAX_LEN}자", fg="#B42318" if n>MAX_LEN else ui.MUTED)
    def choose(self,value):
        if value == "summary":
            text=self.edit.get("1.0","end-1c").strip()
            if not text:
                return ui.warning(self,"현상 요약 확인","요약본이 비어 있습니다.")
            if len(text)>MAX_LEN:
                return ui.warning(self,"현상 요약 확인",f"요약본은 최대 {MAX_LEN}자까지 입력할 수 있습니다.\n현재 {len(text)}자입니다.")
            self.summary=text
        self.result=value; self.destroy()


class WeeklyProblemDialog(tk.Toplevel):
    def __init__(self,parent,summary):
        super().__init__(parent); self.withdraw(); self.result=None; self.title("주간회의 요약 페이지 현상"); self.configure(bg="white"); self.transient(parent); self.resizable(False,False)
        head=tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill="x"); head.pack_propagate(False)
        tk.Label(head,text="주간회의 요약 페이지 현상",bg=ui.NAVY,fg="white",font=("Malgun Gothic",12,"bold")).pack(side="left",padx=22)
        body=tk.Frame(self,bg="white"); body.pack(fill="both",expand=True,padx=24,pady=18)
        tk.Label(body,text="주간회의 요약 페이지에도 확정한 현상 요약본을 기재하시겠습니까?",bg="white",fg=ui.TEXT,font=("Malgun Gothic",10,"bold"),wraplength=690,justify="left").pack(anchor="w")
        tk.Label(body,text="상세 페이지의 2D 현상은 선택과 관계없이 8D 원문 전체를 유지합니다.",bg="white",fg=ui.MUTED,font=("Malgun Gothic",8)).pack(anchor="w",pady=(4,12))
        box=tk.Text(body,height=8,wrap="word",bg="#F7F9FB",fg=ui.TEXT,relief="flat",highlightthickness=1,highlightbackground=ui.BORDER,font=("Malgun Gothic",9),padx=9,pady=7); box.pack(fill="both",expand=True); box.insert("1.0",summary or "(요약 내용 없음)"); box.configure(state="disabled")
        foot=tk.Frame(self,bg="#F6F8FA",height=66); foot.pack(fill="x"); foot.pack_propagate(False); b=tk.Frame(foot,bg="#F6F8FA"); b.pack(side="right",padx=20,pady=13)
        self._button(b,"취소",lambda:self.choose(None),False).pack(side="left",padx=4); self._button(b,"전체 내용 기재",lambda:self.choose("full"),False).pack(side="left",padx=4); self._button(b,"요약본 기재",lambda:self.choose("summary"),True).pack(side="left",padx=4)
        self.protocol("WM_DELETE_WINDOW",lambda:self.choose(None)); ui.center_window(self,parent,740,440); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def _button(self,p,text,cmd,primary): return tk.Button(p,text=text,command=cmd,width=14,bd=0,font=("Malgun Gothic",9,"bold"),bg=ui.BLUE if primary else "#E5EBF0",fg="white" if primary else ui.TEXT,pady=7,cursor="hand2")
    def choose(self,v): self.result=v; self.destroy()


def _problem_dialog_factory(parent, summary):
    # main_enterprise.run calls this after extracting d.  The full text and better
    # summary are supplied by the run wrapper below via temporary attributes.
    full=getattr(parent,"_problem_full_text","")
    suggested=getattr(parent,"_problem_summary_suggested",summary)
    dlg=EditableProblemDialog(parent,suggested,full)
    # Preserve the old interface expected by main_enterprise.run.
    if dlg.result == "summary" and dlg.summary is not None:
        parent._problem_summary_confirmed=dlg.summary
    return dlg


def _weekly_with_problem_choice(src,out,d,g,mode):
    # Only the weekly SUMMARY should use the chosen short wording. Detail 2D must
    # remain full, so temporarily swap problem for the summary-update call and then
    # restore it before the detail-page update runs.
    choice=g.get("_weekly_problem_selected")
    selected=g.get("_weekly_problem_summary")
    if choice != "summary" or not selected:
        return _original_weekly(src,out,d,g,mode)

    original=N(d.get("problem"))
    d["problem"]=selected
    try:
        # weekly implementation updates summary before detail. Mark the original so
        # lower-level detail code can restore it through the wrapper value in g.
        g["_weekly_detail_problem_original"]=original
        return _original_weekly(src,out,d,g,mode)
    finally:
        d["problem"]=original
        g.pop("_weekly_detail_problem_original",None)


# Patch the dialog. A run wrapper prepares the suggestion and asks the independent
# weekly choice. The existing validated run continues to own matching/status/output.
ent.EnterpriseProblemDialog = _problem_dialog_factory


def _run_with_problem_summary(self):
    g=self.gui()
    ppt8d=g.get("ppt8d","")
    if ppt8d:
        try:
            d=ent.base.extract(ppt8d)
            self._problem_full_text=N(d.get("problem"))
            self._problem_summary_suggested=_compact_problem(d)
            self._problem_summary_confirmed=self._problem_summary_suggested
        except Exception:
            self._problem_full_text=""; self._problem_summary_suggested=""; self._problem_summary_confirmed=""

    # The legacy run will show the editable Issue DB dialog only when Excel is used.
    # For weekly-only runs, ask the same editable confirmation first so a summary is available.
    weekly=bool(g.get("pptweekly","").strip()); excel=bool(g.get("xlsx","").strip())
    if weekly and not excel and ppt8d:
        dlg=EditableProblemDialog(self,self._problem_summary_suggested,self._problem_full_text)
        if dlg.result is None:
            return
        self._problem_summary_confirmed=dlg.summary if dlg.result=="summary" else self._problem_summary_suggested

    # Ask weekly choice before processing. If Excel is also selected, the confirmed
    # edit happens inside legacy run, so defer this question until that dialog closes
    # by intercepting its constructor and using a pending marker.
    self._weekly_problem_pending = weekly
    return _original_run(self)


# Extend the factory so weekly selection happens immediately after the editable DB dialog.
def _problem_dialog_factory(parent, summary):
    full=getattr(parent,"_problem_full_text","")
    suggested=getattr(parent,"_problem_summary_suggested",summary)
    dlg=EditableProblemDialog(parent,suggested,full)
    if dlg.result == "summary" and dlg.summary is not None:
        parent._problem_summary_confirmed=dlg.summary
    confirmed=getattr(parent,"_problem_summary_confirmed",suggested)
    if getattr(parent,"_weekly_problem_pending",False) and dlg.result is not None:
        wd=WeeklyProblemDialog(parent,confirmed)
        if wd.result is None:
            dlg.result=None
        else:
            parent._weekly_problem_choice=wd.result
    return dlg

ent.EnterpriseProblemDialog = _problem_dialog_factory

# Inject weekly choice into gui() so the existing run passes it to base.weekly.
_original_gui=v3.EnterpriseAppV3.gui

def _gui_with_problem_choice(self):
    g=_original_gui(self)
    choice=getattr(self,"_weekly_problem_choice",None)
    if choice:
        g["_weekly_problem_selected"]=choice
        g["_weekly_problem_summary"]=getattr(self,"_problem_summary_confirmed","")
    return g

v3.EnterpriseAppV3.gui=_gui_with_problem_choice
v3.EnterpriseAppV3.run=_run_with_problem_summary
ent.base.weekly=_weekly_with_problem_choice
