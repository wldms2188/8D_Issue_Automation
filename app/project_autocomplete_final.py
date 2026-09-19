"""Customer-project autocomplete linked to 담당팀, plus canonical project selection."""
import re
import tkinter as tk
from pathlib import Path
from tkinter import ttk
import main_enterprise_v3 as v3
import ui_enterprise as ui

PROJECTS={
"파우치형Pack개발품질1팀":("Model Care 25","MBAG_EB565M","Ford_V710","Renault_EV2020 87 Mid Ni 400V","Renault_EV2020 87 Mid Ni 800V","Renault_EV2020 CTP 400V","Renault_EV2020 CTP 800V"),
"파우치형Pack개발품질2팀":("STLA_VDA590_2P8S","STLA_VDA590_1P16S","STLA_VDA355","Honda_NA BEV_3P8S","Honda_ZS","Toyota_E-TNGA","Toyota_Li9.B","Ford_FHEV"),
"원통형Pack개발품질팀":("MBAG_VAN EV VarC","Chery_T19C","Audi_SSP41","BDI_BOLT","BDI_ODIE","Honda_MLMA(EV-Fun)","Wheel_TWINS","MBAG_EB-L(EU)","MBAG_EB-L(US)","LGE_TCUA 4pin","PMI_BETA","PMI_Ariane 1.6+","KS 표준팩"),
"ESS System개발품질팀":("JF2 0.25CP CIMC","JF2 0.25CP NEER (CWP)","JF2 0.25CP AC Type","JF2 0.5CP AC Type (Vertech)","JF2 0.25CP 중앙계약시장","JF2S Pack","JF2S 0.5CP DC Type","JF2S Delta (Set_Biz)","JF2S EG4 (CMA_Biz)","JF2 0.25CP Unibody","JF1 0.33CP","JF1 0.33CP 일본, 대만","JF1 0.5CP 유럽","JF1 1.0CP 터키향 (Pack)","JF1 1.0CP 터키향 (BPU)","Fronius JF1 Small C&I Pack","JF1 Fronius (주택용Pack)","JF1 Fronius (주택용BPU)","JF1R Fronius (Pack)","JF1R Fronius (BPU)","JF3 0.5CP AC Type","JF3 0.25CP AC Type","JF3 0.5CP DV Type","JP6 (UPS)","JF2 0.25CP AC Type Ver2","JF2 0.25CP 중앙계약시장 2차","JF2SC (염가형)","S-frame"),
"회로개발품질팀":(),
"양산이슈":(),
}
_original_entry=v3.EnterpriseAppV3._enterprise_entry
_original_run=v3.EnterpriseAppV3.run

def _norm(s): return re.sub(r'[^0-9A-Za-z가-힣]+','',str(s or '')).casefold()
def _no_paren(s): return _norm(re.sub(r'\([^)]*\)','',str(s or '')))
def _all_projects(): return tuple(x for xs in PROJECTS.values() for x in xs)

def split_customer_task(value):
    """Split A_B at the first underscore outside parentheses; underscores inside (...) stay in the project name."""
    s=str(value or '').strip()
    depth=0
    for i,ch in enumerate(s):
        if ch=='(':
            depth+=1
        elif ch==')' and depth:
            depth-=1
        elif ch=='_' and depth==0:
            return s[:i].strip(),s[i+1:].strip()
    return '',s

def project_part(value):
    return split_customer_task(value)[1]

def project_key(value,drop_parentheses=False):
    task=project_part(value)
    return _no_paren(task) if drop_parentheses else _norm(task)

_FILENAME_STOP_WORDS={
    '8d','report','issue','final','rev','revision','update','updated','ppt','pptx',
    'pack','system','quality','claim','problem'
}

def _filename_words(value):
    words=re.findall(r'[0-9A-Za-z가-힣]+',str(value or '').casefold())
    return [w for w in words if len(w)>=2 and w not in _FILENAME_STOP_WORDS]

def filename_project_candidates(path,team=""):
    """Find catalog projects whose name/meaningful words overlap the 8D filename.

    Exact project-name/project-part matches win. If no exact match exists, return
    meaningful word-overlap candidates in best-match order.
    """
    try:
        stem=Path(path).stem
    except Exception:
        stem=str(path or '')
    if not stem:
        return []

    pool=PROJECTS.get(team,()) or _all_projects()
    stem_norm=_norm(stem)
    stem_words=set(_filename_words(stem))
    exact=[]; overlap=[]

    for project in pool:
        full_norm=_norm(project)
        part=project_part(project)
        part_norm=_norm(part)

        # Full catalog phrase or project-only phrase appears in the filename.
        if (full_norm and full_norm in stem_norm) or (part_norm and len(part_norm)>=4 and part_norm in stem_norm):
            exact.append(project)
            continue

        terms=_filename_words(part)
        if not terms:
            continue
        matched=[]
        for term in terms:
            # Token match first; allow compact filename matches for long/model-like tokens.
            if term in stem_words or (len(term)>=4 and _norm(term) in stem_norm):
                matched.append(term)
        if not matched:
            continue

        # A single very short/common token is too weak (e.g. JF1); otherwise keep
        # the candidate so shared model words can surface more than one project.
        strongest=max((len(x) for x in matched),default=0)
        if len(matched)==1 and strongest<4:
            continue
        ratio=len(matched)/max(1,len(terms))
        weight=sum(len(x) for x in matched)
        overlap.append((ratio,weight,len(matched),project))

    if exact:
        # Preserve catalog order and de-duplicate.
        seen=set(); out=[]
        for x in exact:
            if x not in seen:
                seen.add(x); out.append(x)
        return out

    overlap.sort(key=lambda z:(-z[0],-z[1],-z[2],z[3].casefold()))
    return [x[3] for x in overlap[:8]]

def canonical_candidates(value,team=""):
    q=project_key(value); qp=project_key(value,True)
    pool=PROJECTS.get(team,()) or _all_projects()
    exact_paren=[x for x in pool if qp and project_key(x,True)==qp and project_key(x)!=q]
    if exact_paren:return exact_paren
    if not q:return list(pool)
    return [x for x in pool if q in project_key(x) or project_key(x) in q]

def _project_entry(self,parent,label,key,hint):
    if key!="task_name": return _original_entry(self,parent,label,key,hint)
    row=tk.Frame(parent,bg="white"); row.pack(fill="x",pady=4)
    tk.Label(row,text=label,bg="white",fg=ui.TEXT,font=("Malgun Gothic",9),width=17,anchor="w").pack(side="left")
    wrap=tk.Frame(row,bg="white")
    wrap.pack(side="left",fill="x",expand=True)
    entry=ttk.Entry(wrap,textvariable=self.vars[key]); entry.pack(fill="x")
    popup=None; listbox=None
    def close(*_):
        nonlocal popup,listbox
        if popup:
            try: popup.destroy()
            except Exception: pass
        popup=listbox=None

    def _inside_popup(widget):
        if popup is None or widget is None:
            return False
        try:
            return widget.winfo_toplevel() is popup
        except Exception:
            return False

    def choose(*_):
        if not listbox:return
        sel=listbox.curselection()
        if sel:
            self.vars[key].set(listbox.get(sel[0]))
            close()
            try: entry.icursor('end')
            except Exception: pass

    def enter_or_close(_event=None):
        # Enter accepts an explicitly highlighted candidate; otherwise it simply
        # keeps the typed text and dismisses the candidate popup.
        if listbox:
            try:
                if listbox.curselection():
                    choose()
                    return "break"
            except Exception:
                pass
        close()
        return "break"

    def close_if_focus_left():
        # FocusOut can occur while clicking a candidate. Delay the check so a
        # popup/listbox click is not mistaken for clicking somewhere else.
        try:
            focus=self.focus_get()
        except Exception:
            focus=None
        if focus is entry or _inside_popup(focus):
            return
        close()

    def on_global_click(event):
        # Close even when the user clicks a non-focusable area (label/frame),
        # while preserving clicks inside the entry or its candidate popup.
        widget=getattr(event,'widget',None)
        if widget is entry or _inside_popup(widget):
            return
        close()

    def show(*_):
        nonlocal popup,listbox
        close(); vals=canonical_candidates(self.vars[key].get(),self.vars["team"].get().strip())
        if not vals:return
        self.update_idletasks(); popup=tk.Toplevel(self); popup.wm_overrideredirect(True); popup.attributes("-topmost",True)
        x,y=entry.winfo_rootx(),entry.winfo_rooty()+entry.winfo_height(); w=max(entry.winfo_width(),420); h=min(8,len(vals))*26+4
        popup.geometry(f"{w}x{h}+{x}+{y}"); listbox=tk.Listbox(popup,font=("Malgun Gothic",9),exportselection=False)
        listbox.pack(fill="both",expand=True)
        for x0 in vals:listbox.insert("end",x0)
        listbox.bind("<ButtonRelease-1>",choose); listbox.bind("<Return>",choose)
    # Open suggestions on an intentional click or typing.  FocusIn used to reopen
    # the popup immediately after a user selected an item.
    entry.bind("<Button-1>",lambda _e:self.after_idle(show),add="+")
    entry.bind("<KeyRelease>",lambda e: None if e.keysym in ("Up","Down","Return","Escape","Tab") else show(),add="+")
    entry.bind("<Return>",enter_or_close,add="+")
    entry.bind("<Escape>",close,add="+")
    entry.bind("<Tab>",lambda _e:close(),add="+")
    entry.bind("<FocusOut>",lambda _e:self.after_idle(close_if_focus_left),add="+")
    self.bind_all("<Button-1>",on_global_click,add="+")
    tk.Label(
        wrap,
        text="담당팀 연계 과제 후보 · 입력 시 유사 과제 자동완성",
        bg="white",fg="#98A3AD",font=("Malgun Gothic",7),anchor="e"
    ).pack(fill="x",anchor="e")

def _confirm_project(self):
    v=self.vars.get("task_name").get().strip() if self.vars.get("task_name") else ""
    if not v:
        ok=ui.ask_yes_no(self,"고객사/과제명 확인","고객사/과제명을 입력하지 않았습니다.\n\n자동 분류로 진행하시겠습니까?")
        if ok:
            # Do not inject a manual task_name. The original V1 extraction/matching path
            # will classify customer/project from the 8D source as before.
            return True
        return False
    team=self.vars["team"].get().strip()
    pool=PROJECTS.get(team,()) or _all_projects()
    if v in pool:return True
    same=[x for x in pool if project_key(x,True)==project_key(v,True) and project_key(x)!=project_key(v)]
    if same:
        # Simple chooser dialog for parenthesized canonical variants.
        win=tk.Toplevel(self); win.title("유사한 과제가 있습니다."); win.transient(self); win.grab_set(); result={"v":None}
        tk.Label(win,text="유사한 과제가 있습니다.\n아래 과제 중 하나를 선택해 주세요.",font=("Malgun Gothic",10,"bold"),justify="left").pack(anchor="w",padx=20,pady=(18,8))
        lb=tk.Listbox(win,font=("Malgun Gothic",9),height=min(8,len(same)),width=58,exportselection=False); lb.pack(fill="x",padx=20,pady=6)
        for x in same:lb.insert("end",x)
        def ok():
            s=lb.curselection()
            if s:result["v"]=lb.get(s[0]); win.destroy()
        def cancel():win.destroy()
        b=tk.Frame(win); b.pack(pady=14); ttk.Button(b,text="취소",command=cancel).pack(side="left",padx=5); ttk.Button(b,text="선택",command=ok).pack(side="left",padx=5)
        ui.center_window(win,self,540,300); self.wait_window(win)
        if result["v"] is not None:self.vars["task_name"].set(result["v"]); return True
        return ui.ask_yes_no(self,"과제명 확인",f"입력한 과제명 '{v}'으로 입력하겠습니까?")
    if v not in pool:
        return ui.ask_yes_no(self,"과제명 확인",f"입력한 과제명 '{v}'은 등록된 과제 목록에 없습니다.\n\n이 과제명으로 입력하겠습니까?")
    return True

def _run(self):
    if not _confirm_project(self):return
    return _original_run(self)

v3.EnterpriseAppV3._enterprise_entry=_project_entry
v3.EnterpriseAppV3.run=_run
