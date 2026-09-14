import re
import tkinter as tk
from tkinter import ttk, messagebox
import main_recovery_step7 as step7
import main_v310 as v310

base=step7.base
N=v310.N

ORIGINS=('부품','설계','공정','기타','논의 중')

def _cause_text(d):
    parts=[]
    for label,key in [('발생원인','cause_4d'),('유출원인','leak_cause'),('시스템원인','system_cause')]:
        s=N(d.get(key))
        if s: parts.append(f'{label} : {s}')
    return '\n'.join(parts)

def _recommend_origin(d):
    text=_cause_text(d)
    if not text:
        return 'TBD','4D 원인 내용이 아직 없어 TBD로 표시합니다.'
    q=text.lower().replace(' ','')
    scores={k:0 for k in ORIGINS}
    keys={
      '부품':('부품','자재','소재','원재료','supplier','협력사','입고','component','cell','셀불량'),
      '설계':('설계','design','도면','공차','사양','spec','구조','치수','강성','간섭','설계마진'),
      '공정':('공정','작업','조립','체결','토크','용접','검사','검출','설비','조건관리','작업자','생산','가공','도포','압착','공정조건'),
      '기타':('운송','보관','취급','고객','외부','환경','사용조건'),
    }
    for cat,words in keys.items():
        scores[cat]=sum(1 for w in words if w.lower().replace(' ','') in q)
    best=max(('부품','설계','공정','기타'),key=lambda k:scores[k])
    if scores[best]==0:
        return '논의 중','원인 내용은 있으나 부품/설계/공정/기타로 명확히 분류할 근거가 부족합니다.'
    tied=[k for k in ('부품','설계','공정','기타') if scores[k]==scores[best]]
    if len(tied)>1:
        return '논의 중','원인 내용에 여러 이슈기인 범주의 단서가 함께 있어 추가 논의가 필요합니다.'
    evidence=[w for w in keys[best] if w.lower().replace(' ','') in q][:3]
    return best, f"8D 원인 내용에서 {', '.join(evidence)} 관련 표현이 확인되어 {best}을(를) 추천합니다."

class OriginDialog(tk.Toplevel):
    def __init__(self,parent,d):
        super().__init__(parent); self.result=None
        self.title('이슈기인 확인'); self.geometry('760x590'); self.transient(parent); self.grab_set()
        cause=_cause_text(d); rec,reason=_recommend_origin(d)
        ttk.Label(self,text='8D 원인 내용을 기준으로 이슈기인을 추천했습니다.',font=('',10,'bold')).pack(anchor='w',padx=20,pady=(18,8))
        ttk.Label(self,text='8D 원인 내용').pack(anchor='w',padx=20)
        box=tk.Text(self,height=12,wrap='word'); box.pack(fill='both',expand=True,padx=20,pady=(4,10)); box.insert('1.0',cause or '(4D 원인 내용 없음)'); box.configure(state='disabled')
        ttk.Label(self,text=f'추천 이슈기인 : {rec}',font=('',11,'bold')).pack(anchor='w',padx=20,pady=(4,2))
        ttk.Label(self,text='추천 이유 : '+reason,wraplength=700,justify='left').pack(anchor='w',padx=20,pady=(0,10))
        if rec=='TBD':
            self.var=tk.StringVar(value='TBD'); ttk.Label(self,text='원인이 없으므로 이슈기인은 TBD로 반영됩니다.').pack(anchor='w',padx=20,pady=5)
        else:
            self.var=tk.StringVar(value=rec)
            row=ttk.Frame(self); row.pack(anchor='w',padx=20,pady=5)
            for value in ORIGINS:
                label=value + ('  ← 추천' if value==rec else '')
                ttk.Radiobutton(row,text=label,variable=self.var,value=value).pack(side='left',padx=(0,12))
        buttons=ttk.Frame(self); buttons.pack(pady=16)
        ttk.Button(buttons,text='확인',width=16,command=self.ok).pack(side='left',padx=6)
        ttk.Button(buttons,text='취소',width=16,command=self.cancel).pack(side='left',padx=6)
        self.protocol('WM_DELETE_WINDOW',self.cancel); self.wait_visibility(); self.focus_set(); self.wait_window(self)
    def ok(self): self.result=self.var.get(); self.destroy()
    def cancel(self): self.result=None; self.destroy()

# Page-2 metadata is written by v319._page2_meta. Feed the selected value through d
# and then force only the 이슈기인 destination cell after the existing weekly renderer.
_old_weekly=base.weekly

def _set_origin_in_ppt(path,origin):
    from pptx import Presentation
    prs=Presentation(path)
    if len(prs.slides)<2: return
    sl=prs.slides[1]
    for sh in sl.shapes:
        if not getattr(sh,'has_table',False): continue
        tb=sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                if base.compact(tb.cell(r,c).text)=='이슈기인' and c+1<len(tb.columns):
                    cell=tb.cell(r,c+1); cell.text=origin
                    # Preserve template styling as much as possible; only content changes.
                    prs.save(path); return
    prs.save(path)

def weekly_step8(src,out,d,g,mode):
    msg,saved=_old_weekly(src,out,d,g,mode)
    origin=g.get('_issue_origin_selected','')
    if origin:
        _set_origin_in_ppt(saved,origin)
    return msg,saved
base.weekly=weekly_step8

class RecoveryStep8App(step7.RecoveryStep7App):
    def __init__(self):
        super().__init__(); self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP8')
    def run(self):
        # STEP7 owns the DB phenomenon confirmation and normal update flow.
        # Temporarily wrap its DB dialog so the origin dialog occurs immediately after extraction,
        # before any file is written.
        original_dialog=step7.ProblemChoiceDialog
        parent=self
        class ChainedProblemDialog(original_dialog):
            def __init__(self2,p,summary):
                # Extract once here only for recommendation; STEP7 still performs the authoritative run.
                g=parent.gui(); d=base.extract(g['ppt8d'])
                od=OriginDialog(parent,d)
                if od.result is None:
                    self2.result=None
                    return
                parent._step8_origin=od.result
                super().__init__(p,summary)
        step7.ProblemChoiceDialog=ChainedProblemDialog
        old_gui=self.gui
        def gui_with_origin():
            g=old_gui()
            if hasattr(self,'_step8_origin'): g['_issue_origin_selected']=self._step8_origin
            return g
        self.gui=gui_with_origin
        try:
            return super().run()
        finally:
            step7.ProblemChoiceDialog=original_dialog
            self.gui=old_gui

if __name__=='__main__':
    RecoveryStep8App().mainloop()
