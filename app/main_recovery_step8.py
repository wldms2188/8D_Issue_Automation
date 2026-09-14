import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook

import main_recovery_step7 as step7
import main_v310 as v310

base=step7.base
N=v310.N

ORIGINS=('부품','설계','공정','기타','논의 중')


def _cause_text(d):
    parts=[]
    for label,key in [('발생원인','cause_4d'),('유출원인','leak_cause'),('시스템원인','system_cause')]:
        s=N(d.get(key))
        if s:
            parts.append(f'{label} : {s}')
    return '\n'.join(parts)


def _recommend_origin(d):
    text=_cause_text(d)
    if not text:
        return 'TBD','4D 원인 내용이 아직 없어 TBD로 표시합니다.'

    q=text.lower().replace(' ','')
    keys={
        '부품':('부품','자재','소재','원재료','supplier','협력사','입고','component','cell','셀불량'),
        '설계':('설계','design','도면','공차','사양','spec','구조','치수','강성','간섭','설계마진'),
        '공정':('공정','작업','조립','체결','토크','용접','검사','검출','설비','조건관리','작업자','생산','가공','도포','압착','공정조건'),
        '기타':('운송','보관','취급','고객','외부','환경','사용조건'),
    }
    scores={cat:sum(1 for w in words if w.lower().replace(' ','') in q) for cat,words in keys.items()}
    best=max(scores,key=scores.get)
    if scores[best]==0:
        return '논의 중','원인 내용은 있으나 부품/설계/공정/기타로 명확히 분류할 근거가 부족합니다.'

    tied=[k for k,v in scores.items() if v==scores[best]]
    if len(tied)>1:
        return '논의 중','원인 내용에 여러 이슈기인 범주의 단서가 함께 있어 추가 논의가 필요합니다.'

    evidence=[w for w in keys[best] if w.lower().replace(' ','') in q][:3]
    return best, f"8D 원인 내용에서 {', '.join(evidence)} 관련 표현이 확인되어 {best}을(를) 추천합니다."


class OriginDialog(tk.Toplevel):
    def __init__(self,parent,d):
        super().__init__(parent)
        self.result=None
        self.title('이슈기인 확인')
        self.geometry('760x590')
        self.transient(parent)
        self.grab_set()

        cause=_cause_text(d)
        rec,reason=_recommend_origin(d)

        ttk.Label(self,text='8D 원인 내용을 기준으로 이슈기인을 추천했습니다.',font=('',10,'bold')).pack(anchor='w',padx=20,pady=(18,8))
        ttk.Label(self,text='8D 원인 내용').pack(anchor='w',padx=20)
        box=tk.Text(self,height=12,wrap='word')
        box.pack(fill='both',expand=True,padx=20,pady=(4,10))
        box.insert('1.0',cause or '(4D 원인 내용 없음)')
        box.configure(state='disabled')

        ttk.Label(self,text=f'추천 이슈기인 : {rec}',font=('',11,'bold')).pack(anchor='w',padx=20,pady=(4,2))
        ttk.Label(self,text='추천 이유 : '+reason,wraplength=700,justify='left').pack(anchor='w',padx=20,pady=(0,10))

        self.var=tk.StringVar(value=rec)
        if rec=='TBD':
            ttk.Label(self,text='원인이 없으므로 이슈기인은 TBD로 반영됩니다.').pack(anchor='w',padx=20,pady=5)
        else:
            row=ttk.Frame(self)
            row.pack(anchor='w',padx=20,pady=5)
            for value in ORIGINS:
                label=value + ('  ← 추천' if value==rec else '')
                ttk.Radiobutton(row,text=label,variable=self.var,value=value).pack(side='left',padx=(0,12))

        buttons=ttk.Frame(self)
        buttons.pack(pady=16)
        ttk.Button(buttons,text='확인',width=16,command=self.ok).pack(side='left',padx=6)
        ttk.Button(buttons,text='취소',width=16,command=self.cancel).pack(side='left',padx=6)
        self.protocol('WM_DELETE_WINDOW',self.cancel)
        self.wait_visibility()
        self.focus_set()
        self.wait_window(self)

    def ok(self):
        self.result=self.var.get()
        self.destroy()

    def cancel(self):
        self.result=None
        self.destroy()


_old_weekly=base.weekly


def _set_cell_text_preserve(cell,text):
    tf=cell.text_frame
    runs=[r for p in tf.paragraphs for r in p.runs]
    if runs:
        runs[0].text=N(text)
        for r in runs[1:]:
            r.text=''
    else:
        cell.text=N(text)


def _set_origin_in_ppt(path,origin):
    from pptx import Presentation
    prs=Presentation(path)
    if len(prs.slides)<2:
        return False

    sl=prs.slides[1]
    for sh in v310.walk(sl):
        if not getattr(sh,'has_table',False):
            continue
        tb=sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                if base.compact(tb.cell(r,c).text)=='이슈기인' and c+1<len(tb.columns):
                    _set_cell_text_preserve(tb.cell(r,c+1),origin)
                    prs.save(path)
                    return True
    return False


def weekly_step8(src,out,d,g,mode):
    msg,saved=_old_weekly(src,out,d,g,mode)
    origin=N(g.get('_issue_origin_selected'))
    if origin:
        _set_origin_in_ppt(saved,origin)
    return msg,saved


base.weekly=weekly_step8


class RecoveryStep8App(step7.RecoveryStep7App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP8 FIX1')

    def run(self):
        g=self.gui()
        if any(not g[k] or not os.path.exists(g[k]) for k in ['ppt8d','pptweekly','xlsx']):
            return messagebox.showwarning('확인','8D PPT, 주간회의 PPT, 이슈 DB Excel을 모두 선택하세요.')

        try:
            d=base.extract(g['ppt8d'])
            mode=self.mode.get()

            # 1) Decide issue origin first and store it in the SAME g dict used for saving.
            od=OriginDialog(self,d)
            if od.result is None:
                self.log.delete('1.0','end')
                self.log.insert('end','업데이트가 취소되었습니다.')
                return
            g['_issue_origin_selected']=od.result

            # 2) Keep STEP7 Issue-DB phenomenon choice exactly as validated.
            summary=step7._db_summary(d)
            pd=step7.ProblemChoiceDialog(self,summary)
            if pd.result is None:
                self.log.delete('1.0','end')
                self.log.insert('end','업데이트가 취소되었습니다.')
                return
            g['_db_problem_selected']=summary if pd.result=='summary' else N(d.get('problem'))

            # 3) Normal update flow using the same g with both selections included.
            out=Path(g['xlsx']).parent/'자동화_결과'
            out.mkdir(exist_ok=True)
            xo=out/(Path(g['xlsx']).stem+'_업데이트.xlsx')
            po=out/(Path(g['pptweekly']).stem+'_업데이트.pptx')

            wb=load_workbook(g['xlsx'])
            ws=wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active
            r,_=base.find(ws,d,g)
            if mode=='existing' and not r:
                return messagebox.showwarning('업데이트 중단','기존 이슈를 Excel에서 정확하게 찾지 못했습니다. PMS/PLM 이슈번호를 입력하거나 이슈 정보를 확인하세요.')
            if mode=='new' and r and not messagebox.askyesno('중복 가능성',f'유사 이슈(row {r})가 있습니다. 그래도 신규 추가할까요?'):
                return

            a,_=base.update_excel(g['xlsx'],xo,d,g,new=(mode=='new'))
            b,saved_ppt=base.weekly(g['pptweekly'],po,d,g,mode)

            phenomenon_label='현상 요약' if pd.result=='summary' else '전체 내용'
            self.log.delete('1.0','end')
            self.log.insert('end',
                f'=== 완료 ===\n'
                f'이슈 구분: {mode}\n'
                f'Issue DB 현상: {phenomenon_label}\n'
                f'이슈기인: {od.result}\n'
                f'상태: {base.status(d)}\n\n'
                f'{a}\n{b}\n\n'
                f'결과: {out}'
            )
            messagebox.showinfo('완료',f'자동 업데이트가 완료되었습니다.\n\n이슈기인: {od.result}\n\n{out}')

        except Exception as e:
            messagebox.showerror('실행 오류',repr(e))


if __name__=='__main__':
    RecoveryStep8App().mainloop()
