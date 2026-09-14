import os
from pathlib import Path
from tkinter import messagebox
from openpyxl import load_workbook

import main_recovery_step14_fix2 as core
import main_recovery_step9 as step9
import main_recovery_step8 as step8
import main_recovery_step7 as step7
import main_recovery_step10 as step10

base=core.base
N=core.N
OCCURRENCE_SITES=step9.OCCURRENCE_SITES


class FinalApp(core.RecoveryStep14Fix5App):
    def __init__(self):
        super().__init__()
        self.title('8D Issue Automation 1.0 | Pack 개발품질')
        try:
            self.geometry('900x720')
            self.minsize(820,650)
        except Exception:
            pass

    def run(self):
        g=self.gui()
        ppt8d=g.get('ppt8d','')
        weekly=g.get('pptweekly','')
        xlsx=g.get('xlsx','')

        if not ppt8d or not os.path.exists(ppt8d):
            return messagebox.showwarning('입력 확인','8D 원본 PPT는 반드시 선택해 주세요.',parent=self)
        if not weekly and not xlsx:
            return messagebox.showwarning('출력 확인','주간회의 PPT 또는 이슈 DB Excel 중 하나 이상을 선택해 주세요.',parent=self)
        if weekly and not os.path.exists(weekly):
            return messagebox.showwarning('입력 확인','선택한 주간회의 PPT 파일을 찾을 수 없습니다.',parent=self)
        if xlsx and not os.path.exists(xlsx):
            return messagebox.showwarning('입력 확인','선택한 이슈 DB Excel 파일을 찾을 수 없습니다.',parent=self)
        if g.get('occurrence_site') not in OCCURRENCE_SITES:
            return messagebox.showwarning('입력 확인','발생처를 선택해 주세요.',parent=self)

        try:
            d=base.extract(ppt8d)
            mode=self.mode.get()
            do_weekly=bool(weekly)
            do_excel=bool(xlsx)

            excel_row=None
            if do_excel:
                wb=load_workbook(xlsx)
                ws=wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active
                excel_row,_=base.find(ws,d,g)

                if mode=='existing' and not excel_row:
                    if do_weekly:
                        if messagebox.askyesno(
                            '기존 이슈 미확인',
                            'Issue DB에서 동일 이슈를 찾지 못했습니다.\n\n'
                            'Excel은 변경하지 않고 주간회의 PPT만 업데이트할까요?',
                            parent=self):
                            do_excel=False
                        else:
                            return
                    else:
                        return messagebox.showwarning(
                            '기존 이슈 미확인',
                            'Issue DB에서 동일 이슈를 찾지 못했습니다.\n기존 이슈는 임의로 신규 행을 만들지 않습니다.',
                            parent=self)

                if mode=='new' and excel_row:
                    if not messagebox.askyesno('중복 가능성',f'유사 이슈(row {excel_row})가 있습니다. 그래도 신규 추가할까요?',parent=self):
                        return

            # Weekly-only dialog: issue origin is only needed for the weekly detail page.
            if do_weekly:
                od=step8.OriginDialog(self,d)
                if od.result is None:
                    return
                g['_issue_origin_selected']=od.result
            else:
                od=None

            pd_result=None
            final_status=None
            if do_excel:
                summary=step7._db_summary(d)
                pd=step7.ProblemChoiceDialog(self,summary)
                if pd.result is None:
                    return
                pd_result=pd.result
                g['_db_problem_selected']=summary if pd.result=='summary' else N(d.get('problem'))

                judged,reason=step9._judge_issue_status(d)
                final_status=judged
                if judged=='close':
                    sixd=reason or '(6D 내용 없음)'
                    prompt=(
                        '6D 내용 기준으로 close로 판단되었습니다.\n\n'
                        '판단 이유(6D 원문)\n────────────────────\n'
                        f'{sixd}\n────────────────────\n\n'
                        '이슈 상태를 close로 처리하시겠습니까?\n아니오를 선택하면 open으로 처리합니다.'
                    )
                    if not messagebox.askyesno('이슈 상태 확인',prompt,parent=self):
                        final_status='open'
                g['_issue_status_selected']=final_status

            # Put results next to whichever destination file was actually selected.
            anchor=xlsx if do_excel else weekly
            out=Path(anchor).parent/'자동화_결과'
            out.mkdir(exist_ok=True)

            results=[]
            if do_excel:
                xo=out/(Path(xlsx).stem+'_업데이트.xlsx')
                a,_=base.update_excel(xlsx,xo,d,g,new=(mode=='new'))
                results.append(a)

            if do_weekly:
                po=out/(Path(weekly).stem+'_업데이트.pptx')
                b,_=base.weekly(weekly,po,d,g,mode)
                results.append(b)

            self.log.delete('1.0','end')
            targets=[]
            if do_excel: targets.append('Issue DB Excel')
            if do_weekly: targets.append('주간회의 PPT')
            lines=['=== 업데이트 완료 ===',f'이슈 구분: {mode}',f'업데이트 대상: {" + ".join(targets)}']
            if do_weekly and od:
                lines.append(f'이슈기인: {od.result}')
            if do_excel:
                lines.extend([
                    f'Issue DB 현상: {"현상 요약" if pd_result=="summary" else "전체 내용"}',
                    f'Issue DB 상태: {final_status}',
                ])
            lines.extend(['']+results+['',f'결과 폴더: {out}'])
            self.log.insert('end','\n'.join(lines))

            messagebox.showinfo(
                '완료',
                '선택한 자료만 업데이트했습니다.\n\n'
                + '\n'.join('✓ '+x for x in targets)
                + f'\n\n결과 폴더\n{out}',
                parent=self)

        except Exception as e:
            messagebox.showerror('실행 오류',repr(e),parent=self)


if __name__=='__main__':
    FinalApp().mainloop()
