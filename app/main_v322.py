# 8D Issue Automation v3.2.2
# 6D-driven issue status + explicit close confirmation before update.
import re

import main_v321 as v321

base = v321.base
N = v321.N
C = v321.C


def _status_reason(d):
    """Return (status, reason) using 6D only.

    Priority:
    1) progress/future wording -> open
    2) completion/confirmation wording -> close
    3) any other meaningful 6D content -> close
    4) blank 6D -> open
    """
    raw=N(d.get('verification_6d'))
    q=C(raw)
    if not q:
        return 'open', '6D 내용이 기재되어 있지 않아 open으로 판단했습니다.'

    # 진행/예정 표현은 완료/확인보다 우선한다.
    progress_patterns=(
        r'진행\s*중', r'검증\s*중', r'확인\s*중', r'모니터링\s*중',
        r'시험\s*중', r'적용\s*중', r'조치\s*중', r'분석\s*중',
        r'예정', r'계획', r'추후', r'진행중', r'검증중', r'확인중', r'모니터링중'
    )
    for pat in progress_patterns:
        if re.search(pat, raw, re.I):
            return 'open', f'6D에 진행/예정 표현("{re.search(pat, raw, re.I).group(0)}")이 있어 open으로 판단했습니다.'

    if re.search(r'완료',raw,re.I):
        return 'close', '6D에 "완료" 표현이 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'
    if re.search(r'확인',raw,re.I):
        return 'close', '6D에 "확인" 표현이 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'

    return 'close', '6D에 검증/확인 결과 내용이 기재되어 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'


def _excel_status(d,g=None):
    # User may reject the automatic close judgment; in that case keep it open.
    if g and g.get('_status_override') in ('open','close'):
        return g['_status_override']
    return _status_reason(d)[0]


def _write_row(ws,r,d,g):
    y,m=base.parse_year_month(d.get('occurrence_date'))
    cause='\n'.join(x for x in [d.get('cause_4d'),d.get('leak_cause'),d.get('system_cause')] if N(x))
    vals={
        2:v321.datetime.date.today(),3:g.get('plm_no',''),4:g.get('form_factor',''),5:g.get('product_type',''),
        6:y,7:v321._month_label(m),8:g.get('team',''),9:g.get('owner',''),10:base.task(d),11:g.get('sample',''),
        12:g.get('stage',''),13:g.get('occurrence_site') or d.get('occurrence_site',''),14:N(d.get('problem')),15:'',
        16:cause,17:N(d.get('action_5d')),18:_excel_status(d,g)
    }
    for c,v in vals.items():
        ws.cell(r,c).value=v
    v321._add_excel_photo(ws,r,base.representative_image(d))
    return vals[18]


# Patch the shared write path used by update_excel.
v321._excel_status=lambda d: _excel_status(d)
v321._write_row=_write_row
base.write_row=_write_row


class App(v321.App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.2')

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d'):
            return v321.messagebox.showwarning('확인','8D PPT를 선택하세요.')
        try:
            d=base.extract(g['ppt8d'])
            kind,event=v321._event_name(d)
            st,reason=_status_reason(d)
            self.log.delete('1.0','end')
            self.log.insert('end','=== 8D 추출 결과 ===\n')
            for k in ['issue_name','task_name','customer','occurrence_date','problem','cause_4d','leak_cause','system_cause','action_5d','verification_6d','spread_7d']:
                self.log.insert('end',f'{k}: {d.get(k)}\n')
            self.log.insert('end',f'\n시험/빌드 판정: {kind} / {event}\n이슈기인 기본값: {"선택 필요" if v321._cause_known(d) else "TBD"}\nDB 상태: {st}\n상태 판정 이유: {reason}')
        except Exception as e:
            v321.messagebox.showerror('오류',repr(e))

    def _prepare_choices(self,d,g):
        g=super()._prepare_choices(d,g)
        if g is None:
            return None

        st,reason=_status_reason(d)
        if st=='close':
            six=N(d.get('verification_6d'))
            msg=(
                '6D까지 작성된 내용을 기준으로 이 이슈를 close로 판단했습니다.\n\n'
                f'[6D 내용]\n{six}\n\n'
                f'[판단 이유]\n{reason}\n\n'
                '이슈 상태를 close로 처리할까요?\n\n'
                '예: close로 처리\n아니오: open으로 처리'
            )
            ans=v321.messagebox.askyesno('이슈 상태 최종 확인',msg)
            g['_status_override']='close' if ans else 'open'
        else:
            g['_status_override']='open'
        return g

    def run(self):
        # Inherited run() dynamically calls this class's _prepare_choices().
        return super().run()


# Keep all v3.2.1 GUI/weekly/extraction behavior except the status rule above.
base.extract=v321.base.extract
base.weekly=v321.weekly

if __name__=='__main__':
    App().mainloop()
