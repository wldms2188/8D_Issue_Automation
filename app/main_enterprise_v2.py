"""Enterprise UI v2: input guidance, DB validation, and dedicated 8D preview window."""
import tkinter as tk

import main_enterprise as ent
import ui_enterprise as ui

base=ent.base
N=ent.N


class EnterpriseAppV2(ent.EnterpriseApp):
    def __init__(self):
        super().__init__()
        self._replace_label_text('예: DUT3 / Sample No.', '예: A1, B1, C2')
        self._replace_label_text('기존값이 있을 때 입력', 'Issue DB 업데이트 시 필수 입력')

    def _replace_label_text(self, old, new):
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text') == old:
                        child.configure(text=new)
                except Exception:
                    pass
                walk(child)
        walk(self)

    def _show_preview_window(self, d):
        win=tk.Toplevel(self); win.withdraw(); win.title('8D 내용 미리보기'); win.configure(bg='white'); win.resizable(True,True); win.transient(self)
        head=tk.Frame(win,bg=ui.NAVY,height=64); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='8D 내용 미리보기',bg=ui.NAVY,fg='white',font=('Malgun Gothic',14,'bold')).pack(side='left',padx=24)
        tk.Label(head,text='원본 8D에서 추출된 내용을 확인합니다.',bg=ui.NAVY,fg='#C9D8E6',font=('Malgun Gothic',8)).pack(side='right',padx=24)
        body=tk.Frame(win,bg='white'); body.pack(fill='both',expand=True,padx=22,pady=18)
        meta=tk.Frame(body,bg='#F4F7FA',highlightbackground=ui.BORDER,highlightthickness=1); meta.pack(fill='x',pady=(0,12))
        meta_items=[('이슈명','issue_name'),('고객사','customer'),('과제명','task_name'),('발생일자','occurrence_date')]
        for i,(label,key) in enumerate(meta_items):
            cell=tk.Frame(meta,bg='#F4F7FA'); cell.grid(row=0,column=i,sticky='nsew',padx=12,pady=10); meta.grid_columnconfigure(i,weight=1)
            tk.Label(cell,text=label,bg='#F4F7FA',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w')
            tk.Label(cell,text=N(d.get(key)) or '-',bg='#F4F7FA',fg=ui.TEXT,font=('Malgun Gothic',9,'bold'),wraplength=180,justify='left').pack(anchor='w',pady=(3,0))
        tk.Label(body,text='추출 결과',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(anchor='w',pady=(0,7))
        text_frame=tk.Frame(body,bg='white'); text_frame.pack(fill='both',expand=True)
        scroll=tk.Scrollbar(text_frame); scroll.pack(side='right',fill='y')
        txt=tk.Text(text_frame,wrap='word',yscrollcommand=scroll.set,bg='#FBFCFD',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=14,pady=12); txt.pack(side='left',fill='both',expand=True); scroll.config(command=txt.yview)
        sections=[('현상','problem'),('3D 임시조치','temporary_action'),('4D 발생원인','cause_4d'),('4D 유출원인','leak_cause'),('4D 시스템원인','system_cause'),('5D 개선대책','action_5d'),('6D 효과검증','verification_6d'),('고객대응','customer_action'),('수평전개','horizontal_deployment'),('요청사항','request')]
        for label,key in sections:
            value=N(d.get(key))
            txt.insert('end',f'■ {label}\n','heading'); txt.insert('end',(value or '추출 내용 없음')+'\n\n')
        txt.tag_configure('heading',font=('Malgun Gothic',9,'bold'),foreground=ui.NAVY,spacing1=5,spacing3=4); txt.configure(state='disabled')
        foot=tk.Frame(win,bg='#F6F8FA',height=72); foot.pack(fill='x'); foot.pack_propagate(False)
        tk.Label(foot,text='※ 미리보기는 파일을 수정하지 않습니다.',bg='#F6F8FA',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='left',padx=22)
        tk.Button(foot,text='확인',command=win.destroy,bg=ui.BLUE,fg='white',activebackground='#12598F',activeforeground='white',bd=0,font=('Malgun Gothic',9,'bold'),width=12,padx=4,pady=10,cursor='hand2').pack(side='right',padx=22,pady=14)
        ui.center_window(win,self,900,680); win.minsize(760,560); win.deiconify(); win.grab_set(); win.focus_force(); self.wait_window(win)

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d','').strip():
            return ui.warning(self,'입력 확인','8D 원본 PPT를 선택해 주세요.')
        try:
            self.status_var.set('ANALYZING  ·  8D 내용을 추출하고 있습니다...'); self.update_idletasks()
            d=base.extract(g['ppt8d'])
            self.status_var.set('READY  ·  8D 추출 완료')
            self._show_preview_window(d)
        except Exception as e:
            self.status_var.set('ERROR  ·  추출 실패')
            ui.error(self,'미리보기 오류',repr(e))

    def run(self):
        g=self.gui()
        if g.get('xlsx','').strip() and not g.get('plm_no','').strip():
            self.status_var.set('READY  ·  PMS/PLM 이슈번호 입력이 필요합니다.')
            return ui.warning(
                self,
                'PMS/PLM 이슈번호 확인',
                'Issue DB Excel을 업데이트하려면 PMS/PLM 이슈번호를 입력해 주세요.\n\n'
                '주간회의 PPT만 업데이트하는 경우에는 입력하지 않아도 됩니다.'
            )
        return super().run()


if __name__ == '__main__':
    EnterpriseAppV2().mainloop()
