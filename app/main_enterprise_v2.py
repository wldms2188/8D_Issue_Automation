"""Enterprise UI v2: guided inputs, explicit classification confirmation, DB validation and 8D preview."""
import tkinter as tk
from tkinter import ttk

import main_enterprise as ent
import ui_enterprise as ui

base=ent.base
N=ent.N


class EnterpriseAppV2(ent.EnterpriseApp):
    def __init__(self):
        super().__init__()
        self._replace_label_text('예: DUT3 / Sample No.', '예: A1, B1, C2')
        self._replace_label_text('기존값이 있을 때 입력', 'Issue DB 업데이트 시 필수 입력')
        self._install_input_checks()

    def _replace_label_text(self, old, new):
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text') == old: child.configure(text=new)
                except Exception: pass
                walk(child)
        walk(self)

    def _find_labeled_row(self, label_text):
        found=[]
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text') == label_text: found.append(child.master)
                except Exception: pass
                walk(child)
        walk(self)
        return found[0] if found else None

    def _status_mark(self, row):
        mark=tk.Label(row,text='○',bg='white',fg='#AAB5BE',font=('Malgun Gothic',12,'bold'),width=3)
        mark.pack(side='right',padx=(8,0)); return mark

    def _install_input_checks(self):
        self._field_marks={}; self._confirm_vars={}; self._confirm_marks={}
        # Free-text inputs: check automatically when a value is entered.
        for label,key in [('담당팀','team'),('담당자','owner'),('발생 샘플','sample'),('PMS/PLM 이슈번호','plm_no')]:
            row=self._find_labeled_row(label)
            if not row: continue
            mark=self._status_mark(row); self._field_marks[key]=mark
            self.vars[key].trace_add('write',lambda *_args,k=key:self._refresh_field_mark(k))
            self._refresh_field_mark(key)
        # Dropdowns: a default value is not considered confirmed. User must select a value
        # or explicitly click the confirmation circle when the default is correct.
        for label,key in [('폼팩터','form_factor'),('제품 타입','product_type'),('발생처','occurrence_site'),('개발 단계','stage')]:
            row=self._find_labeled_row(label)
            if not row: continue
            var=tk.BooleanVar(value=False); self._confirm_vars[key]=var
            btn=tk.Button(row,text='○',command=lambda k=key:self._toggle_confirm(k),bg='white',fg='#8A98A5',activebackground='white',activeforeground=ui.GREEN,bd=0,font=('Malgun Gothic',13,'bold'),width=3,cursor='hand2')
            btn.pack(side='right',padx=(8,0)); self._confirm_marks[key]=btn
            combo=None
            for child in row.winfo_children():
                if isinstance(child,ttk.Combobox): combo=child; break
            if combo is not None: combo.bind('<<ComboboxSelected>>',lambda e,k=key:self._confirm_dropdown(k),add='+')
        # Keep marks in sync with update target because PLM is conditional.
        self.vars['xlsx'].trace_add('write',lambda *_:self._refresh_field_mark('plm_no'))

    def _refresh_field_mark(self,key):
        mark=self._field_marks.get(key)
        if not mark:return
        value=self.vars[key].get().strip()
        # PLM is optional when Issue DB is not selected.
        if key=='plm_no' and not self.vars.get('xlsx',tk.StringVar()).get().strip():
            mark.configure(text='—',fg='#9BA8B2'); return
        mark.configure(text='✓' if value else '○',fg=ui.GREEN if value else '#AAB5BE')

    def _confirm_dropdown(self,key):
        self._confirm_vars[key].set(True); self._paint_confirm(key)
    def _toggle_confirm(self,key):
        self._confirm_vars[key].set(not self._confirm_vars[key].get()); self._paint_confirm(key)
    def _paint_confirm(self,key):
        yes=self._confirm_vars[key].get(); self._confirm_marks[key].configure(text='✓' if yes else '○',fg=ui.GREEN if yes else '#8A98A5')

    def _unconfirmed_classifications(self):
        names={'form_factor':'폼팩터','product_type':'제품 타입','occurrence_site':'발생처','stage':'개발 단계'}
        return [names[k] for k,v in self._confirm_vars.items() if not v.get()]

    def _show_preview_window(self, d):
        win=tk.Toplevel(self); win.withdraw(); win.title('8D 내용 미리보기'); win.configure(bg='white'); win.resizable(True,True); win.transient(self)
        head=tk.Frame(win,bg=ui.NAVY,height=64); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='8D 내용 미리보기',bg=ui.NAVY,fg='white',font=('Malgun Gothic',14,'bold')).pack(side='left',padx=24); tk.Label(head,text='원본 8D에서 추출된 내용을 확인합니다.',bg=ui.NAVY,fg='#C9D8E6',font=('Malgun Gothic',8)).pack(side='right',padx=24)
        body=tk.Frame(win,bg='white'); body.pack(fill='both',expand=True,padx=22,pady=18); meta=tk.Frame(body,bg='#F4F7FA',highlightbackground=ui.BORDER,highlightthickness=1); meta.pack(fill='x',pady=(0,12))
        for i,(label,key) in enumerate([('이슈명','issue_name'),('고객사','customer'),('과제명','task_name'),('발생일자','occurrence_date')]):
            cell=tk.Frame(meta,bg='#F4F7FA'); cell.grid(row=0,column=i,sticky='nsew',padx=12,pady=10); meta.grid_columnconfigure(i,weight=1); tk.Label(cell,text=label,bg='#F4F7FA',fg=ui.MUTED,font=('Malgun Gothic',8,'bold')).pack(anchor='w'); tk.Label(cell,text=N(d.get(key)) or '-',bg='#F4F7FA',fg=ui.TEXT,font=('Malgun Gothic',9,'bold'),wraplength=180,justify='left').pack(anchor='w',pady=(3,0))
        tk.Label(body,text='추출 결과',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(anchor='w',pady=(0,7)); text_frame=tk.Frame(body,bg='white'); text_frame.pack(fill='both',expand=True); scroll=tk.Scrollbar(text_frame); scroll.pack(side='right',fill='y'); txt=tk.Text(text_frame,wrap='word',yscrollcommand=scroll.set,bg='#FBFCFD',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=14,pady=12); txt.pack(side='left',fill='both',expand=True); scroll.config(command=txt.yview)
        for label,key in [('현상','problem'),('3D 임시조치','temporary_action'),('4D 발생원인','cause_4d'),('4D 유출원인','leak_cause'),('4D 시스템원인','system_cause'),('5D 개선대책','action_5d'),('6D 효과검증','verification_6d'),('고객대응','customer_action'),('수평전개','horizontal_deployment'),('요청사항','request')]:
            txt.insert('end',f'■ {label}\n','heading'); txt.insert('end',(N(d.get(key)) or '추출 내용 없음')+'\n\n')
        txt.tag_configure('heading',font=('Malgun Gothic',9,'bold'),foreground=ui.NAVY,spacing1=5,spacing3=4); txt.configure(state='disabled'); foot=tk.Frame(win,bg='#F6F8FA',height=72); foot.pack(fill='x'); foot.pack_propagate(False); tk.Label(foot,text='※ 미리보기는 파일을 수정하지 않습니다.',bg='#F6F8FA',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='left',padx=22); tk.Button(foot,text='확인',command=win.destroy,bg=ui.BLUE,fg='white',bd=0,font=('Malgun Gothic',9,'bold'),width=12,padx=4,pady=10,cursor='hand2').pack(side='right',padx=22,pady=14); ui.center_window(win,self,900,680); win.minsize(760,560); win.deiconify(); win.grab_set(); win.focus_force(); self.wait_window(win)

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d','').strip(): return ui.warning(self,'입력 확인','8D 원본 PPT를 선택해 주세요.')
        try:
            self.status_var.set('ANALYZING  ·  8D 내용을 추출하고 있습니다...'); self.update_idletasks(); d=base.extract(g['ppt8d']); self.status_var.set('READY  ·  8D 추출 완료'); self._show_preview_window(d)
        except Exception as e: self.status_var.set('ERROR  ·  추출 실패'); ui.error(self,'미리보기 오류',repr(e))

    def run(self):
        g=self.gui()
        if g.get('xlsx','').strip() and not g.get('plm_no','').strip():
            self.status_var.set('READY  ·  PMS/PLM 이슈번호 입력이 필요합니다.'); return ui.warning(self,'PMS/PLM 이슈번호 확인','Issue DB Excel을 업데이트하려면 PMS/PLM 이슈번호를 입력해 주세요.\n\n주간회의 PPT만 업데이트하는 경우에는 입력하지 않아도 됩니다.')
        missing=self._unconfirmed_classifications()
        if missing:
            self.status_var.set('READY  ·  분류 정보 확인이 필요합니다.')
            return ui.warning(self,'분류 정보 확인','아래 분류값을 확인해 주세요.\n\n'+' / '.join(missing)+'\n\n값을 직접 선택하거나, 현재 기본값이 맞으면 오른쪽 ○를 클릭해 ✓로 확인해 주세요.')
        return super().run()


if __name__ == '__main__': EnterpriseAppV2().mainloop()
