"""Enterprise UI v2: balanced focus panels and always-visible workflow progress."""
import tkinter as tk
from tkinter import ttk
from pathlib import Path

import main_enterprise as ent
import ui_enterprise as ui

base=ent.base
N=ent.N

def progress_state(status,current=0):
    s=str(status or '').strip(); low=s.lower(); current=int(current or 0)
    if 'error' in low or '실패' in s:return current,'오류 발생'
    if '8d 추출 완료' in low:return max(current,25),'8D 추출 완료'
    if 'complete' in low or '업데이트 완료' in s:return 100,'업데이트 완료'
    if '주간' in s or 'weekly' in low or ('ppt' in low and ('업데이트' in s or '저장' in s)):return max(current,85),'주간회의 PPT 업데이트 중'
    if ('excel' in low and ('업데이트' in s or '저장' in s)) or 'issue db 업데이트' in low:return max(current,70),'Issue DB 업데이트 중'
    if '이슈기인' in s:return max(current,55),'이슈기인 확인 중'
    if '현상' in s and ('확인' in s or '선택' in s):return max(current,45),'Issue DB 현상 확인 중'
    if '매칭' in s or 'match' in low:return max(current,30),'기존 이슈 매칭 확인 중'
    if '추출' in s or 'analyz' in low:return max(current,15),'8D 내용 추출 중'
    if 'running' in low:return max(10,current),(s.split('·',1)[-1].strip() if '·' in s else '처리 중')
    if 'ready' in low:return 0,'실행 대기'
    return current,(s or '처리 중')


class EnterpriseAppV2(ent.EnterpriseApp):
    def __init__(self):
        super().__init__()
        self._replace_label_text('예: DUT3 / Sample No.', '예: A1, B1, C2')
        self._replace_label_text('기존값이 있을 때 입력', 'Issue DB 업데이트 시 필수 입력')
        self._install_input_checks()
        self._compact_layout()
        self._balance_and_focus_cards()
        self._install_progress_panel()

    def _compact_layout(self):
        self.update_idletasks()
        sw,sh=self.winfo_screenwidth(),self.winfo_screenheight()
        w=min(1180,max(1040,sw-80)); h=min(820,max(740,sh-70))
        self.minsize(1040,720)
        self.geometry(f'{w}x{h}+{max(0,(sw-w)//2)}+{max(0,(sh-h)//2)}')
        try:self.log.configure(height=2)
        except Exception:pass
        for z in getattr(self,'dropzones',{}).values():
            try:z.pack_configure(pady=2)
            except Exception:pass

    def _find_section_card(self,number):
        found=[]
        def walk(w):
            for c in w.winfo_children():
                try:
                    if c.cget('text')==number:found.append(c)
                except Exception:pass
                walk(c)
        walk(self)
        if not found:return None
        w=found[0]
        while w is not self:
            try:
                if isinstance(w,tk.Frame) and w.cget('bg')=='white' and int(w.cget('highlightthickness'))==1:return w
            except Exception:pass
            w=w.master
        return None

    def _balance_and_focus_cards(self):
        self.card01=self._find_section_card('01'); self.card02=self._find_section_card('02')
        if not self.card01 or not self.card02 or self.card01.master is not self.card02.master:return
        p=self.card01.master; self._cards_parent=p
        try:
            self.card01.pack_forget(); self.card02.pack_forget()
            p.grid_columnconfigure(0,weight=1,uniform='half',minsize=0)
            p.grid_columnconfigure(1,weight=1,uniform='half',minsize=0)
            p.grid_rowconfigure(0,weight=1)
            self.card01.grid(row=0,column=0,sticky='nsew',padx=(0,7))
            self.card02.grid(row=0,column=1,sticky='nsew',padx=(7,0))
        except Exception:return
        # Bind both mouse and focus recursively. Combobox internal clicks are also caught globally below.
        def bind_tree(w,side):
            w.bind('<ButtonPress-1>',lambda e,s=side:self._expand_card(s),add='+')
            w.bind('<FocusIn>',lambda e,s=side:self._expand_card(s),add='+')
            for c in w.winfo_children():bind_tree(c,side)
        bind_tree(self.card01,1); bind_tree(self.card02,2)
        self.bind_all('<ButtonPress-1>',self._route_card_click,add='+')
        self.bind_all('<FocusIn>',self._route_card_focus,add='+')
        self._set_card_ratio(0)

    def _is_descendant(self,w,parent):
        if parent is None:return False
        while w is not None:
            if w is parent:return True
            try:w=w.master
            except Exception:return False
        return False

    def _route_card_click(self,event):
        if self._is_descendant(event.widget,getattr(self,'card02',None)):self._set_card_ratio(2)
        elif self._is_descendant(event.widget,getattr(self,'card01',None)):self._set_card_ratio(1)

    def _route_card_focus(self,event):
        if self._is_descendant(event.widget,getattr(self,'card02',None)):self._set_card_ratio(2)
        elif self._is_descendant(event.widget,getattr(self,'card01',None)):self._set_card_ratio(1)

    def _set_card_ratio(self,active=0):
        p=getattr(self,'_cards_parent',None)
        if not p:return
        try:
            # uniform MUST be cleared on expansion; otherwise Tk forces equal widths regardless of weights.
            p.grid_columnconfigure(0,uniform=''); p.grid_columnconfigure(1,uniform='')
            if active==1:
                p.grid_columnconfigure(0,weight=13); p.grid_columnconfigure(1,weight=7)
            elif active==2:
                p.grid_columnconfigure(0,weight=7); p.grid_columnconfigure(1,weight=13)
            else:
                p.grid_columnconfigure(0,weight=1,uniform='half'); p.grid_columnconfigure(1,weight=1,uniform='half')
            p.update_idletasks()
        except Exception:pass

    def _expand_card(self,side):self._set_card_ratio(side)

    def _install_progress_panel(self):
        """Replace the flexible result body with a fixed visible progress strip + compact log."""
        result=self.log.master
        # Keep result area tall enough so Windows DPI scaling cannot clip the progress bar.
        try:
            result.pack_propagate(False)
            result.configure(height=112)
        except Exception:pass
        self.progress_value=tk.DoubleVar(value=0); self.progress_text=tk.StringVar(value='0% · 실행 대기')
        panel=tk.Frame(result,bg='white',height=34)
        panel.pack(fill='x',side='top',padx=16,pady=(0,5),before=self.log); panel.pack_propagate(False)
        tk.Label(panel,text='진행률',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8,'bold'),width=7,anchor='w').pack(side='left')
        self.progress_canvas=tk.Canvas(panel,height=14,bg='#D9E2E9',highlightthickness=1,highlightbackground='#B9C8D3',bd=0)
        self.progress_canvas.pack(side='left',fill='x',expand=True,padx=(5,12),pady=9)
        self.progress_fill=self.progress_canvas.create_rectangle(0,0,0,14,fill='#1769AA',outline='')
        tk.Label(panel,textvariable=self.progress_text,bg='white',fg=ui.NAVY,font=('Malgun Gothic',8,'bold'),width=27,anchor='e').pack(side='right')
        self.progress_canvas.bind('<Configure>',lambda e:self._paint_progress_bar())
        self.status_var.trace_add('write',self._sync_progress_from_status)
        self._sync_progress_from_status()

    def _paint_progress_bar(self):
        try:
            self.update_idletasks()
            width=max(2,self.progress_canvas.winfo_width()); height=max(2,self.progress_canvas.winfo_height())
            pct=max(0.0,min(100.0,float(self.progress_value.get())))
            self.progress_canvas.coords(self.progress_fill,0,0,max(0,width*pct/100.0),height)
            self.progress_canvas.tag_raise(self.progress_fill)
        except Exception:pass

    def _set_progress(self,pct,label):
        pct=max(0,min(100,int(pct))); self.progress_value.set(pct); self.progress_text.set(f'{pct}% · {label}')
        self._paint_progress_bar()
        try:self.update()
        except Exception:pass

    def _sync_progress_from_status(self,*_):
        pct,label=progress_state(self.status_var.get(),self.progress_value.get())
        self._set_progress(pct,label)

    def _replace_label_text(self,old,new):
        def walk(w):
            for c in w.winfo_children():
                try:
                    if c.cget('text')==old:c.configure(text=new)
                except Exception:pass
                walk(c)
        walk(self)

    def _find_labeled_row(self,label_text):
        found=[]
        def walk(w):
            for c in w.winfo_children():
                try:
                    if c.cget('text')==label_text:found.append(c.master)
                except Exception:pass
                walk(c)
        walk(self); return found[0] if found else None

    def _status_mark(self,row):
        m=tk.Label(row,text='○',bg='white',fg='#AAB5BE',font=('Malgun Gothic',11,'bold'),width=2); m.pack(side='right',padx=(5,0)); return m

    def _install_input_checks(self):
        self._field_marks={}; self._confirm_vars={}; self._confirm_marks={}
        for label,key in [('담당팀','team'),('담당자','owner'),('발생 샘플','sample'),('PMS/PLM 이슈번호','plm_no')]:
            row=self._find_labeled_row(label)
            if not row:continue
            row.pack_configure(pady=1); m=self._status_mark(row); self._field_marks[key]=m
            self.vars[key].trace_add('write',lambda *_a,k=key:self._refresh_field_mark(k)); self._refresh_field_mark(key)
        first=None
        for label,key in [('폼팩터','form_factor'),('제품 타입','product_type'),('발생처','occurrence_site'),('개발 단계','stage')]:
            row=self._find_labeled_row(label)
            if not row:continue
            if first is None:first=row
            v=tk.BooleanVar(value=False); self._confirm_vars[key]=v
            b=tk.Button(row,text='○',command=lambda k=key:self._toggle_confirm(k),bg='white',fg='#8A98A5',activebackground='white',bd=0,font=('Malgun Gothic',11,'bold'),width=2,cursor='hand2'); b.pack(side='right',padx=(5,0)); self._confirm_marks[key]=b
            combo=next((c for c in row.winfo_children() if isinstance(c,ttk.Combobox)),None)
            if combo is not None:combo.bind('<<ComboboxSelected>>',lambda e,k=key:self._confirm_dropdown(k),add='+')
        if first is not None:
            tk.Label(first.master,text='※ 분류값 확인: 값 변경 시 자동 ✓  |  기본값이 맞으면 오른쪽 ○ 클릭 → 선택 완료 ✓',bg='white',fg='#5D7488',font=('Malgun Gothic',8,'bold'),anchor='e').pack(fill='x',pady=(0,2),before=first)
        self.vars['xlsx'].trace_add('write',lambda *_:self._refresh_field_mark('plm_no'))

    def _refresh_field_mark(self,key):
        m=self._field_marks.get(key)
        if not m:return
        val=self.vars[key].get().strip()
        if key=='plm_no' and not self.vars['xlsx'].get().strip():m.configure(text='—',fg='#9BA8B2'); return
        m.configure(text='✓' if val else '○',fg=ui.GREEN if val else '#AAB5BE')
    def _confirm_dropdown(self,key):self._confirm_vars[key].set(True); self._paint_confirm(key)
    def _toggle_confirm(self,key):self._confirm_vars[key].set(not self._confirm_vars[key].get()); self._paint_confirm(key)
    def _paint_confirm(self,key):
        yes=self._confirm_vars[key].get(); self._confirm_marks[key].configure(text='✓' if yes else '○',fg=ui.GREEN if yes else '#8A98A5')
    def _unconfirmed_classifications(self):
        names={'form_factor':'폼팩터','product_type':'제품 타입','occurrence_site':'발생처','stage':'개발 단계'}
        return [names[k] for k,v in self._confirm_vars.items() if not v.get()]

    def _show_preview_window(self,d):
        win=tk.Toplevel(self); win.withdraw(); win.title('8D 내용 미리보기'); win.configure(bg='white'); win.transient(self)
        head=tk.Frame(win,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='8D 내용 미리보기',bg=ui.NAVY,fg='white',font=('Malgun Gothic',13,'bold')).pack(side='left',padx=22)
        body=tk.Frame(win,bg='white'); body.pack(fill='both',expand=True,padx=22,pady=16)
        txt=tk.Text(body,wrap='word',bg='#FBFCFD',fg=ui.TEXT,relief='flat',highlightthickness=1,highlightbackground=ui.BORDER,font=('Malgun Gothic',9),padx=14,pady=12); txt.pack(fill='both',expand=True)
        for label,key in [('2D 현상','problem'),('3D 임시조치','temporary_action'),('4D 발생원인','cause_4d'),('4D 유출원인','leak_cause'),('4D 시스템원인','system_cause'),('5D 개선대책','action_5d'),('6D 효과검증','verification_6d'),('고객대응','customer_action'),('수평전개','horizontal_deployment'),('요청사항','request')]:
            txt.insert('end',f'■ {label}\n','h'); txt.insert('end',(N(d.get(key)) or '추출 내용 없음')+'\n\n')
        txt.tag_configure('h',font=('Malgun Gothic',9,'bold'),foreground=ui.NAVY); txt.configure(state='disabled')
        foot=tk.Frame(win,bg='#F6F8FA',height=60); foot.pack(fill='x'); foot.pack_propagate(False)
        tk.Button(foot,text='확인',command=win.destroy,bg=ui.BLUE,fg='white',bd=0,font=('Malgun Gothic',9,'bold'),width=12,pady=8).pack(side='right',padx=20,pady=12)
        ui.center_window(win,self,860,640); win.deiconify(); win.grab_set(); self.wait_window(win)

    def preview(self):
        g=self.gui()
        if not g.get('ppt8d','').strip():return ui.warning(self,'입력 확인','8D 원본 PPT를 선택해 주세요.')
        try:
            self.status_var.set('ANALYZING · 8D 내용을 추출하고 있습니다...'); self._set_progress(15,'8D 내용 추출 중')
            d=base.extract(g['ppt8d'])
            self._last_preview_path=g['ppt8d']
            self.log.delete('1.0','end')
            self.log.insert('end','[ 8D 추출 완료 ]\n'+'─'*72+'\n')
            self.log.insert('end',f'원본 파일       : {Path(g["ppt8d"]).name}\n')
            self.log.insert('end','상태            : 8D 내용 추출 완료 · 미리보기 확인 가능\n')
            self.log.insert('end','※ 아래 [8D 내용 미리보기] 창에서 2D~8D 추출 내용을 확인해 주세요.\n')
            self.status_var.set('READY · 8D 추출 완료')
            self._show_preview_window(d)
        except Exception as e:self.status_var.set('ERROR · 추출 실패'); ui.error(self,'미리보기 오류',repr(e))

    def run(self):
        g=self.gui()
        if g.get('xlsx','').strip() and not g.get('plm_no','').strip():self.status_var.set('READY · PMS/PLM 이슈번호 입력이 필요합니다.'); return ui.warning(self,'PMS/PLM 이슈번호 확인','Issue DB Excel을 업데이트하려면 PMS/PLM 이슈번호를 입력해 주세요.\n\n주간회의 PPT만 업데이트하는 경우에는 입력하지 않아도 됩니다.')
        missing=self._unconfirmed_classifications()
        if missing:self.status_var.set('READY · 분류 정보 확인이 필요합니다.'); return ui.warning(self,'분류 정보 확인','아래 분류값을 확인해 주세요.\n\n'+' / '.join(missing)+'\n\n값을 직접 선택하거나, 현재 기본값이 맞으면 오른쪽 ○를 클릭해 ✓로 확인해 주세요.')
        self._set_progress(10,'실행 준비 중')
        return super().run()


if __name__=='__main__':EnterpriseAppV2().mainloop()
