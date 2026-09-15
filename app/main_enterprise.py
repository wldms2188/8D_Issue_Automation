"""Enterprise presentation layer for 8D Issue Automation.
Keeps the validated FinalApp processing methods and replaces only the main window UI.
"""
import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path

import main_final as legacy
import ui_enterprise as ui

try:
    from tkinterdnd2 import TkinterDnD
    _RootBase=TkinterDnD.Tk
except Exception:
    _RootBase=tk.Tk


class EnterpriseApp(legacy.FinalApp, _RootBase):
    def __init__(self):
        # Initialize DnD-aware Tk root directly; do not run legacy UI builder.
        _RootBase.__init__(self)
        self.title('8D Issue Automation 1.0 | Pack 개발품질')
        self.geometry('1180x850'); self.minsize(1080,780); self.configure(bg=ui.BG)
        self.vars={}; self.mode=tk.StringVar(value='existing')
        self.status_var=tk.StringVar(value='READY  ·  8D 원본을 선택해 주세요.')
        self._configure_styles(); self._build_enterprise_ui(); self._center_main()

    def _center_main(self):
        self.update_idletasks(); w=1180; h=850
        x=max(0,(self.winfo_screenwidth()-w)//2); y=max(0,(self.winfo_screenheight()-h)//2)
        self.geometry(f'{w}x{h}+{x}+{y}')

    def _build_enterprise_ui(self):
        head=tk.Frame(self,bg=ui.NAVY,height=96); head.pack(fill='x'); head.pack_propagate(False)
        tk.Label(head,text='8D Issue Automation',bg=ui.NAVY,fg='white',font=('Malgun Gothic',20,'bold')).place(x=32,y=18)
        tk.Label(head,text='Issue Management Productivity Tool  ·  v1.0',bg=ui.NAVY,fg='#C9D8E6',font=('Malgun Gothic',9)).place(x=34,y=58)
        tk.Label(head,text='PACK DEVELOPMENT QUALITY',bg=ui.NAVY,fg='#D7E4EF',font=('Segoe UI',8,'bold')).place(relx=1,x=-32,y=38,anchor='e')

        body=tk.Frame(self,bg=ui.BG); body.pack(fill='both',expand=True,padx=26,pady=18)
        mode=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); mode.pack(fill='x',pady=(0,12))
        tk.Label(mode,text='이슈 처리 방식',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left',padx=(18,24),pady=13)
        for text,value in [('기존 이슈 업데이트','existing'),('신규 이슈 등록','new')]:
            ttk.Radiobutton(mode,text=text,variable=self.mode,value=value,style='Mode.TRadiobutton').pack(side='left',padx=(0,20))
        tk.Label(mode,text='기존 이슈는 동일 이슈 행/페이지를 찾아 업데이트합니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='right',padx=18)

        content=tk.Frame(body,bg=ui.BG); content.pack(fill='x')
        left=tk.Frame(content,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); left.pack(side='left',fill='both',expand=True,padx=(0,7))
        right=tk.Frame(content,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); right.pack(side='left',fill='both',expand=True,padx=(7,0))
        self._section_title(left,'01','입력 파일','Drag & Drop 지원')
        filebody=tk.Frame(left,bg='white'); filebody.pack(fill='x',padx=16,pady=(0,14))
        for key in ('ppt8d','pptweekly','xlsx'): self.vars[key]=tk.StringVar()
        self.dropzones={}
        specs=[('8D 원본 PPT','ppt8d',('.pptx',),True),('주간회의 PPT','pptweekly',('.pptx',),False),('Issue DB Excel','xlsx',('.xlsx',),False)]
        for title,key,exts,req in specs:
            z=ui.DropZone(filebody,title,key,self.vars[key],exts,self.pick,self._refresh_target,req); z.pack(fill='x',pady=5); self.dropzones[key]=z
        dnd='사용 가능' if ui.DND_AVAILABLE else '미사용 · [찾기] 버튼 사용'
        tk.Label(filebody,text=f'Drag & Drop: {dnd}  |  8D는 필수, 나머지는 선택 입력',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',pady=(4,0))

        self._section_title(right,'02','담당 및 분류 정보','표준 입력')
        fields=tk.Frame(right,bg='white'); fields.pack(fill='x',padx=18,pady=(0,10))
        for key in ('team','owner','sample','plm_no'): self.vars[key]=tk.StringVar()
        self._enterprise_entry(fields,'담당팀','team','예: Pack개발품질1팀')
        self._enterprise_entry(fields,'담당자','owner','예: 홍길동')
        self._enterprise_entry(fields,'발생 샘플','sample','예: DUT3 / Sample No.')
        self._enterprise_entry(fields,'PMS/PLM 이슈번호','plm_no','기존값이 있을 때 입력')
        tk.Frame(fields,bg=ui.BORDER,height=1).pack(fill='x',pady=10)
        self.vars['form_factor']=tk.StringVar(value='파우치형'); self.vars['product_type']=tk.StringVar(value='EV Pack'); self.vars['occurrence_site']=tk.StringVar(value=legacy.OCCURRENCE_SITES[0]); self.vars['stage']=tk.StringVar(value='DV')
        self._enterprise_combo(fields,'폼팩터','form_factor',legacy.FORM_FACTORS)
        self._enterprise_combo(fields,'제품 타입','product_type',legacy.PRODUCT_TYPES)
        self._enterprise_combo(fields,'발생처','occurrence_site',legacy.OCCURRENCE_SITES)
        self._enterprise_combo(fields,'개발 단계','stage',legacy.STAGES)

        target=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); target.pack(fill='x',pady=12)
        tk.Label(target,text='UPDATE TARGET',bg='white',fg=ui.MUTED,font=('Segoe UI',8,'bold')).pack(side='left',padx=(16,12),pady=12)
        self.target_label=tk.Label(target,text='주간회의 또는 Issue DB를 선택해 주세요.',bg='white',fg=ui.NAVY,font=('Malgun Gothic',9,'bold')); self.target_label.pack(side='left')
        self.target_weekly=tk.Label(target,text='WEEKLY  OFF',bg='#EDF1F4',fg=ui.MUTED,font=('Segoe UI',8,'bold'),padx=9,pady=4); self.target_weekly.pack(side='right',padx=(5,16))
        self.target_excel=tk.Label(target,text='ISSUE DB  OFF',bg='#EDF1F4',fg=ui.MUTED,font=('Segoe UI',8,'bold'),padx=9,pady=4); self.target_excel.pack(side='right')

        actions=tk.Frame(body,bg=ui.BG); actions.pack(fill='x',pady=(0,10))
        tk.Button(actions,text='8D 내용 미리보기',command=self.preview,bg='#E5EBF0',fg=ui.NAVY,bd=0,font=('Malgun Gothic',9,'bold'),padx=18,pady=9,cursor='hand2').pack(side='left')
        tk.Button(actions,text='자동 업데이트 실행  →',command=self.run,bg=ui.BLUE,fg='white',activebackground='#12598F',activeforeground='white',bd=0,font=('Malgun Gothic',10,'bold'),padx=24,pady=10,cursor='hand2').pack(side='right')

        result=tk.Frame(body,bg='white',highlightbackground=ui.BORDER,highlightthickness=1); result.pack(fill='both',expand=True)
        rt=tk.Frame(result,bg='white'); rt.pack(fill='x',padx=16,pady=(12,7))
        tk.Label(rt,text='실행 결과',bg='white',fg=ui.NAVY,font=('Malgun Gothic',10,'bold')).pack(side='left')
        tk.Label(rt,textvariable=self.status_var,bg='#E9F1F8',fg='#315A7D',font=('Malgun Gothic',8),padx=9,pady=4).pack(side='right')
        self.log=tk.Text(result,height=7,wrap='word',font=('Consolas',9),bg='#FBFCFD',fg='#283A49',relief='flat',highlightthickness=1,highlightbackground='#E1E7EC',padx=10,pady=8)
        self.log.pack(fill='both',expand=True,padx=16,pady=(0,14)); self.log.insert('end','8D 원본을 Drag & Drop하거나 [찾기]로 선택해 주세요.\n')
        foot=tk.Frame(self,bg='#E8EEF3',height=30); foot.pack(fill='x'); foot.pack_propagate(False)
        tk.Label(foot,text='8D Issue Automation 1.0   ·   원본 파일은 직접 덮어쓰지 않습니다.',bg='#E8EEF3',fg='#607180',font=('Malgun Gothic',8)).pack(side='left',padx=26)
        self._refresh_target()

    def _section_title(self,parent,no,title,sub):
        f=tk.Frame(parent,bg='white'); f.pack(fill='x',padx=18,pady=(15,10))
        tk.Label(f,text=no,bg=ui.NAVY,fg='white',font=('Segoe UI',8,'bold'),width=3,pady=3).pack(side='left')
        tk.Label(f,text=title,bg='white',fg=ui.NAVY,font=('Malgun Gothic',11,'bold')).pack(side='left',padx=9)
        tk.Label(f,text=sub,bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(side='right')
    def _enterprise_entry(self,parent,label,key,hint):
        row=tk.Frame(parent,bg='white'); row.pack(fill='x',pady=4); tk.Label(row,text=label,bg='white',fg=ui.TEXT,font=('Malgun Gothic',9),width=17,anchor='w').pack(side='left')
        e=ttk.Entry(row,textvariable=self.vars[key]); e.pack(side='left',fill='x',expand=True); tk.Label(parent,text=hint,bg='white',fg='#98A3AD',font=('Malgun Gothic',7)).pack(anchor='e')
    def _enterprise_combo(self,parent,label,key,values):
        row=tk.Frame(parent,bg='white'); row.pack(fill='x',pady=5); tk.Label(row,text=label,bg='white',fg=ui.TEXT,font=('Malgun Gothic',9),width=17,anchor='w').pack(side='left'); ttk.Combobox(row,textvariable=self.vars[key],values=values,state='readonly').pack(side='left',fill='x',expand=True)

    def pick(self,k,desc=None,pattern=None):
        is_excel=(k=='xlsx'); types=[('Excel (*.xlsx)','*.xlsx')] if is_excel else [('PowerPoint (*.pptx)','*.pptx')]
        f=filedialog.askopenfilename(title='파일 선택',filetypes=types+[('모든 파일','*.*')],parent=self)
        if f: self.vars[k].set(f); self._refresh_target()

    def _refresh_target(self):
        if not hasattr(self,'target_label'): return
        weekly=bool(self.vars.get('pptweekly') and self.vars['pptweekly'].get().strip()); excel=bool(self.vars.get('xlsx') and self.vars['xlsx'].get().strip())
        if weekly and excel: text='주간회의 PPT + Issue DB Excel'
        elif weekly: text='주간회의 PPT만 업데이트'
        elif excel: text='Issue DB Excel만 업데이트'
        else: text='주간회의 또는 Issue DB를 선택해 주세요.'
        self.target_label.configure(text=text)
        self.target_weekly.configure(text='WEEKLY  ON' if weekly else 'WEEKLY  OFF',bg='#E7F4ED' if weekly else '#EDF1F4',fg=ui.GREEN if weekly else ui.MUTED)
        self.target_excel.configure(text='ISSUE DB  ON' if excel else 'ISSUE DB  OFF',bg='#E7F4ED' if excel else '#EDF1F4',fg=ui.GREEN if excel else ui.MUTED)


if __name__=='__main__':
    EnterpriseApp().mainloop()
