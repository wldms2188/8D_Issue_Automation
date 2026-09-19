"""Enterprise UI v3.
Windows/Tk production layout:
- fixed 50/50 and 65/35 cards independent of long filenames;
- always-visible progress strip;
- explicit active-card border state for click, browse, focus, wheel and drag/drop interaction;
- exactly one visible progress percentage.
Business logic remains inherited unchanged from main_enterprise_v2.
"""
import re
import tkinter as tk

import main_enterprise_v2 as v2
import ui_enterprise as ui


class EnterpriseAppV3(v2.EnterpriseAppV2):
    ACTIVE_BORDER = ui.NAVY
    INACTIVE_BORDER = ui.BORDER

    @staticmethod
    def _initial_window_size(sw,sh):
        # Use most of the available desktop so the result/status area is visible
        # at common Windows DPI settings, while keeping a small taskbar margin.
        w=max(1000,min(1380,int(sw)-40))
        h=max(720,min(980,int(sh)-45))
        return min(w,int(sw)),min(h,int(sh))

    def _compact_layout(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,h=self._initial_window_size(sw,sh)
        self.minsize(min(1080,w),min(740,h))
        self.geometry(f'{w}x{h}+{max(0,(sw-w)//2)}+{max(0,(sh-h)//2)}')
        # Production UI should open maximized on Windows, matching the validated enterprise layout.
        # Keep the calculated geometry as a safe fallback on environments that do not support zoomed state.
        try:
            self.state('zoomed')
        except Exception:
            pass
        try: self.log.configure(height=5)
        except Exception: pass
        for z in getattr(self, 'dropzones', {}).values():
            try: z.pack_configure(pady=2)
            except Exception: pass

    def _balance_and_focus_cards(self):
        self.card01 = self._find_section_card('01')
        self.card02 = self._find_section_card('02')
        if not self.card01 or not self.card02 or self.card01.master is not self.card02.master: return
        p = self.card01.master; self._cards_parent = p
        self.update_idletasks()
        natural_h = max(p.winfo_height(), self.card01.winfo_reqheight(), self.card02.winfo_reqheight(), 410)
        try: self.card01.pack_forget(); self.card02.pack_forget()
        except Exception: pass
        try: self.card01.grid_forget(); self.card02.grid_forget()
        except Exception: pass
        p.configure(height=natural_h); p.pack_propagate(False); p.grid_propagate(False)

        def bind_tree(widget, side):
            widget.bind('<ButtonPress-1>', lambda _e,s=side:self._set_card_ratio(s), add='+')
            widget.bind('<FocusIn>', lambda _e,s=side:self._set_card_ratio(s), add='+')
            widget.bind('<MouseWheel>', lambda _e,s=side:self._set_card_ratio(s), add='+')
            widget.bind('<Button-4>', lambda _e,s=side:self._set_card_ratio(s), add='+')
            widget.bind('<Button-5>', lambda _e,s=side:self._set_card_ratio(s), add='+')
            for child in widget.winfo_children(): bind_tree(child, side)
        bind_tree(self.card01,1); bind_tree(self.card02,2)
        self.bind_all('<ButtonPress-1>',self._route_card_click_v3,add='+')
        self.bind_all('<FocusIn>',self._route_card_focus_v3,add='+')
        self.bind_all('<MouseWheel>',self._route_card_wheel_v3,add='+')

        # Real tkinterdnd2 callbacks are registered at DropZone construction time.
        # Bind directly to the Tcl <<Drop>> event AND to the selected-file variable.
        # Either path activates input card 01; this is deliberately redundant for Windows DnD.
        for zone in getattr(self,'dropzones',{}).values():
            if ui.DND_AVAILABLE:
                for target in (zone, getattr(zone,'name',None)):
                    if target is None: continue
                    try: target.dnd_bind('<<Drop>>', lambda _e:self._set_card_ratio(1), add='+')
                    except Exception: pass
            try: zone.var.trace_add('write', lambda *_:self._set_card_ratio(1))
            except Exception: pass
        self._set_card_ratio(0)

    def _route_card_click_v3(self,event):
        if self._is_descendant(event.widget,getattr(self,'card02',None)): self._set_card_ratio(2)
        elif self._is_descendant(event.widget,getattr(self,'card01',None)): self._set_card_ratio(1)
    def _route_card_focus_v3(self,event):
        if self._is_descendant(event.widget,getattr(self,'card02',None)): self._set_card_ratio(2)
        elif self._is_descendant(event.widget,getattr(self,'card01',None)): self._set_card_ratio(1)
    def _route_card_wheel_v3(self,event):
        if self._is_descendant(event.widget,getattr(self,'card02',None)): self._set_card_ratio(2)
        elif self._is_descendant(event.widget,getattr(self,'card01',None)): self._set_card_ratio(1)

    def _paint_active_card(self,active):
        try:
            if active==1:
                self.card01.configure(highlightbackground=self.ACTIVE_BORDER,highlightcolor=self.ACTIVE_BORDER,highlightthickness=2)
                self.card02.configure(highlightbackground=self.INACTIVE_BORDER,highlightcolor=self.INACTIVE_BORDER,highlightthickness=1)
            elif active==2:
                self.card01.configure(highlightbackground=self.INACTIVE_BORDER,highlightcolor=self.INACTIVE_BORDER,highlightthickness=1)
                self.card02.configure(highlightbackground=self.ACTIVE_BORDER,highlightcolor=self.ACTIVE_BORDER,highlightthickness=2)
            else:
                for card in (self.card01,self.card02): card.configure(highlightbackground=self.INACTIVE_BORDER,highlightcolor=self.INACTIVE_BORDER,highlightthickness=1)
        except Exception: pass

    def _set_card_ratio(self,active=0):
        if not getattr(self,'card01',None) or not getattr(self,'card02',None): return
        try:
            left,right=(0.65,0.35) if active==1 else ((0.35,0.65) if active==2 else (0.50,0.50))
            self.card01.place(x=0,rely=0,relwidth=left,relheight=1.0)
            self.card02.place(relx=left,x=7,rely=0,relwidth=right,width=-7,relheight=1.0)
            self._paint_active_card(active); self.card01.lift(); self.card02.lift(); self.update_idletasks()
        except Exception: pass

    def _find_result_header(self):
        found=[]
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text')=='실행 결과': found.append(child.master)
                except Exception: pass
                walk(child)
        walk(self); return found[0] if found else None

    @staticmethod
    def _clean_status_for_display(text):
        """Never show a second numeric progress value next to the dedicated percent label."""
        s=str(text or '').strip()
        # Handles '85 RUNNING', '85% RUNNING', '85% · RUNNING', and repeated variants.
        s=re.sub(r'^\s*(?:\d{1,3}\s*%?\s*(?:[·|:\-]\s*)?)+(?=[A-Za-z가-힣])', '', s).strip()
        return s or 'READY'

    def _install_progress_panel(self):
        """One progress bar + one percentage + a separately sanitized status label."""
        self.progress_value=tk.DoubleVar(value=0)
        header=self._find_result_header()
        if header is None:return

        # The original Enterprise UI already owns the status label. Point it at a display-only
        # variable so business status_var can never leak a duplicate numeric percentage onscreen.
        self.status_display_var=tk.StringVar(value=self._clean_status_for_display(self.status_var.get()))
        def retarget_status_label(widget):
            for child in widget.winfo_children():
                try:
                    if str(child.cget('textvariable')) == str(self.status_var):
                        child.configure(textvariable=self.status_display_var)
                except Exception: pass
                retarget_status_label(child)
        retarget_status_label(self)

        progress_wrap=tk.Frame(header,bg='white'); progress_wrap.pack(side='right',fill='x',expand=True,padx=(20,10))
        self.progress_percent=tk.Label(progress_wrap,text='0%',bg='white',fg=ui.NAVY,font=('Malgun Gothic',8,'bold'),width=5,anchor='e')
        self.progress_percent.pack(side='right',padx=(8,0))
        self.progress_canvas=tk.Canvas(progress_wrap,height=12,bg='#D9E2E9',highlightthickness=1,highlightbackground='#B9C8D3',bd=0)
        self.progress_canvas.pack(side='right',fill='x',expand=True,pady=3)
        self.progress_fill=self.progress_canvas.create_rectangle(0,0,0,12,fill='#1769AA',outline='')
        self.progress_canvas.bind('<Configure>',lambda _e:self._paint_progress_bar())
        self.status_var.trace_add('write',self._sync_progress_from_status)
        self.status_var.trace_add('write',self._sync_status_display)
        self._sync_progress_from_status(); self._sync_status_display()

    def _sync_status_display(self,*_):
        try: self.status_display_var.set(self._clean_status_for_display(self.status_var.get()))
        except Exception: pass

    def _paint_progress_bar(self):
        try:
            width=max(2,self.progress_canvas.winfo_width()); height=max(2,self.progress_canvas.winfo_height())
            pct=max(0.0,min(100.0,float(self.progress_value.get())))
            self.progress_canvas.coords(self.progress_fill,0,0,width*pct/100.0,height); self.progress_canvas.tag_raise(self.progress_fill)
            self.progress_percent.configure(text=f'{int(pct)}%')
        except Exception: pass

    def _set_progress(self,pct,label):
        """Update only the dedicated percentage widget; never add pct to status text."""
        pct=max(0,min(100,int(pct))); self.progress_value.set(pct); self._paint_progress_bar()
        try: self.update_idletasks(); self.update()
        except Exception: pass


if __name__=='__main__':
    EnterpriseAppV3().mainloop()
