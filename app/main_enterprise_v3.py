"""Enterprise UI v3.
Fixes two Windows/Tk layout issues seen with populated file cards:
1) requested child widths preventing the opposite card from expanding;
2) progress strip being clipped below the visible result area.
Business logic remains inherited unchanged from main_enterprise_v2.
"""
import tkinter as tk

import main_enterprise_v2 as v2
import ui_enterprise as ui


class EnterpriseAppV3(v2.EnterpriseAppV2):
    def _compact_layout(self):
        """Keep enough vertical room at common Windows display scaling."""
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(1180, max(1040, sw - 60))
        # Do not request a window taller than the usable screen area.
        h = min(820, max(700, sh - 55))
        self.minsize(min(1040, w), min(690, h))
        self.geometry(f'{w}x{h}+{max(0,(sw-w)//2)}+{max(0,(sh-h)//2)}')
        try:
            self.log.configure(height=2)
        except Exception:
            pass
        for z in getattr(self, 'dropzones', {}).values():
            try:
                z.pack_configure(pady=2)
            except Exception:
                pass

    def _balance_and_focus_cards(self):
        """Use place geometry so child requested widths can never block 50/50 or 35/65 sizing."""
        self.card01 = self._find_section_card('01')
        self.card02 = self._find_section_card('02')
        if not self.card01 or not self.card02 or self.card01.master is not self.card02.master:
            return
        p = self.card01.master
        self._cards_parent = p

        # Capture the natural height while the original packed cards are still rendered.
        self.update_idletasks()
        natural_h = max(p.winfo_height(), self.card01.winfo_reqheight(), self.card02.winfo_reqheight())
        natural_h = max(410, natural_h)

        try:
            self.card01.pack_forget()
            self.card02.pack_forget()
        except Exception:
            pass
        try:
            self.card01.grid_forget()
            self.card02.grid_forget()
        except Exception:
            pass

        # Critical: freeze parent height and use place. Unlike grid/pack, place ignores
        # the long filename / Entry requested width that was keeping card 01 too wide.
        p.configure(height=natural_h)
        p.pack_propagate(False)
        p.grid_propagate(False)

        def bind_tree(widget, side):
            widget.bind('<ButtonPress-1>', lambda _e, s=side: self._set_card_ratio(s), add='+')
            widget.bind('<FocusIn>', lambda _e, s=side: self._set_card_ratio(s), add='+')
            for child in widget.winfo_children():
                bind_tree(child, side)

        bind_tree(self.card01, 1)
        bind_tree(self.card02, 2)
        self.bind_all('<ButtonPress-1>', self._route_card_click_v3, add='+')
        self.bind_all('<FocusIn>', self._route_card_focus_v3, add='+')
        self._set_card_ratio(0)

    def _route_card_click_v3(self, event):
        if self._is_descendant(event.widget, getattr(self, 'card02', None)):
            self._set_card_ratio(2)
        elif self._is_descendant(event.widget, getattr(self, 'card01', None)):
            self._set_card_ratio(1)

    def _route_card_focus_v3(self, event):
        if self._is_descendant(event.widget, getattr(self, 'card02', None)):
            self._set_card_ratio(2)
        elif self._is_descendant(event.widget, getattr(self, 'card01', None)):
            self._set_card_ratio(1)

    def _set_card_ratio(self, active=0):
        """Default 50/50; selected card 65%, other card 35%, regardless of its contents."""
        if not getattr(self, 'card01', None) or not getattr(self, 'card02', None):
            return
        try:
            if active == 1:
                left, right = 0.65, 0.35
            elif active == 2:
                left, right = 0.35, 0.65
            else:
                left, right = 0.50, 0.50

            # Small absolute gap between cards; widths are forced by relwidth.
            self.card01.place(x=0, rely=0, relwidth=left, relheight=1.0)
            self.card02.place(relx=left, x=7, rely=0, relwidth=right, width=-7, relheight=1.0)
            self.card01.lift()
            self.card02.lift()
            self.update_idletasks()
        except Exception:
            pass

    def _find_result_header(self):
        found = []
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text') == '실행 결과':
                        found.append(child.master)
                except Exception:
                    pass
                walk(child)
        walk(self)
        return found[0] if found else None

    def _install_progress_panel(self):
        """Put progress in the already-visible '실행 결과' header, never below it."""
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_text = tk.StringVar(value='0% · 실행 대기')
        header = self._find_result_header()
        if header is None:
            return

        # Existing status badge is already on the right. Pack progress immediately beside it.
        progress_wrap = tk.Frame(header, bg='white')
        progress_wrap.pack(side='right', fill='x', expand=True, padx=(20, 10))

        self.progress_percent = tk.Label(
            progress_wrap, text='0%', bg='white', fg=ui.NAVY,
            font=('Malgun Gothic', 8, 'bold'), width=5, anchor='e'
        )
        self.progress_percent.pack(side='right', padx=(8, 0))

        self.progress_canvas = tk.Canvas(
            progress_wrap, height=12, bg='#D9E2E9',
            highlightthickness=1, highlightbackground='#B9C8D3', bd=0
        )
        self.progress_canvas.pack(side='right', fill='x', expand=True, pady=3)
        self.progress_fill = self.progress_canvas.create_rectangle(
            0, 0, 0, 12, fill='#1769AA', outline=''
        )
        self.progress_canvas.bind('<Configure>', lambda _e: self._paint_progress_bar())
        self.status_var.trace_add('write', self._sync_progress_from_status)
        self._sync_progress_from_status()

    def _paint_progress_bar(self):
        try:
            width = max(2, self.progress_canvas.winfo_width())
            height = max(2, self.progress_canvas.winfo_height())
            pct = max(0.0, min(100.0, float(self.progress_value.get())))
            self.progress_canvas.coords(self.progress_fill, 0, 0, width * pct / 100.0, height)
            self.progress_canvas.tag_raise(self.progress_fill)
            self.progress_percent.configure(text=f'{int(pct)}%')
        except Exception:
            pass

    def _set_progress(self, pct, label):
        pct = max(0, min(100, int(pct)))
        self.progress_value.set(pct)
        self.progress_text.set(f'{pct}% · {label}')
        self._paint_progress_bar()
        try:
            self.update_idletasks()
            self.update()
        except Exception:
            pass


if __name__ == '__main__':
    EnterpriseAppV3().mainloop()
