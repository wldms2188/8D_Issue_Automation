import os
import tkinter as tk
from pathlib import Path
from tkinter import ttk

NAVY='#173A5E'; BLUE='#1769AA'; BG='#F3F6F9'; WHITE='#FFFFFF'; TEXT='#263746'; MUTED='#70808F'; BORDER='#D9E3EC'; PALE='#F5F8FB'; GREEN='#2F7A55'; RED='#B94B4B'

try:
    from tkinterdnd2 import DND_FILES
    DND_AVAILABLE=True
except Exception:
    DND_FILES=None; DND_AVAILABLE=False


def center_window(win, parent=None, width=500, height=250):
    win.update_idletasks()
    if parent is not None and parent.winfo_exists():
        parent.update_idletasks()
        x=parent.winfo_rootx()+max(0,(parent.winfo_width()-width)//2)
        y=parent.winfo_rooty()+max(0,(parent.winfo_height()-height)//2)
    else:
        sw=win.winfo_screenwidth(); sh=win.winfo_screenheight()
        x=(sw-width)//2; y=(sh-height)//2
    win.geometry(f'{width}x{height}+{x}+{y}')


def _dialog_height(message, requested=250):
    """Keep the footer/buttons visible even for multiline confirmation text."""
    lines=0
    for raw in str(message).splitlines() or ['']:
        lines += max(1, (len(raw)+43)//44)
    return max(requested, min(640, 195 + lines*24))


def dialog(parent, title, message, kind='info', buttons=(('확인', True),), width=500, height=250, button_width=11):
    height=_dialog_height(message,height)
    win=tk.Toplevel(parent); win.withdraw(); win.title(title); win.configure(bg=WHITE); win.resizable(False,False); win.transient(parent)
    result={'value':None}
    head=tk.Frame(win,bg=NAVY,height=56); head.pack(fill='x',side='top'); head.pack_propagate(False)
    tk.Label(head,text=title,bg=NAVY,fg=WHITE,font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
    # Taller footer + button padding: visually matches the primary action buttons in the main UI.
    foot=tk.Frame(win,bg='#F6F8FA',height=76); foot.pack(fill='x',side='bottom'); foot.pack_propagate(False)
    box=tk.Frame(foot,bg='#F6F8FA'); box.pack(side='right',padx=20,pady=14)
    def choose(v): result['value']=v; win.destroy()
    for i,(label,value) in enumerate(buttons):
        primary=(i==len(buttons)-1)
        b=tk.Button(box,text=label,command=lambda v=value:choose(v),font=('Malgun Gothic',9,'bold'),width=button_width,bd=0,cursor='hand2',bg=BLUE if primary else '#E5EBF0',fg=WHITE if primary else TEXT,activebackground='#12598F' if primary else '#D9E2E9',activeforeground=WHITE if primary else TEXT,padx=4,pady=10)
        b.pack(side='left',padx=(8,0))
    icon={'info':'i','warning':'!','error':'×','question':'?'}.get(kind,'i')
    col={'info':BLUE,'warning':'#C98424','error':RED,'question':BLUE}.get(kind,BLUE)
    body=tk.Frame(win,bg=WHITE); body.pack(fill='both',expand=True,padx=24,pady=18)
    tk.Label(body,text=icon,bg=col,fg=WHITE,font=('Malgun Gothic',14,'bold'),width=2,height=1).pack(side='left',anchor='n',padx=(0,15))
    tk.Label(body,text=message,bg=WHITE,fg=TEXT,font=('Malgun Gothic',10),justify='left',anchor='nw',wraplength=max(300,width-110)).pack(side='left',fill='both',expand=True)
    win.protocol('WM_DELETE_WINDOW',lambda:choose(None)); center_window(win,parent,width,height); win.deiconify(); win.grab_set(); win.focus_force(); parent.wait_window(win)
    return result['value']


def info(parent,title,message): return dialog(parent,title,message,'info')
def warning(parent,title,message): return dialog(parent,title,message,'warning')
def error(parent,title,message): return dialog(parent,title,message,'error')
def ask_yes_no(parent,title,message): return dialog(parent,title,message,'question',(('아니오',False),('확인',True))) is True


def parse_drop_files(widget, data):
    try: return list(widget.tk.splitlist(data))
    except Exception: return [data.strip('{}')]


def enable_drop(widget, callback):
    if not DND_AVAILABLE: return False
    try:
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind('<<Drop>>',lambda e: callback(parse_drop_files(widget,e.data)))
        return True
    except Exception: return False


class DropZone(tk.Frame):
    def __init__(self,parent,title,key,var,extensions,on_browse,on_change,required=False):
        super().__init__(parent,bg=WHITE,highlightbackground=BORDER,highlightthickness=1,cursor='hand2')
        self.key=key; self.var=var; self.extensions=tuple(x.lower() for x in extensions); self.on_change=on_change
        top=tk.Frame(self,bg=WHITE); top.pack(fill='x',padx=12,pady=(9,2))
        tk.Label(top,text=title,bg=WHITE,fg=TEXT,font=('Malgun Gothic',9,'bold')).pack(side='left')
        tk.Label(top,text='필수' if required else '선택',bg='#E7F4ED' if required else '#E8F1F8',fg=GREEN if required else '#3E6C8E',font=('Malgun Gothic',7,'bold'),padx=6,pady=1).pack(side='right')
        self.name=tk.Label(self,text='파일을 끌어다 놓거나 클릭하여 선택',bg=PALE,fg=MUTED,font=('Malgun Gothic',8),anchor='w',padx=10,pady=8)
        self.name.pack(fill='x',padx=12,pady=(4,5))
        bottom=tk.Frame(self,bg=WHITE); bottom.pack(fill='x',padx=12,pady=(0,9))
        self.state=tk.Label(bottom,text='미선택',bg=WHITE,fg=MUTED,font=('Malgun Gothic',8)); self.state.pack(side='left')
        tk.Button(bottom,text='지우기',command=self.clear,bg=WHITE,fg=MUTED,bd=0,font=('Malgun Gothic',8),cursor='hand2').pack(side='right')
        tk.Button(bottom,text='찾기',command=lambda:on_browse(key),bg='#E8EFF5',fg=NAVY,bd=0,font=('Malgun Gothic',8,'bold'),padx=10,pady=3,cursor='hand2').pack(side='right',padx=(0,5))
        self.name.bind('<Button-1>',lambda e:on_browse(key)); self.bind('<Button-1>',lambda e:on_browse(key))
        enable_drop(self,self._drop); enable_drop(self.name,self._drop)
        var.trace_add('write',lambda *_:self._var_changed()); self.refresh()
    def _var_changed(self):
        self.refresh()
        try:self.on_change()
        except Exception:pass
    def _drop(self,paths):
        if not paths:return
        path=paths[0]
        if Path(path).suffix.lower() not in self.extensions:
            warning(self.winfo_toplevel(),'파일 형식 확인',f"허용되는 파일 형식: {', '.join(self.extensions)}")
            return
        self.var.set(os.path.normpath(path))
    def clear(self): self.var.set('')
    def refresh(self):
        p=self.var.get().strip()
        if p:
            self.name.configure(text=Path(p).name,fg=TEXT,bg='#F1F7FB'); self.state.configure(text='✓ 선택 완료',fg=GREEN)
        else:
            self.name.configure(text='파일을 끌어다 놓거나 클릭하여 선택',fg=MUTED,bg=PALE); self.state.configure(text='미선택',fg=MUTED)
