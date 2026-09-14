# Recovery STEP6: requested GUI dropdowns only.
# Preserve STEP5 DB behavior and STEP4 weekly PPT/image/layout behavior.

import tkinter as tk
from tkinter import ttk
import main_recovery_step5 as step5

base = step5.base

FORM_FACTORS = ('원통형','파우치형')
PRODUCT_TYPES = ('ESS BPU','ESS Link','EV Cell-Unit','EV Module','EV Pack','IT Pack','LEV Pack','PHEV Pack')
OCCURRENCE_SITES = ('etc.','고객/상위 Claim','고객/상위 생산','고객/상위 시험','고객/상위 운송/보관','부품 생산','제품 생산','제품 시험','제품 운송/보관')
STAGES = ('CV','DV','PD','MP')


class RecoveryStep6App(step5.RecoveryStep5App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP6')
        # Convert only the four existing entry widgets to readonly dropdowns.
        # Keep the inherited GUI structure/run logic untouched.
        self._convert_existing_field('form_factor', FORM_FACTORS)
        self._convert_existing_field('product_type', PRODUCT_TYPES)
        self._convert_existing_field('occurrence_site', OCCURRENCE_SITES)
        self._convert_existing_field('stage', STAGES)

    def _convert_existing_field(self,key,values):
        var=getattr(self,'vars',{}).get(key)
        if var is None:
            return False
        # Find Entry bound to this exact StringVar and replace it in-place.
        def walk(w):
            for ch in w.winfo_children():
                yield ch
                yield from walk(ch)
        for widget in list(walk(self)):
            if not isinstance(widget,ttk.Entry):
                continue
            try:
                if str(widget.cget('textvariable')) != str(var):
                    continue
            except Exception:
                continue
            parent=widget.master
            info=widget.pack_info()
            width=int(widget.cget('width') or 60)
            widget.destroy()
            cb=ttk.Combobox(parent,textvariable=var,values=values,state='readonly',width=width)
            # Reuse the original pack geometry as closely as possible.
            opts={k:v for k,v in info.items() if k not in ('in','in_')}
            cb.pack(**opts)
            if not var.get() and values:
                var.set(values[0])
            return True
        return False


if __name__=='__main__':
    RecoveryStep6App().mainloop()
