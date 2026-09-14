# 8D Issue Automation v3.0.3
# Activates the collision-aware flow renderer for the weekly page-2 renderer.
import sys
from pathlib import Path
APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))
import main_v301 as v301
base = v301.base
renderer = v301.impl  # main_v300; weekly_v300 resolves _fill_page2 in this module

FLOW = {"left":["2D","3D","4D_CAUSE"], "right":["4D_LEAK","5D","6D"]}
MIN_H = {"2D":0.62,"3D":0.70,"4D_CAUSE":0.90,"4D_LEAK":0.58,"5D":0.82,"6D":0.60}
BOTTOM = {"left":7.17,"right":6.85}
IMAGE_KEY = {"2D":"2D","3D":"3D","4D_CAUSE":"4D","4D_LEAK":"4D","5D":"5D","6D":"6D"}

def _text_for(k,d):
    if k=="2D": return str(d.get("problem") or "검토 중")
    if k=="3D": return "\n".join(x for x in (d.get("temporary_action"),d.get("customer_response")) if str(x or "").strip()) or "검토 중"
    if k=="4D_CAUSE": return str(d.get("cause_4d") or "검토 중")
    if k=="4D_LEAK": return "\n".join(str(x) for x in (d.get("leak_cause"),d.get("system_cause")) if str(x or "").strip()) or "검토 중"
    if k=="5D": return str(d.get("action_5d") or "검토 중")
    if k=="6D": return str(d.get("verification_6d") or "검토 중")
    return ""

def _need(k,text,width,h,has_img):
    tw=max(.65,width*(.64 if has_img else .94))
    c=max(7,int(tw*9.2))
    lines=sum(max(1,(len(s)+c-1)//c) for s in str(text).splitlines() or [""])
    return max(h,.26+lines*.145)

def _flow(layout,d):
    out={k:dict(v) for k,v in layout.items()}
    imgs=d.get("_section_images",{}) or {}
    for col,keys in FLOW.items():
        cursor=min(out[k]["y"] for k in keys)
        bottom=BOTTOM[col]
        desired={k:_need(k,_text_for(k,d),out[k]["w"],out[k]["h"],bool(imgs.get(IMAGE_KEY[k],[]))) for k in keys}
        for i,k in enumerate(keys):
            out[k]["y"]=cursor
            room=bottom-cursor-sum(MIN_H[z] for z in keys[i+1:])
            out[k]["h"]=max(MIN_H[k],min(desired[k],max(MIN_H[k],room)))
            cursor=out[k]["y"]+out[k]["h"]
    return out

_original_fill=renderer._fill_page2

def _collision_fill(sl,d,g,layout):
    _original_fill(sl,d,g,_flow(layout,d))

renderer._fill_page2=_collision_fill

if __name__=="__main__":
    base.App().mainloop()
