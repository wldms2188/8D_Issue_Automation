"""Skip duplicate 8D attachment pages before native PowerPoint COM insertion."""
import hashlib, os, shutil, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import main_recovery_step14_fix2 as core
_original_append=core._append_8d_attachments

def _shape_parts(sh):
    parts=[str(getattr(sh,'shape_type','')),str(getattr(sh,'left','')),str(getattr(sh,'top','')),str(getattr(sh,'width','')),str(getattr(sh,'height',''))]
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
        for child in sh.shapes:parts.extend(_shape_parts(child))
    if getattr(sh,'has_table',False):
        tb=sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):parts.append('T:'+str(tb.cell(r,c).text or '').strip())
    elif hasattr(sh,'text'):parts.append('X:'+str(getattr(sh,'text','') or '').strip())
    if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:
        try:parts.append('I:'+hashlib.sha1(sh.image.blob).hexdigest())
        except Exception:pass
    return parts

def _slide_signature(sl):
    parts=[]
    for sh in sl.shapes:parts.extend(_shape_parts(sh))
    return hashlib.sha1('\x1f'.join(parts).encode('utf-8','ignore')).hexdigest()

def _delete_slide(prs,index):
    slide_id=prs.slides._sldIdLst[index]; prs.part.drop_rel(slide_id.rId); prs.slides._sldIdLst.remove(slide_id)

def _remove_current_auto(out_path,prefix):
    prs=Presentation(out_path); changed=False
    for i in range(len(prs.slides)-1,-1,-1):
        if str(getattr(prs.slides[i],'name','') or '').startswith(prefix):_delete_slide(prs,i); changed=True
    if changed:prs.save(out_path)

def _append_without_duplicates(out_path,src8d_path,detail_index,d):
    src=Presentation(src8d_path)
    if len(src.slides)<=1:return 0
    dest=Presentation(out_path); current_prefix='AUTO_8D_ATTACH_'+core._attachment_key(d)+'_'; existing=set()
    for sl in dest.slides:
        if str(getattr(sl,'name','') or '').startswith(current_prefix):continue
        existing.add(_slide_signature(sl))
    keep=[]; seen=set(existing)
    for i in range(1,len(src.slides)):
        sig=_slide_signature(src.slides[i])
        if sig in seen:continue
        seen.add(sig); keep.append(i)
    if len(keep)==len(src.slides)-1:return _original_append(out_path,src8d_path,detail_index,d)
    if not keep:
        _remove_current_auto(out_path,current_prefix)
        return 0
    fd,tmp=tempfile.mkstemp(prefix='8d_attach_',suffix='.pptx'); os.close(fd)
    try:
        shutil.copy2(src8d_path,tmp); reduced=Presentation(tmp); keep_set={0,*keep}
        for idx in range(len(reduced.slides)-1,-1,-1):
            if idx not in keep_set:_delete_slide(reduced,idx)
        reduced.save(tmp); return _original_append(out_path,tmp,detail_index,d)
    finally:
        try:Path(tmp).unlink(missing_ok=True)
        except Exception:pass
core._append_8d_attachments=_append_without_duplicates
