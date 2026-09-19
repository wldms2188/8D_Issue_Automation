"""Final output type selector: full updated files or update-only reduced files."""
import copy
import re
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation

import main_enterprise as ent
import ui_enterprise as ui

_original_info = ui.info


def _latest_version(source_path):
    src=Path(source_path)
    out=src.parent/'자동화_결과'
    root=src.stem
    pat=re.compile(rf'^{re.escape(root)}_v0[.](\d+){re.escape(src.suffix)}$',re.I)
    found=[]
    if out.exists():
        for p in out.iterdir():
            m=pat.match(p.name)
            if m:
                found.append((int(m.group(1)),p))
    return max(found,key=lambda x:x[0])[1] if found else None


def _slide_signature(slide):
    """Compare slide content while ignoring package-level metadata."""
    try:return str(slide._element.xml)
    except Exception:return ''


def _remove_slide(prs,index):
    slide_id=prs.slides._sldIdLst[index]
    rel_id=slide_id.rId
    prs.part.drop_rel(rel_id)
    del prs.slides._sldIdLst[index]


def _ppt_update_only(source,saved):
    """Keep only genuinely new/changed slides, regardless of slide reordering.

    Comparing source/destination by index makes every old detail page look changed
    when a new summary/detail page is inserted before it. Match slide signatures as
    a multiset instead, so unchanged pages are discarded even if their index moved.
    """
    src=Presentation(source); dst=Presentation(saved)
    source_counts=Counter(_slide_signature(sl) for sl in src.slides)
    keep=[]
    for i,slide in enumerate(dst.slides):
        sig=_slide_signature(slide)
        if source_counts.get(sig,0)>0:
            source_counts[sig]-=1
        else:
            keep.append(i)
    if not keep:
        return None,0
    keep_set=set(keep)
    for i in range(len(dst.slides)-1,-1,-1):
        if i not in keep_set:_remove_slide(dst,i)
    target=Path(saved).with_name(Path(saved).stem+'_업데이트사항만'+Path(saved).suffix)
    dst.save(target)
    return target,len(keep)


def _row_signature(ws,r,max_col):
    return tuple(str(ws.cell(r,c).value or '') for c in range(1,max_col+1))

def _anchor_row(obj):
    """Return a 1-based worksheet row for an image/chart anchor when possible."""
    anchor=getattr(obj,'anchor',None)
    if isinstance(anchor,str):
        m=re.fullmatch(r'[A-Za-z]+(\d+)',anchor.strip())
        return int(m.group(1)) if m else None
    try:
        return int(anchor._from.row)+1
    except Exception:
        return None

def _move_anchor_to_row(obj,new_row):
    anchor=getattr(obj,'anchor',None)
    old=_anchor_row(obj)
    if old is None:
        return
    delta=int(new_row)-int(old)
    if isinstance(anchor,str):
        m=re.fullmatch(r'([A-Za-z]+)\d+',anchor.strip())
        if m:
            obj.anchor=f'{m.group(1)}{int(new_row)}'
        return
    try:
        anchor._from.row=max(0,int(anchor._from.row)+delta)
    except Exception:
        pass
    try:
        anchor.to.row=max(0,int(anchor.to.row)+delta)
    except Exception:
        pass

def _prune_data_drawings(ws,changed_rows):
    """Keep header drawings and only the updated row's representative drawing(s).

    For representative images, when an update left both the old and new image on
    the same row, the most recently-added image wins.
    """
    row_map={old:7+i for i,old in enumerate(sorted(changed_rows))}
    for attr in ('_images','_charts'):
        items=list(getattr(ws,attr,[]) or [])
        header=[]; by_row={}
        for obj in items:
            row=_anchor_row(obj)
            if row is None or row<=6:
                header.append(obj)
                continue
            if row in row_map:
                by_row.setdefault(row,[]).append(obj)

        kept=list(header)
        for row in sorted(by_row):
            objs=by_row[row]
            if attr=='_images' and objs:
                # Issue DB has one representative image per data row. Keep the
                # newest one if both old/new images survived the full update.
                objs=[objs[-1]]
            for obj in objs:
                _move_anchor_to_row(obj,row_map[row])
                kept.append(obj)
        try:
            setattr(ws,attr,kept)
        except Exception:
            pass

def _detect_changed_rows(ws0,ws):
    max_col=max(ws0.max_column,ws.max_column)
    source_counts=Counter(_row_signature(ws0,r,max_col) for r in range(7,ws0.max_row+1))
    changed=[]
    for r in range(7,ws.max_row+1):
        sig=_row_signature(ws,r,max_col)
        if source_counts.get(sig,0)>0:
            source_counts[sig]-=1
        else:
            changed.append(r)
    return changed

def _focus_reduced_excel_view(wb,ws,updated_row=7):
    """Open the reduced Issue DB at the top with the updated row immediately visible."""
    try:
        wb.active=ws
    except Exception:
        try:wb.active=wb.index(ws)
        except Exception:pass

    # Keep the workbook scrolled to the top. If rows 1-6 are frozen, A7 is the
    # first scrollable cell and places the updated row directly under the headers.
    top_cell='A1'
    try:
        fp=ws.freeze_panes
        fp_row=getattr(fp,'row',None)
        if fp_row is None and isinstance(fp,str):
            m=re.search(r'(\d+)$',fp)
            fp_row=int(m.group(1)) if m else None
        if fp_row and int(fp_row)>=7:
            top_cell='A7'
    except Exception:
        pass

    try:ws.sheet_view.topLeftCell=top_cell
    except Exception:pass

    # Select the surviving updated row so Excel opens with the relevant record in focus.
    target=f'A{max(7,int(updated_row or 7))}'
    try:
        sels=list(ws.sheet_view.selection or [])
        if sels:
            sels[0].activeCell=target
            sels[0].sqref=target
        else:
            from openpyxl.worksheet.views import Selection
            ws.sheet_view.selection=[Selection(activeCell=target,sqref=target)]
    except Exception:
        pass


def _excel_update_only(source,saved,preferred_row=None):
    """Create a true reduced Issue DB: rows 1-6 + only the updated data row(s).

    All other data rows and their floating representative images/charts are removed.
    The surviving updated row(s) are compacted to row 7 onward.
    """
    src=load_workbook(source,rich_text=True)
    dst=load_workbook(saved,rich_text=True)
    ws0=src['Sheet1'] if 'Sheet1' in src.sheetnames else src.active
    ws=dst['Sheet1'] if 'Sheet1' in dst.sheetnames else dst.active

    changed=_detect_changed_rows(ws0,ws)
    if preferred_row and 7<=int(preferred_row)<=ws.max_row:
        preferred_row=int(preferred_row)
        # Exact row from the just-finished update wins for existing-issue updates.
        if preferred_row not in changed:
            changed=[preferred_row]
        else:
            changed=[preferred_row]+[r for r in changed if r!=preferred_row]

    if not changed:
        return None,0

    changed=sorted(set(r for r in changed if 7<=r<=ws.max_row))
    _prune_data_drawings(ws,changed)

    # Remove every old data row except the row(s) that were just updated.
    # Deleting bottom-up naturally compacts the kept rows to 7, 8, ...
    keep=set(changed)
    for r in range(ws.max_row,6,-1):
        if r not in keep:
            ws.delete_rows(r,1)

    # Reset the saved workbook view so the updated row is immediately visible.
    _focus_reduced_excel_view(dst,ws,7)

    target=Path(saved).with_name(Path(saved).stem+'_업데이트사항만'+Path(saved).suffix)
    dst.save(target)
    return target,len(changed)


class OutputTypeDialog(ent.tk.Toplevel):
    def __init__(self,parent):
        super().__init__(parent); self.withdraw(); self.result=None
        self.title('업데이트 파일 유형 선택'); self.configure(bg='white'); self.resizable(False,False); self.transient(parent)
        head=ent.tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        ent.tk.Label(head,text='업데이트 파일 유형 선택',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
        body=ent.tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=20)
        ent.tk.Label(body,text='저장할 Output 파일 유형을 선택해 주세요.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold')).pack(anchor='w',pady=(0,14))
        self.var=ent.tk.StringVar(value='full')
        ent.ttk.Radiobutton(body,text='1. 전체 업데이트 본',variable=self.var,value='full',style='Mode.TRadiobutton').pack(anchor='w',pady=8)
        ent.tk.Label(body,text='기존 주간회의 전체 페이지 / Issue DB 전체 행을 유지합니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',padx=24)
        ent.ttk.Radiobutton(body,text='2. 업데이트 사항 이외 삭제본',variable=self.var,value='reduced',style='Mode.TRadiobutton').pack(anchor='w',pady=(16,8))
        ent.tk.Label(body,text='주간회의는 업데이트된 페이지만, Issue DB는 헤더와 업데이트된 행만 남깁니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',padx=24)
        foot=ent.tk.Frame(self,bg='#F6F8FA',height=64); foot.pack(fill='x'); foot.pack_propagate(False)
        ent.tk.Button(foot,text='선택 완료',command=self.ok,bg=ui.BLUE,fg='white',bd=0,font=('Malgun Gothic',9,'bold'),width=14,pady=8).pack(side='right',padx=20,pady=13)
        self.protocol('WM_DELETE_WINDOW',self.cancel); ui.center_window(self,parent,620,330); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def ok(self):self.result=self.var.get(); self.destroy()
    def cancel(self):self.result='full'; self.destroy()


def _select_and_finish(parent,title,message):
    if title!='업데이트 완료':
        return _original_info(parent,title,message)
    dlg=OutputTypeDialog(parent)
    if dlg.result=='reduced':
        made=[]; errors=[]
        weekly=parent.vars.get('pptweekly').get().strip() if parent.vars.get('pptweekly') else ''
        excel=parent.vars.get('xlsx').get().strip() if parent.vars.get('xlsx') else ''
        try:
            if weekly:
                saved=Path(getattr(parent,'_last_saved_outputs',{}).get('weekly','')) if getattr(parent,'_last_saved_outputs',{}).get('weekly') else _latest_version(weekly)
                if saved and Path(saved).exists():
                    reduced,n=_ppt_update_only(weekly,saved)
                    if reduced:
                        made.append(f'주간회의: 업데이트 페이지 {n}개')
                        try:Path(saved).unlink()
                        except Exception:pass
        except Exception as e:errors.append('주간회의 축약본 생성 실패: '+str(e))
        try:
            if excel:
                saved=Path(getattr(parent,'_last_saved_outputs',{}).get('excel','')) if getattr(parent,'_last_saved_outputs',{}).get('excel') else _latest_version(excel)
                if saved and Path(saved).exists():
                    reduced,n=_excel_update_only(excel,saved,getattr(parent,'_last_excel_updated_row',None))
                    if reduced:
                        made.append(f'Issue DB: 업데이트 행 {n}개')
                        try:Path(saved).unlink()
                        except Exception:pass
        except Exception as e:errors.append('Issue DB 축약본 생성 실패: '+str(e))
        if errors:
            extra='\n'.join(made+errors)
            return _original_info(parent,'업데이트 완료','일부 축약본 생성에 실패했습니다. 전체 업데이트 본은 보존했습니다.\n\n'+extra)
        if not made:
            return _original_info(parent,'업데이트 완료','변경된 행/페이지를 찾지 못해 축약본을 만들지 않았습니다. 전체 업데이트 본은 보존했습니다.')
        extra='\n'.join(made)
        return _original_info(parent,'업데이트 완료','업데이트 사항 이외 삭제본으로 저장했습니다.\n\n'+extra)
    return _original_info(parent,'업데이트 완료',message+'\n\n선택 유형: 전체 업데이트 본')


ui.info=_select_and_finish
,fp)
            fp_row=int(m.group(1)) if m else None
        if fp_row and int(fp_row)>=7:
            top_cell='A7'
    except Exception:
        pass

    try:ws.sheet_view.topLeftCell=top_cell
    except Exception:pass

    # Select the surviving updated row so Excel opens with the relevant record in focus.
    target=f'A{max(7,int(updated_row or 7))}'
    try:
        sels=list(ws.sheet_view.selection or [])
        if sels:
            sels[0].activeCell=target
            sels[0].sqref=target
        else:
            from openpyxl.worksheet.views import Selection
            ws.sheet_view.selection=[Selection(activeCell=target,sqref=target)]
    except Exception:
        pass

def _excel_update_only(source,saved,preferred_row=None):
    """Create a true reduced Issue DB: rows 1-6 + only the updated data row(s).

    All other data rows and their floating representative images/charts are removed.
    The surviving updated row(s) are compacted to row 7 onward.
    """
    src=load_workbook(source,rich_text=True)
    dst=load_workbook(saved,rich_text=True)
    ws0=src['Sheet1'] if 'Sheet1' in src.sheetnames else src.active
    ws=dst['Sheet1'] if 'Sheet1' in dst.sheetnames else dst.active

    changed=_detect_changed_rows(ws0,ws)
    if preferred_row and 7<=int(preferred_row)<=ws.max_row:
        preferred_row=int(preferred_row)
        # Exact row from the just-finished update wins for existing-issue updates.
        if preferred_row not in changed:
            changed=[preferred_row]
        else:
            changed=[preferred_row]+[r for r in changed if r!=preferred_row]

    if not changed:
        return None,0

    changed=sorted(set(r for r in changed if 7<=r<=ws.max_row))
    _prune_data_drawings(ws,changed)

    # Remove every old data row except the row(s) that were just updated.
    # Deleting bottom-up naturally compacts the kept rows to 7, 8, ...
    keep=set(changed)
    for r in range(ws.max_row,6,-1):
        if r not in keep:
            ws.delete_rows(r,1)

    target=Path(saved).with_name(Path(saved).stem+'_업데이트사항만'+Path(saved).suffix)
    dst.save(target)
    return target,len(changed)


class OutputTypeDialog(ent.tk.Toplevel):
    def __init__(self,parent):
        super().__init__(parent); self.withdraw(); self.result=None
        self.title('업데이트 파일 유형 선택'); self.configure(bg='white'); self.resizable(False,False); self.transient(parent)
        head=ent.tk.Frame(self,bg=ui.NAVY,height=58); head.pack(fill='x'); head.pack_propagate(False)
        ent.tk.Label(head,text='업데이트 파일 유형 선택',bg=ui.NAVY,fg='white',font=('Malgun Gothic',12,'bold')).pack(side='left',padx=22)
        body=ent.tk.Frame(self,bg='white'); body.pack(fill='both',expand=True,padx=24,pady=20)
        ent.tk.Label(body,text='저장할 Output 파일 유형을 선택해 주세요.',bg='white',fg=ui.TEXT,font=('Malgun Gothic',10,'bold')).pack(anchor='w',pady=(0,14))
        self.var=ent.tk.StringVar(value='full')
        ent.ttk.Radiobutton(body,text='1. 전체 업데이트 본',variable=self.var,value='full',style='Mode.TRadiobutton').pack(anchor='w',pady=8)
        ent.tk.Label(body,text='기존 주간회의 전체 페이지 / Issue DB 전체 행을 유지합니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',padx=24)
        ent.ttk.Radiobutton(body,text='2. 업데이트 사항 이외 삭제본',variable=self.var,value='reduced',style='Mode.TRadiobutton').pack(anchor='w',pady=(16,8))
        ent.tk.Label(body,text='주간회의는 업데이트된 페이지만, Issue DB는 헤더와 업데이트된 행만 남깁니다.',bg='white',fg=ui.MUTED,font=('Malgun Gothic',8)).pack(anchor='w',padx=24)
        foot=ent.tk.Frame(self,bg='#F6F8FA',height=64); foot.pack(fill='x'); foot.pack_propagate(False)
        ent.tk.Button(foot,text='선택 완료',command=self.ok,bg=ui.BLUE,fg='white',bd=0,font=('Malgun Gothic',9,'bold'),width=14,pady=8).pack(side='right',padx=20,pady=13)
        self.protocol('WM_DELETE_WINDOW',self.cancel); ui.center_window(self,parent,620,330); self.deiconify(); self.grab_set(); self.focus_force(); parent.wait_window(self)
    def ok(self):self.result=self.var.get(); self.destroy()
    def cancel(self):self.result='full'; self.destroy()


def _select_and_finish(parent,title,message):
    if title!='업데이트 완료':
        return _original_info(parent,title,message)
    dlg=OutputTypeDialog(parent)
    if dlg.result=='reduced':
        made=[]; errors=[]
        weekly=parent.vars.get('pptweekly').get().strip() if parent.vars.get('pptweekly') else ''
        excel=parent.vars.get('xlsx').get().strip() if parent.vars.get('xlsx') else ''
        try:
            if weekly:
                saved=Path(getattr(parent,'_last_saved_outputs',{}).get('weekly','')) if getattr(parent,'_last_saved_outputs',{}).get('weekly') else _latest_version(weekly)
                if saved and Path(saved).exists():
                    reduced,n=_ppt_update_only(weekly,saved)
                    if reduced:
                        made.append(f'주간회의: 업데이트 페이지 {n}개')
                        try:Path(saved).unlink()
                        except Exception:pass
        except Exception as e:errors.append('주간회의 축약본 생성 실패: '+str(e))
        try:
            if excel:
                saved=Path(getattr(parent,'_last_saved_outputs',{}).get('excel','')) if getattr(parent,'_last_saved_outputs',{}).get('excel') else _latest_version(excel)
                if saved and Path(saved).exists():
                    reduced,n=_excel_update_only(excel,saved,getattr(parent,'_last_excel_updated_row',None))
                    if reduced:
                        made.append(f'Issue DB: 업데이트 행 {n}개')
                        try:Path(saved).unlink()
                        except Exception:pass
        except Exception as e:errors.append('Issue DB 축약본 생성 실패: '+str(e))
        if errors:
            extra='\n'.join(made+errors)
            return _original_info(parent,'업데이트 완료','일부 축약본 생성에 실패했습니다. 전체 업데이트 본은 보존했습니다.\n\n'+extra)
        if not made:
            return _original_info(parent,'업데이트 완료','변경된 행/페이지를 찾지 못해 축약본을 만들지 않았습니다. 전체 업데이트 본은 보존했습니다.')
        extra='\n'.join(made)
        return _original_info(parent,'업데이트 완료','업데이트 사항 이외 삭제본으로 저장했습니다.\n\n'+extra)
    return _original_info(parent,'업데이트 완료',message+'\n\n선택 유형: 전체 업데이트 본')


ui.info=_select_and_finish
