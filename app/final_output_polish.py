"""Final output polish: persistent version names, Excel date/status styling, PPT empty-placeholder cleanup."""
import datetime
import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import main_enterprise as ent
import main_recovery_step14 as s14

base=ent.base
BLUE="0033FF"
_original_excel=base.update_excel
_original_weekly=base.weekly


def _version_path(requested):
    """Return root_v0.N.ext in the requested output folder, scanning existing versions."""
    p=Path(requested)
    stem=re.sub(r'_업데이트$','',p.stem,flags=re.I)
    m=re.match(r'^(.*)_v(\\d+)$',stem,re.I)
    root=m.group(1) if m else stem
    current=int(m.group(2)) if m else 0
    pat=re.compile(rf'^{re.escape(root)}_v(\\d+){re.escape(p.suffix)}$',re.I)
    nums=[current]
    if p.parent.exists():
        for f in p.parent.iterdir():
            mm=pat.match(f.name)
            if mm: nums.append(int(mm.group(1)))
    n=max(nums)+1
    return p.with_name(f'{root}_v0.{n}{p.suffix}')


def _blue_inline(cell):
    f=cell.font
    return InlineFont(rFont=f.name or '맑은 고딕',sz=f.sz,b=f.b,i=f.i,strike=f.strike,color=BLUE)


def _date_text(value):
    """Return a stable date-only display string; never preserve an Excel time part."""
    if isinstance(value,datetime.datetime):return value.strftime('%Y-%m-%d')
    if isinstance(value,datetime.date):return value.strftime('%Y-%m-%d')
    text=str(value or '').strip()
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%d','%Y/%m/%d %H:%M:%S','%Y/%m/%d'):
        try:return datetime.datetime.strptime(text,fmt).strftime('%Y-%m-%d')
        except ValueError:pass
    m=re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s+.*)?$',text)
    if m:
        try:return datetime.date(int(m.group(1)),int(m.group(2)),int(m.group(3))).strftime('%Y-%m-%d')
        except ValueError:pass
    return text


def _polish_excel(saved,src,d,g,new):
    """Final modified date is display text; changed issue status is wholly blue."""
    try:
        wb=load_workbook(saved,rich_text=True); ws=wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active
        row,_=base.find(ws,d,g)
        if not row:return
        date_cell=ws.cell(row,2)
        date_cell.value=_date_text(date_cell.value)
        date_cell.number_format='@'
        old_status=''
        if not new:
            try:
                wb0=load_workbook(src,rich_text=True); ws0=wb0['Sheet1'] if 'Sheet1' in wb0.sheetnames else wb0.active
                old_row,_=base.find(ws0,d,g)
                if old_row:old_status=str(ws0.cell(old_row,18).value or '')
            except Exception:pass
        cell=ws.cell(row,18); new_status=str(cell.value or '')
        if new_status and (new or old_status.strip().casefold()!=new_status.strip().casefold()):
            cell.value=CellRichText([TextBlock(_blue_inline(cell),new_status)])
        wb.save(saved)
    except Exception:
        pass


def update_excel_versioned(src,out,d,g,new=False):
    versioned=_version_path(out)
    msg,saved=_original_excel(src,versioned,d,g,new=new)
    _polish_excel(saved,src,d,g,new)
    return msg+f' / 버전 저장={Path(saved).name}',saved


def _empty_placeholder(sh):
    try:
        if sh.shape_type!=MSO_SHAPE_TYPE.PLACEHOLDER:return False
        if getattr(sh,'has_table',False):return False
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:return False
        return not str(getattr(sh,'text','') or '').strip()
    except Exception:return False


def _clean_summary_placeholders(saved):
    """Remove unused content placeholders on weekly summary pages (table/chart icons in edit view)."""
    try:
        prs=Presentation(saved); changed=False
        pages=s14._summary_pages(prs)
        indices={x[0] for x in pages}
        for i in indices:
            if i<0 or i>=len(prs.slides):continue
            sl=prs.slides[i]
            for sh in list(sl.shapes):
                if _empty_placeholder(sh):
                    try:sh._element.getparent().remove(sh._element); changed=True
                    except Exception:pass
        if changed:prs.save(saved)
    except Exception:pass


def weekly_versioned(src,out,d,g,mode):
    versioned=_version_path(out)
    msg,saved=_original_weekly(src,versioned,d,g,mode)
    _clean_summary_placeholders(saved)
    return msg+f' / 버전 저장={Path(saved).name}',saved

base.update_excel=update_excel_versioned
base.weekly=weekly_versioned
