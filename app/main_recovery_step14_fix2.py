import datetime
import hashlib
import subprocess
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_recovery_step12 as s12
import main_recovery_step10 as s10
import main_v319 as v319
import main_v310 as v310

base=s14.base
N=v310.N


# -----------------------------------------------------------------------------
# Project/detail-page detection
# -----------------------------------------------------------------------------
def _project_parts(d):
    return (
        s13._k(s13._customer_task(d)),
        s13._k(d.get('task_name')),
        s13._k(d.get('customer')),
    )


def _project_score(sl,d):
    q=s13._k(s13._slide_text(sl))
    full,task,customer=_project_parts(d)
    if full and full in q:
        return 160
    if task and customer and task in q and customer in q:
        return 145
    if task and task in q:
        return 120
    return 0


def _summary_indices(prs):
    return {si for si,_,_ in s14._summary_pages(prs)}


def _is_detail_like(sl):
    text=s13._slide_text(sl)
    q=s13._k(text)
    d_hits=sum(1 for token in ('2d','3d','4d','5d','6d') if token in q)
    meta_hits=sum(1 for token in ('signal','이슈기인','발생단계') if token in q)
    return d_hits>=3 or (d_hits>=2 and meta_hits>=1)


# -----------------------------------------------------------------------------
# PowerPoint native section handling
# -----------------------------------------------------------------------------
def _slide_id_pairs(prs):
    out=[]
    for i,sldId in enumerate(prs.slides._sldIdLst):
        try:
            sid=int(sldId.get('id'))
        except Exception:
            try: sid=int(sldId.id)
            except Exception: continue
        out.append((i,sid))
    return out


def _native_sections(prs):
    root=prs._element
    id_to_index={sid:i for i,sid in _slide_id_pairs(prs)}
    sections=[]
    try:
        sec_nodes=root.xpath('.//*[local-name()="sectionLst"]/*[local-name()="section"]')
    except Exception:
        sec_nodes=[]
    for sec in sec_nodes:
        name=N(sec.get('name'))
        ids=[]
        try:
            id_nodes=sec.xpath('.//*[local-name()="sldId"]')
        except Exception:
            id_nodes=[]
        for n in id_nodes:
            raw=n.get('id')
            try:
                sid=int(raw)
            except Exception:
                continue
            ids.append(sid)
        indices=sorted(id_to_index[sid] for sid in ids if sid in id_to_index)
        sections.append({'name':name,'element':sec,'slide_ids':ids,'indices':indices})
    return sections


def _section_match_score(name,d):
    q=s13._k(name)
    full,task,customer=_project_parts(d)
    if not q:
        return 0
    if full and q==full:
        return 300
    if task and q==task:
        return 280
    if full and full in q:
        return 260
    if task and customer and task in q and customer in q:
        return 250
    if task and task in q:
        return 220
    return 0


def _matching_native_section(prs,d):
    scored=[]
    for sec in _native_sections(prs):
        sc=_section_match_score(sec['name'],d)
        if sc:
            scored.append((sc,sec))
    if not scored:
        return None
    scored.sort(key=lambda x:(x[0],len(x[1]['indices'])),reverse=True)
    return scored[0][1]


def _slide_id_at(prs,index):
    sldId=prs.slides._sldIdLst[index]
    try: return int(sldId.get('id'))
    except Exception:
        try: return int(sldId.id)
        except Exception: return None


def _add_slide_to_native_section(prs,section,slide_index):
    if not section:
        return False
    sid=_slide_id_at(prs,slide_index)
    if sid is None:
        return False
    sec=section['element']
    try:
        for n in sec.xpath('.//*[local-name()="sldId"]'):
            if str(n.get('id'))==str(sid):
                return True
        lists=sec.xpath('./*[local-name()="sldIdLst"]')
        if not lists:
            return False
        lst=lists[0]
        existing=lst.xpath('./*[local-name()="sldId"]')
        if not existing:
            return False
        import copy
        node=copy.deepcopy(existing[-1])
        node.set('id',str(sid))
        lst.append(node)
        return True
    except Exception:
        return False


def _project_detail_indices(prs,d):
    sec=_matching_native_section(prs,d)
    summaries=_summary_indices(prs)
    if sec and sec['indices']:
        return [i for i in sec['indices'] if i not in summaries and _is_detail_like(prs.slides[i])]
    return [i for i,sl in enumerate(prs.slides)
            if i not in summaries and _is_detail_like(sl) and _project_score(sl,d)>=120]


def _find_existing_detail(prs,d):
    pages=_project_detail_indices(prs,d)
    if not pages:
        return None
    issue=s13._k(s13._issue_display(d))
    best=None
    for i in pages:
        text=s13._slide_text(prs.slides[i])
        q=s13._k(text)
        score=200
        score+=s13._detail_structure_score(prs.slides[i])*4
        if issue and issue in q:
            score+=150
        else:
            from difflib import SequenceMatcher
            ratios=[SequenceMatcher(None,issue,s13._k(x)).ratio()
                    for x in text.splitlines() if s13._k(x)] if issue else []
            if ratios:
                score+=int(70*max(ratios))
        if best is None or score>best[0]:
            best=(score,i)
    return best[1] if best and best[0]>=285 else None


def _new_detail_position(prs,d):
    sec=_matching_native_section(prs,d)
    if sec and sec['indices']:
        return max(sec['indices'])+1
    pages=_project_detail_indices(prs,d)
    return max(pages)+1 if pages else len(prs.slides)


# -----------------------------------------------------------------------------
# Clone another real detail page ONLY as a layout shell.
# -----------------------------------------------------------------------------
def _template_detail_index(prs,d):
    summaries=_summary_indices(prs)
    same=_project_detail_indices(prs,d)
    if same:
        return same[-1]

    candidates=[]
    for i,sl in enumerate(prs.slides):
        if i in summaries or not _is_detail_like(sl):
            continue
        structure=s13._detail_structure_score(sl)
        q=s13._k(s13._slide_text(sl))
        meta=sum(1 for x in ('signal','이슈기인','발생단계') if x in q)
        candidates.append((structure*10+meta*8,i))
    if not candidates:
        raise ValueError('주간회의 PPT에서 복제할 수 있는 2D~6D 상세 페이지 양식을 찾지 못했습니다.')
    candidates.sort(reverse=True)
    return candidates[0][1]


def _overlap_ratio(sh,zone):
    try:
        x=float(sh.left)/v310.EMU; y=float(sh.top)/v310.EMU
        w=float(sh.width)/v310.EMU; h=float(sh.height)/v310.EMU
    except Exception:
        return 0.0
    zx,zy,zw,zh=zone
    ix=max(0,min(x+w,zx+zw)-max(x,zx))
    iy=max(0,min(y+h,zy+zh)-max(y,zy))
    inter=ix*iy
    area=max(w*h,1e-6)
    return inter/area


def _content_zones():
    return [
        (.38,2.35,5.05,1.30),
        (.38,3.72,5.05,1.50),
        (.38,5.18,5.05,2.05),
        (5.58,2.30,5.10,1.25),
        (5.58,3.56,5.10,2.10),
        (5.58,5.72,5.10,1.25),
    ]


def _static_detail_label(text):
    q=s13._k(text)
    static={
        '2d','3d','4d','5d','6d','7d',
        '현상','임시대책필요시','원인분석','개선대책','유효성점검','수평전개',
        'signal','이슈기인','발생단계'
    }
    return q in static


def _clear_cloned_issue_content(sl):
    v310.remove_previous_auto(sl)
    zones=_content_zones()

    for sh in list(sl.shapes):
        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
            continue

        in_work_area=any(_overlap_ratio(sh,z)>=0.35 for z in zones)
        if not in_work_area:
            continue

        if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.PICTURE:
            try:
                sh._element.getparent().remove(sh._element)
            except Exception:
                pass
            continue

        if hasattr(sh,'text_frame'):
            old=N(getattr(sh,'text',''))
            if old and not _static_detail_label(old):
                try:
                    sh.text=''
                except Exception:
                    try: sh.text_frame.clear()
                    except Exception: pass


def _clone_detail_shell(prs,d,insert_at):
    template_index=_template_detail_index(prs,d)
    matched_section=_matching_native_section(prs,d)
    s13._clone_slide_with_rels(prs,template_index)
    s13._move_last_slide_to(prs,insert_at)
    sl=prs.slides[insert_at]
    _clear_cloned_issue_content(sl)
    _add_slide_to_native_section(prs,matched_section,insert_at)
    return sl,template_index,matched_section


# -----------------------------------------------------------------------------
# Append source 8D pages 2..end as native PowerPoint slides.
# This uses PowerPoint itself (via Windows PowerShell COM) so pictures, groups,
# charts, fonts and the source slide appearance are preserved exactly.  The pages
# are inserted immediately AFTER the issue detail page, therefore they also sit in
# the same project section in normal PowerPoint section order.
# -----------------------------------------------------------------------------
def _attachment_key(d):
    raw=(s13._customer_task(d)+'|'+s13._issue_display(d)).encode('utf-8','ignore')
    return hashlib.sha1(raw).hexdigest()[:12]


def _ps_quote(path):
    return str(path).replace("'","''")


def _append_8d_attachments(out_path,src8d_path,detail_index,d):
    try:
        source_count=len(Presentation(src8d_path).slides)
    except Exception as e:
        raise RuntimeError('8D 원본의 유첨 페이지 수를 확인하지 못했습니다: '+repr(e))
    if source_count<=1:
        return 0

    # PowerPoint COM is used instead of python-pptx cross-file XML copying because
    # the latter can lose source relationships/theme parts and corrupt the layout.
    # detail_index is zero-based here; PowerPoint COM is one-based.
    insert_after=detail_index+1
    key=_attachment_key(d)
    prefix='AUTO_8D_ATTACH_'+key+'_'

    out_q=_ps_quote(Path(out_path).resolve())
    src_q=_ps_quote(Path(src8d_path).resolve())
    prefix_q=prefix.replace("'","''")

    script=f"""
$ErrorActionPreference = 'Stop'
$ppt = $null
$pres = $null
try {{
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Open('{out_q}', 0, 0, 0)

    # Remove only attachment slides previously generated for this same issue.
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) {{
        $nm = [string]$pres.Slides.Item($i).Name
        if ($nm.StartsWith('{prefix_q}')) {{
            $pres.Slides.Item($i).Delete()
        }}
    }}

    $added = $pres.Slides.InsertFromFile('{src_q}', {insert_after}, 2, {source_count})
    for ($j = 1; $j -le $added; $j++) {{
        $pres.Slides.Item({insert_after} + $j).Name = '{prefix_q}' + $j
    }}
    $pres.Save()
    Write-Output $added
}}
finally {{
    if ($pres -ne $null) {{ $pres.Close() }}
    if ($ppt -ne $null) {{ $ppt.Quit() }}
}}
"""

    try:
        p=subprocess.run(
            ['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-Command',script],
            capture_output=True,text=True,timeout=120
        )
    except FileNotFoundError:
        raise RuntimeError('Windows PowerShell을 찾지 못해 8D 유첨 페이지를 붙이지 못했습니다.')
    except subprocess.TimeoutExpired:
        raise RuntimeError('PowerPoint 유첨 페이지 복사 작업이 120초를 초과했습니다.')

    if p.returncode!=0:
        detail=(p.stderr or p.stdout or '').strip()
        raise RuntimeError('8D 유첨 페이지 자동 복사 실패: '+detail[-800:])

    # InsertFromFile returns the number of inserted slides; parse it if possible.
    nums=[]
    for line in (p.stdout or '').splitlines():
        q=line.strip()
        if q.isdigit():
            nums.append(int(q))
    return nums[-1] if nums else source_count-1


# -----------------------------------------------------------------------------
# Weekly update
# -----------------------------------------------------------------------------
def weekly_fix5(src,out,d,g,mode):
    prs=Presentation(src)
    _,_,summary_action=s14._update_summary_by_task(prs,d,g,mode)

    matched_section=_matching_native_section(prs,d)
    target=_find_existing_detail(prs,d) if mode=='existing' else None
    detail_action='업데이트'

    if target is None:
        insert_at=_new_detail_position(prs,d)
        _,template_index,matched_section=_clone_detail_shell(prs,d,insert_at)
        target=insert_at
        sec_name=N(matched_section.get('name')) if matched_section else '(텍스트 기반 구역)'
        detail_action=(f'과제 구역 [{sec_name}] 끝에 상세 페이지 추가 '
                       f'(양식 원본 slide {template_index+1}, 기존 내용 초기화)')

    s13._update_detail_slide(prs.slides[target],d,g,mode)

    Path(out).parent.mkdir(parents=True,exist_ok=True)
    try:
        prs.save(out); saved=out
    except PermissionError:
        p=Path(out)
        saved=str(p.with_name(p.stem+'_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+p.suffix))
        prs.save(saved)

    # After the weekly file is safely saved, append every source 8D slide from
    # page 2 onward directly behind this issue's detail page.
    attached=_append_8d_attachments(saved,g.get('ppt8d',''),target,d)

    sec_name=N(matched_section.get('name')) if matched_section else '텍스트 기반 탐색'
    attach_msg=f' / 8D 유첨 {attached}페이지 추가' if attached else ' / 8D 유첨 없음'
    return ('주간회의 PPT: 요약 '+summary_action+' / 상세 '+detail_action+
            f' / 배치 구역={sec_name} / 과제명={N(d.get("task_name"))}'+attach_msg,saved)


base.weekly=weekly_fix5


class RecoveryStep14Fix5App(s10.RecoveryStep10App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP14 FIX5')


if __name__=='__main__':
    RecoveryStep14Fix5App().mainloop()
