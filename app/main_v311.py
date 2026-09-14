# 8D Issue Automation v3.1.1
# Focused fixes on top of clean v3.1.0 renderer.
import re, datetime
from pathlib import Path

import main_v310 as v310
from pptx import Presentation
from pptx.util import Pt

base = v310.base
N = v310.N
C = v310.C
walk = v310.walk
box = v310.box
set_text = v310.set_text


def _progress_text(d):
    """Safe local progress text; no dependency on old v3.x patch chain."""
    for key in ('progress', 'progress_status', 'current_progress', 'status_detail'):
        val = N(d.get(key))
        if val:
            return val
    parts = []
    for key in ('temporary_action', 'cause_4d', 'action_5d', 'verification_6d'):
        val = N(d.get(key))
        if val:
            parts.append(val)
    return '\n'.join(parts[-2:]) if parts else N(d.get('status'))


def _set_cell_text_preserve_size(cell, text, size=8):
    cell.text = N(text)
    for p in cell.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(size)


def update_page1(prs, d, g):
    """Update summary table and title independently so one never prevents the other."""
    if not prs.slides:
        return
    sl = prs.slides[0]
    team = N(g.get('team'))
    values = {
        '과제명': N(d.get('task_name')),
        '이슈': N(d.get('issue_name')),
        '문제/현상': N(d.get('problem')),
        '진행': _progress_text(d),
    }

    # 1) Update every matching title shape first; preserve existing run formatting.
    if team:
        for sh in walk(sl):
            if not hasattr(sh, 'text_frame'):
                continue
            full = N(getattr(sh, 'text', ''))
            if '팀 주요 논의 사항' not in full:
                continue
            changed = False
            for p in sh.text_frame.paragraphs:
                for run in p.runs:
                    if '팀 주요 논의 사항' in run.text:
                        run.text = re.sub(r'[^\s]*팀\s*주요\s*논의\s*사항', team + '팀 주요 논의 사항', run.text)
                        changed = True
            if not changed:
                # Split-run title fallback. Only text is changed; no global font rewrite.
                sh.text_frame.text = re.sub(r'[^\s]*팀\s*주요\s*논의\s*사항', team + '팀 주요 논의 사항', full)

    # 2) Update summary table separately.
    for sh in walk(sl):
        if not getattr(sh, 'has_table', False):
            continue
        tb = sh.table
        hm = {}
        for c in range(len(tb.columns)):
            q = C(tb.cell(0, c).text)
            if '과제명' in q:
                hm['과제명'] = c
            elif q == '이슈' or '이슈명' in q:
                hm['이슈'] = c
            elif '문제' in q or '현상' in q:
                hm['문제/현상'] = c
            elif '진행' in q:
                hm['진행'] = c
            elif 'signal' in q:
                hm['Signal'] = c
        if not ({'과제명', '이슈'} & set(hm)):
            continue
        row = 1 if len(tb.rows) > 1 else 0
        for r in range(1, len(tb.rows)):
            if not any(N(tb.cell(r, c).text) for c in range(len(tb.columns))):
                row = r
                break
        for key, c in hm.items():
            if key in values:
                _set_cell_text_preserve_size(tb.cell(row, c), values[key], 8)
        if 'Signal' in hm:
            try:
                base.signal(tb.cell(row, hm['Signal']), base.status(d))
            except Exception:
                pass
        break


def _occurrence_value(d, g):
    stage = N(g.get('stage') or g.get('occurrence_stage') or d.get('occurrence_stage') or d.get('stage'))
    date = N(d.get('occurrence_date') or g.get('occurrence_date'))
    if stage and date:
        return f'{stage} ({date})'
    return stage or date


def _fill_occurrence(sl, d, g):
    target = _occurrence_value(d, g)
    if not target:
        return

    # Table form.
    for sh in walk(sl):
        if not getattr(sh, 'has_table', False):
            continue
        tb = sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                q = C(tb.cell(r, c).text)
                if '발생단계' in q and '발생일자' in q and c + 1 < len(tb.columns):
                    _set_cell_text_preserve_size(tb.cell(r, c + 1), target, 9)
                    return

    # Text-box form: fill closest box to the right of the label.
    labels = []
    for sh in walk(sl):
        if hasattr(sh, 'text_frame'):
            q = C(getattr(sh, 'text', ''))
            if '발생단계' in q and '발생일자' in q:
                labels.append(sh)
    for lab in labels:
        lx, ly, lw, lh = box(lab)
        best = None
        best_score = float('inf')
        for sh in walk(sl):
            if sh is lab or not hasattr(sh, 'text_frame'):
                continue
            sx, sy, sw, shh = box(sh)
            gap = sx - (lx + lw)
            if gap < -0.05 * v310.EMU:
                continue
            if abs(sy - ly) > 0.6 * v310.EMU:
                continue
            score = max(0, gap) + abs(sy - ly)
            if score < best_score:
                best, best_score = sh, score
        if best is not None:
            set_text(best, target, 9)
            return


def update_page2(sl, d, g, mode):
    # Keep v3.1.0's safe section/image renderer, then apply reliable occurrence fill.
    v310.update_page2(sl, d, g, mode)
    _fill_occurrence(sl, d, g)


def weekly(src, out, d, g, mode):
    prs = Presentation(src)
    update_page1(prs, d, g)
    if len(prs.slides) > 1:
        update_page2(prs.slides[1], d, g, mode)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(out)
        saved = out
    except PermissionError:
        p = Path(out)
        saved = str(p.with_name(p.stem + '_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + p.suffix))
        prs.save(saved)
    return '주간회의 PPT 업데이트: ' + base.status(d), saved


base.weekly = weekly
base.extract = v310.extract

if __name__ == '__main__':
    base.App().mainloop()
