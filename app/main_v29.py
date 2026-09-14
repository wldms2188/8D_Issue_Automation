# 8D Issue Automation v2.9 patch layer
# v2.9 keeps v2.8 extraction/output logic and fixes:
# 1) preserve line breaks and never add ellipsis/truncate issue text
# 2) weekly meeting text = Malgun Gothic 8pt
# 3) robust slide-2/detail-box filling with semantic + geometry fallback
# 4) representative photo = image attached to the 2D/problem area, not largest image in PPT
# 5) Excel photo is fitted inside the representative-photo cell area

import io, re, copy, datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

import main_v28 as base

# ---------- text handling ----------
def v29_one(x, n=None):
    """Keep meaningful line breaks. Never truncate with an ellipsis."""
    s = base.norm(x)
    # Do not collapse newlines. Remove trailing spaces on each line only.
    s = '\n'.join(line.rstrip() for line in s.split('\n'))
    return s.strip()

base.one = v29_one


def v29_progress(d):
    def block(label, items):
        lines = [f'{label}:']
        found = False
        for name, value in items:
            value = v29_one(value)
            if value:
                lines.append(f'  {name}:')
                lines.append(value)
                found = True
        if not found:
            lines.append('  검토 중')
        return '\n'.join(lines)

    return '\n\n'.join([
        block('원인', [
            ('발생 원인', d.get('cause_4d')),
            ('유출 원인', d.get('leak_cause')),
            ('시스템 원인', d.get('system_cause')),
        ]),
        block('진행 현황', [
            ('임시 조치', d.get('temporary_action')),
            ('고객 대응', d.get('customer_response')),
            ('개선 대책', d.get('action_5d')),
            ('유효성 검증', d.get('verification_6d')),
            ('수평 전개', d.get('spread_7d')),
        ]),
        '요청 사항:\n  '
    ])

base.progress = v29_progress

# ---------- representative image selection ----------
def _shape_center(sh):
    return (float(sh.left + sh.width / 2), float(sh.top + sh.height / 2))


def _distance(a, b):
    ax, ay = a; bx, by = b
    return ((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5


def _shape_texts(slide):
    out = []
    for sh in base.flat(slide):
        t = base.norm(getattr(sh, 'text', ''))
        if t:
            out.append((sh, t))
        if getattr(sh, 'has_table', False):
            tb = sh.table
            if any(base.norm(tb.cell(r, c).text) for r in range(len(tb.rows)) for c in range(len(tb.columns))):
                cell_text = '\n'.join(base.norm(tb.cell(r, c).text) for r in range(len(tb.rows)) for c in range(len(tb.columns)))
                out.append((sh, cell_text))
    return out


def _pick_2d_picture(prs, problem=''):
    pictures = []
    anchors = []
    pkey = base.compact(problem)[:40] if problem else ''

    for si, sl in enumerate(prs.slides):
        texts = _shape_texts(sl)
        for sh, txt in texts:
            q = base.compact(txt)
            if ('2d' in q or '문제현황' in q or '불량현상' in q or (pkey and pkey in q)):
                anchors.append((si, _shape_center(sh)))
        for sh in base.flat(sl):
            if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = sh.image.blob
                    with PILImage.open(io.BytesIO(blob)) as im:
                        w, h = im.size
                    if w * h >= 10000:
                        pictures.append((si, _shape_center(sh), w * h, w, h, blob))
                except Exception:
                    pass

    if not pictures:
        return None

    # Strong preference: same slide as 2D/problem area, nearest picture.
    if anchors:
        scored = []
        for pic in pictures:
            si, center, area, w, h, blob = pic
            same = [a for a in anchors if a[0] == si]
            if same:
                dist = min(_distance(center, a[1]) for a in same)
                scored.append((0, dist, -area, blob))
        if scored:
            scored.sort(key=lambda x: (x[0], x[1], x[2]))
            return scored[0][3]

    # If no 2D picture is on the same slide, prefer the slide containing 2D.
    if anchors:
        anchor_slides = {a[0] for a in anchors}
        same_slide = [p for p in pictures if p[0] in anchor_slides]
        if same_slide:
            same_slide.sort(key=lambda x: -x[2])
            return same_slide[0][5]

    # Last-resort fallback: largest image.
    return max(pictures, key=lambda x: x[2])[5]


def v29_extract(path):
    d = base._v29_original_extract(path) if hasattr(base, '_v29_original_extract') else base.extract(path)
    try:
        prs = Presentation(path)
        chosen = _pick_2d_picture(prs, d.get('problem', ''))
        if chosen:
            d['_images'] = [(1, 1, 1, chosen)]
    except Exception:
        pass
    return d

base._v29_original_extract = base.extract
base.extract = v29_extract

# ---------- Excel: no truncation + representative photo fitted to O cell ----------
def v29_add_excel_photo(ws, row, blob):
    if not blob:
        return
    try:
        # Remove an existing image anchored to the same row, preventing overlays.
        for old in list(getattr(ws, '_images', [])):
            try:
                anchor = old.anchor._from
                if anchor.row == row - 1 and anchor.col == 14:
                    ws._images.remove(old)
            except Exception:
                pass

        with PILImage.open(io.BytesIO(blob)) as im:
            w, h = im.size
            col_width = ws.column_dimensions['O'].width or 20
            # Excel column width -> approximate pixels; leave a small margin.
            max_w = max(60, int(col_width * 7) - 8)
            max_h = 95
            scale = min(max_w / max(w, 1), max_h / max(h, 1))
            new_w = max(1, int(w * scale))
            new_h = max(1, int(h * scale))
            img = XLImage(io.BytesIO(blob))
            img.width = new_w
            img.height = new_h
            ws.add_image(img, f'O{row}')
            ws.row_dimensions[row].height = max(float(ws.row_dimensions[row].height or 15), min(105, new_h * 0.75 + 8))
    except Exception:
        pass


def v29_write_row(ws, r, d, g):
    y, m = base.parse_year_month(d.get('occurrence_date'))
    st = base.status(d)
    cause = '\n'.join(x for x in [d.get('cause_4d'), d.get('leak_cause'), d.get('system_cause')] if v29_one(x))
    vals = {
        2: datetime.date.today(),
        3: g.get('plm_no', ''),
        4: g.get('form_factor', ''),
        5: g.get('product_type', ''),
        6: y,
        7: m,
        8: g.get('team', ''),
        9: g.get('owner', ''),
        10: base.task(d),
        11: g.get('sample', ''),
        12: g.get('stage', ''),
        13: v29_one(d.get('occurrence_site')),
        14: v29_one(d.get('problem')),
        15: '',
        16: cause,
        17: v29_one(d.get('action_5d')),
        18: st,
    }
    for c, v in vals.items():
        ws.cell(r, c).value = v
    v29_add_excel_photo(ws, r, base.representative_image(d))
    return st

base.write_row = v29_write_row
base.add_excel_photo = v29_add_excel_photo

# ---------- PPT text formatting ----------
def v29_font_run(run, size=8):
    run.font.name = '맑은 고딕'
    run.font.size = Pt(size)
    try:
        r = run._r.get_or_add_rPr()
        r.set('a:latin', '맑은 고딕')
        r.set('a:ea', '맑은 고딕')
        r.set('a:cs', '맑은 고딕')
    except Exception:
        pass


def v29_set_shape_text(sh, text, size=8):
    sh.text = v29_one(text)
    if hasattr(sh, 'text_frame'):
        sh.text_frame.word_wrap = True
        for p in sh.text_frame.paragraphs:
            for run in p.runs:
                v29_font_run(run, size)


def v29_set_cell_text(cl, text, size=8):
    cl.text = v29_one(text)
    cl.text_frame.word_wrap = True
    for p in cl.text_frame.paragraphs:
        for run in p.runs:
            v29_font_run(run, size)

base.font_run = v29_font_run
base.set_shape_text = v29_set_shape_text
base.set_cell_text = v29_set_cell_text

# ---------- weekly detail fallback ----------
def _fill_weekly_geometry(sl, d):
    """Fallback for weekly slide-2 boxes when numbered annotations are absent."""
    contents = base.section_contents(d)
    mapping = [('2d', 0.53, 2.53, 4.76, 0.98),
               ('3d', 0.53, 3.91, 4.76, 1.14),
               ('4d_cause', 1.25, 4.64, 3.92, 1.52),
               ('4d_cause', 0.53, 5.38, 4.76, 1.79),
               ('5d', 5.77, 3.74, 4.76, 1.77),
               ('6d', 5.77, 5.91, 4.78, 0.94),
               ('6d', 5.77, 3.74, 4.76, 1.77),
               ('3d', 5.77, 2.47, 4.76, 0.87)]
    # Use only likely content rectangles: reasonably sized, text empty or placeholder.
    candidates = []
    for sh in base.flat(sl):
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.PICTURE:
            continue
        if not hasattr(sh, 'text_frame'):
            continue
        t = base.norm(getattr(sh, 'text', ''))
        if t in {'', '예시', '내용', '입력', '내용 입력'}:
            candidates.append(sh)

    used = set()
    sw, shh = float(sl.part.slide_layout.slide_master.part.presentation.slide_width if False else 1), 1
    # Match against example-2 coordinates with tolerance in inches.
    for k, x, y, w, h in mapping:
        val = contents.get(k, '')
        if not val:
            continue
        best = None
        best_score = 10**9
        for sh in candidates:
            if id(sh) in used:
                continue
            sx, sy = sh.left / 914400, sh.top / 914400
            swid, shgt = sh.width / 914400, sh.height / 914400
            score = abs(sx-x) + abs(sy-y) + abs(swid-w) + abs(shgt-h)
            if abs(sx-x) <= 0.65 and abs(sy-y) <= 0.65 and abs(swid-w) <= 0.9 and abs(shgt-h) <= 0.9 and score < best_score:
                best, best_score = sh, score
        if best is not None:
            v29_set_shape_text(best, val, 8)
            used.add(id(best))


def v29_fill_detail_slide(sl, d, g):
    # First run v2.8 semantic/numbered logic with non-truncating text.
    base.fill_detail_slide(sl, d, g)
    # Then add geometry fallback for weekly-style blank boxes.
    _fill_weekly_geometry(sl, d)
    # Force all populated text in the detail slide to Malgun Gothic 8pt.
    for sh in base.flat(sl):
        if hasattr(sh, 'text_frame'):
            for p in sh.text_frame.paragraphs:
                for run in p.runs:
                    v29_font_run(run, 8)

base.fill_detail_slide = v29_fill_detail_slide

# ---------- replace weekly to remove summary truncation and guarantee detail pages ----------
def v29_weekly(src, out, d, g, mode):
    prs = Presentation(src)
    st = base.status(d)
    title = base.task(d)
    issue = v29_one(d.get('issue_name'))

    # Find summary table semantically.
    summary = None
    summary_idx = None
    for i, sl in enumerate(prs.slides):
        for sh in base.flat(sl):
            if getattr(sh, 'has_table', False):
                tb = sh.table
                header = ' '.join(base.norm(tb.cell(0,c).text) for c in range(len(tb.columns)))
                if '과제명' in header and 'Signal' in header:
                    summary = tb
                    summary_idx = i
                    break
        if summary is not None:
            break

    if summary is not None:
        r = None
        for x in range(1, len(summary.rows)):
            if base.key(title) and base.key(title) in base.key(summary.cell(x,0).text) and base.key(issue) in base.key(summary.cell(x,1).text):
                r = x; break
        if r is None:
            for x in range(1, len(summary.rows)):
                if not any(base.norm(summary.cell(x,c).text) for c in range(min(5,len(summary.columns))) if c != 4):
                    r = x; break
        if r is None:
            summary._tbl.append(copy.deepcopy(summary.rows[-1]._tr)); r = len(summary.rows)-1
        values = [title, issue, v29_one(d.get('problem')), v29_progress(d), '●']
        for c, v in enumerate(values):
            v29_set_cell_text(summary.cell(r,c), v, 8)
        base.signal(summary.cell(r,4), st)

    if summary_idx is None:
        summary_idx = 0

    # IMPORTANT: process every slide after the summary. This is the 2-page/detail fix.
    details = [prs.slides[i] for i in range(summary_idx + 1, len(prs.slides))]
    if not details:
        # Preserve the existing 2-page template when a source has only one slide.
        template = prs.slides[1] if len(prs.slides) > 1 else prs.slides[0]
        details = [base.clone_slide(prs, template)]

    for sl in details:
        v29_fill_detail_slide(sl, d, g)
        # Metadata and signal fallbacks.
        for tb, r, c, cl in base.table_cells(sl):
            if 'signal' in base.compact(cl.text) and c+1 < len(tb.columns):
                base.signal(tb.cell(r,c+1), st)
        for sh in base.flat(sl):
            if hasattr(sh, 'text') and sh.text:
                txt = base.norm(sh.text)
                if '00팀 담당자' in txt and '000' in txt:
                    v29_set_shape_text(sh, f"{g.get('team','')} 담당자 : {g.get('owner','')}", 8)

    Path(out).parent.mkdir(exist_ok=True)
    try:
        prs.save(out); saved = out
    except PermissionError:
        p = Path(out)
        saved = p.with_name(p.stem + '_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + p.suffix)
        prs.save(saved)
    return '주간회의 PPT 업데이트: ' + st, saved

base.weekly = v29_weekly

# Launch the existing GUI, but all callbacks now use patched v2.9 functions.
if __name__ == '__main__':
    base.App().mainloop()
