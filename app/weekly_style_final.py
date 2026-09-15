"""Final weekly-PPT presentation polish.

Rules:
- Signal dot: preserve status color and center horizontally/vertically.
- Existing issue: every managed field is compared with its previous value; only actual changes are blue.
- Summary task/issue identity fields use semantic comparison so harmless PPT whitespace/line-break differences do not become false changes.
- Unchanged/repeated content is black regardless of D stage (2D~6D included).
- New issue: newly populated automation content is blue because there is no previous issue content.
- Static template labels remain untouched.
"""
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt
import re

import main_recovery_step14_fix2 as core
import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_v310 as v310

UPDATE_BLUE = RGBColor(0x00, 0x33, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)


def _color_text_frame(tf, color):
    if tf is None:
        return
    for p in tf.paragraphs:
        for run in p.runs:
            run.font.color.rgb = color


def _color_cell(cell, color):
    try:
        _color_text_frame(cell.text_frame, color)
    except Exception:
        pass


def _color_shape(shape, color):
    try:
        if hasattr(shape, 'text_frame'):
            _color_text_frame(shape.text_frame, color)
    except Exception:
        pass


def _norm(x):
    return '\n'.join(str(x or '').replace('\r\n', '\n').replace('\r', '\n').splitlines()).strip()


def _identity_norm(x):
    """Compare identity text independent of PPT-only wrapping/spacing.

    PowerPoint can expose a visually identical cell as spaces, tabs, NBSP or line
    breaks depending on how the template was authored. Those are formatting, not
    an issue-name change.
    """
    s = str(x or '').replace('\u00a0', ' ')
    return re.sub(r'\s+', '', s).casefold()


def _auto_text_snapshot(sl):
    out = {}
    for sh in v310.walk(sl):
        name = str(getattr(sh, 'name', '') or '')
        if name.startswith('AUTO_8D_TEXT_'):
            key = name[len('AUTO_8D_TEXT_'):].upper()
            out[key] = _norm(getattr(sh, 'text', ''))
    return out


_original_signal = core.base.signal


def centered_signal(cell, status):
    _original_signal(cell, status)
    try:
        tf = cell.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.bold = True
    except Exception:
        pass


core.base.signal = centered_signal


# Summary page: compare every managed value before overwriting it.
_original_write_summary_row = s14._write_summary_row


def refined_write_summary_row(tb, row, hr, d, g):
    try:
        hm0 = s13._summary_map(tb, hr)
        old = {k: _norm(tb.cell(row, c).text) for k, c in hm0.items()
               if k in ('task', 'issue', 'problem', 'progress')}
    except Exception:
        old = {}

    _original_write_summary_row(tb, row, hr, d, g)

    try:
        hm = s13._summary_map(tb, hr)
        for key in ('task', 'issue', 'problem', 'progress'):
            if key not in hm:
                continue
            new = _norm(tb.cell(row, hm[key]).text)
            previous = old.get(key, '')
            # Task/issue are identity fields. Ignore visual-only whitespace/wrapping
            # introduced by PowerPoint so identical content remains black.
            if key in ('task', 'issue'):
                changed = (_identity_norm(previous) != _identity_norm(new))
            else:
                changed = (previous != new)
            _color_cell(tb.cell(row, hm[key]), UPDATE_BLUE if changed else BLACK)
    except Exception:
        pass


s14._write_summary_row = refined_write_summary_row


# Detail page: no stage is special. Compare 2D~6D identically.
_original_update_detail_slide = s13._update_detail_slide


def refined_update_detail_slide(sl, d, g, mode):
    before = _auto_text_snapshot(sl)
    _original_update_detail_slide(sl, d, g, mode)
    after = _auto_text_snapshot(sl)
    existing = (s13._k(mode) in ('existing', '기존', '기존이슈'))

    for sh in v310.walk(sl):
        name = str(getattr(sh, 'name', '') or '')
        if not name.startswith('AUTO_8D_TEXT_'):
            continue
        key = name[len('AUTO_8D_TEXT_'):].upper()
        new_text = after.get(key, _norm(getattr(sh, 'text', '')))
        old_text = before.get(key, '')

        if existing:
            if key not in before:
                changed = False
            else:
                changed = (old_text != new_text)
            _color_shape(sh, UPDATE_BLUE if changed else BLACK)
        else:
            _color_shape(sh, UPDATE_BLUE if bool(new_text) else BLACK)

    labels = {'이슈기인', '발생단계'}
    label_keys = {s13._k(x) for x in labels}
    for sh in v310.walk(sl):
        if not getattr(sh, 'has_table', False):
            continue
        tb = sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                if s13._k(tb.cell(r, c).text) in label_keys and c + 1 < len(tb.columns):
                    _color_cell(tb.cell(r, c + 1), BLACK)


s13._update_detail_slide = refined_update_detail_slide
