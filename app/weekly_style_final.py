"""Final weekly-PPT presentation polish.

Rules:
- Signal dot: preserve status color and center horizontally/vertically.
- Template-reference vivid blue is used only for genuinely new/changed automation content.
- Summary row through 현상 stays black; 진행사항 is blue only when it changed.
- Existing issue detail: 2D stays black. 3D~6D are blue only when their text actually changed.
- New issue detail: 2D stays black; 3D~6D are blue.
- Static template labels remain untouched.
"""
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt

import main_recovery_step14_fix2 as core
import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_v310 as v310

# Vivid royal blue matched to the user's weekly-template reference image.
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
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
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


_original_write_summary_row = s14._write_summary_row


def refined_write_summary_row(tb, row, hr, d, g):
    try:
        hm_before = s13._summary_map(tb, hr)
        old_progress = _norm(tb.cell(row, hm_before['progress']).text) if 'progress' in hm_before else ''
    except Exception:
        old_progress = ''

    _original_write_summary_row(tb, row, hr, d, g)

    try:
        hm = s13._summary_map(tb, hr)
        for key in ('task', 'issue', 'problem'):
            if key in hm:
                _color_cell(tb.cell(row, hm[key]), BLACK)
        if 'progress' in hm:
            new_progress = _norm(tb.cell(row, hm['progress']).text)
            changed = (old_progress != new_progress)
            _color_cell(tb.cell(row, hm['progress']), UPDATE_BLUE if changed else BLACK)
    except Exception:
        pass


s14._write_summary_row = refined_write_summary_row


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

        if key.startswith('2D'):
            _color_shape(sh, BLACK)
            continue

        if existing:
            changed = bool(old_text) and old_text != new_text
            _color_shape(sh, UPDATE_BLUE if changed else BLACK)
        else:
            _color_shape(sh, UPDATE_BLUE)

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
