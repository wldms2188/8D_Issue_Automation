"""Final weekly-PPT presentation polish.

Applies only to text written by the automation:
- Signal dot is centered horizontally and vertically in its table cell.
- Updated summary/detail text is blue so meeting participants can see changes.
Template labels/static text are intentionally left untouched.
"""
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt

import main_recovery_step14_fix2 as core
import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_v310 as v310

BLUE = RGBColor(0x17, 0x69, 0xAA)


def _blue_text_frame(tf):
    if tf is None:
        return
    for p in tf.paragraphs:
        for run in p.runs:
            run.font.color.rgb = BLUE


def _blue_cell(cell):
    try:
        _blue_text_frame(cell.text_frame)
    except Exception:
        pass


def _blue_shape(shape):
    try:
        if hasattr(shape, 'text_frame'):
            _blue_text_frame(shape.text_frame)
    except Exception:
        pass


# --- Signal: preserve status color, but force true cell centering. ---
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
                # Slightly larger than body text makes the status dot visually centered.
                run.font.size = Pt(10)
                run.font.bold = True
    except Exception:
        pass


core.base.signal = centered_signal


# --- Summary row: only the four values written by automation become blue.
# Signal keeps its own green/yellow/red color. ---
_original_write_summary_row = s14._write_summary_row


def blue_write_summary_row(tb, row, hr, d, g):
    _original_write_summary_row(tb, row, hr, d, g)
    try:
        hm = s13._summary_map(tb, hr)
        for key in ('task', 'issue', 'problem', 'progress'):
            if key in hm:
                _blue_cell(tb.cell(row, hm[key]))
    except Exception:
        pass


s14._write_summary_row = blue_write_summary_row


# --- Detail page: color generated D-content and metadata values written by the app.
# Static template labels (2D/3D/4D/5D/6D, Signal, 이슈기인, 발생단계...) stay original. ---
_original_update_detail_slide = s13._update_detail_slide


def blue_update_detail_slide(sl, d, g, mode):
    _original_update_detail_slide(sl, d, g, mode)

    # D-section content created by the renderer is explicitly named AUTO_8D_TEXT_*.
    for sh in v310.walk(sl):
        try:
            if str(getattr(sh, 'name', '')).startswith('AUTO_8D_TEXT_'):
                _blue_shape(sh)
        except Exception:
            pass

    # Metadata value cells: color the cell to the right of labels the automation updates.
    labels = {'이슈기인', '발생단계'}
    for sh in v310.walk(sl):
        if not getattr(sh, 'has_table', False):
            continue
        tb = sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                q = s13._k(tb.cell(r, c).text)
                if q in {s13._k(x) for x in labels} and c + 1 < len(tb.columns):
                    _blue_cell(tb.cell(r, c + 1))

    # Header/title/team-owner shapes are replaced in-place rather than generated,
    # so identify them by the values that were just written.
    candidates = [
        s13._customer_task(d),
        s13._issue_display(d),
        str(g.get('team', '') or ''),
        str(g.get('owner', '') or ''),
        str(g.get('stage', '') or ''),
    ]
    keys = [s13._k(x) for x in candidates if s13._k(x)]
    for sh in v310.walk(sl):
        if not hasattr(sh, 'text_frame'):
            continue
        text_key = s13._k(getattr(sh, 'text', ''))
        if text_key and any(k in text_key for k in keys):
            _blue_shape(sh)


s13._update_detail_slide = blue_update_detail_slide
