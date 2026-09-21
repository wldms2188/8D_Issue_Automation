"""Continuation fixes for weekly summary placement and detail title construction.

Applied last in main_enterprise_final so these user-confirmed rules win:
1) For a task that appears on multiple summary pages, continue from the LAST matching page.
   If one more row fits, insert there. If it does not fit, insert the continuation
   summary page IMMEDIATELY AFTER that last matching page.
2) When all five standard GUI inputs are present, detail title is:
   고객사_과제명_발생샘플_개발단계_발생처 이슈 발생
   The customer/project portion is the canonical combined GUI value (task_name).
   There is intentionally NO underscore before "이슈 발생".
   If any required GUI input is missing, preserve the existing title logic.
"""

import main_recovery_step14 as s14
import main_recovery_step4 as s4
import main_v310 as v310

N=v310.N

_original_prepare_new_summary_page=s14._prepare_new_summary_page
_original_detail_title_text=s4._detail_title_text


def _prepare_new_summary_page_after_last_match(prs,pages,g,template_index=None):
    """Keep legacy new-task placement, but overflow continuation goes right after the matched page."""
    if template_index is None:
        return _original_prepare_new_summary_page(prs,pages,g,template_index=None)

    # The original helper already clones the matched page, clears old rows, keeps
    # its local styling, and returns the new table/row. Only its insertion index
    # needs changing from "end of leading summary block" to "right after matched page".
    original_index_picker=s14._first_summary_section_insert_index
    try:
        s14._first_summary_section_insert_index=lambda _prs,_pages: min(len(_prs.slides),int(template_index)+1)
        return _original_prepare_new_summary_page(prs,pages,g,template_index=template_index)
    finally:
        s14._first_summary_section_insert_index=original_index_picker


def _detail_title_from_user_inputs(d,g):
    g=g or {}
    customer_project=N(g.get('task_name'))
    sample=N(g.get('sample'))
    stage=N(g.get('stage'))
    occurrence_site=N(g.get('occurrence_site'))

    # User-entered/selected values are authoritative when complete.
    if all((customer_project,sample,stage,occurrence_site)):
        return f'{customer_project}_{sample}_{stage}_{occurrence_site} 이슈 발생'

    # Missing input -> keep all previously established test/build/fallback rules.
    return _original_detail_title_text(d,g)


s14._prepare_new_summary_page=_prepare_new_summary_page_after_last_match
s4._detail_title_text=_detail_title_from_user_inputs
