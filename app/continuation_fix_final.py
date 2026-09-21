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
import main_recovery_step13 as s13
import main_v310 as v310
import project_autocomplete_final as catalog

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
    d=d or {}

    # Per-field priority confirmed by the user:
    # GUI input > existing extraction/legacy value.
    customer_project=(
        N(g.get('task_name'))
        or N(s13._customer_task(d))
    )
    sample=(
        N(g.get('sample'))
        or N(d.get('sample'))
        or N(d.get('occurrence_sample'))
    )
    stage=(
        N(g.get('stage'))
        or N(d.get('development_stage'))
        or N(d.get('occurrence_stage'))
        or N(d.get('stage'))
    )
    occurrence_site=(
        N(g.get('occurrence_site'))
        or N(d.get('_origin_occurrence_site'))
        or N(d.get('occurrence_site'))
    )

    # Build the requested full title whenever every component can be resolved.
    if all((customer_project,sample,stage,occurrence_site)):
        return f'{customer_project}_{sample}_{stage}_{occurrence_site} 이슈 발생'

    # If a component cannot be resolved at all, preserve the previous safe title logic.
    return _original_detail_title_text(d,g)


s14._prepare_new_summary_page=_prepare_new_summary_page_after_last_match
s14._update_summary_by_task=_update_summary_customer_project_then_project
s4._detail_title_text=_detail_title_from_user_inputs
