"""Five narrowly-scoped weekly-meeting fixes on stable baseline 667623cc.

This module is imported LAST by main_enterprise_final.py.
It deliberately leaves the stable weekly writer, attachment flow, Excel logic,
and native-section creation implementation untouched. Only helper functions
for title text, summary placement/matching, 7D positioning, and clone cleanup
are patched.
"""

import re
import datetime
import subprocess
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE

import main_recovery_step4 as s4
import main_recovery_step12 as s12
import main_recovery_step13 as s13
import main_recovery_step14 as s14
import main_recovery_step14_fix2 as core
import main_v310 as v310
import project_autocomplete_final as catalog
import main_enterprise_v3 as enterprise_v3
import ui_enterprise as ui
import final_output_polish as final_polish
import output_variant_final as output_variant

N = v310.N


# ---------------------------------------------------------------------------
# 1) Detail title:
#    고객사_과제명_발생샘플 샘플_개발단계_발생처 이슈 발생
# ---------------------------------------------------------------------------
_original_detail_title_text = s4._detail_title_text


def _title_customer_project(d, g):
    selected = N((g or {}).get("task_name")) or N(s13._customer_task(d or {}))
    customer, project = catalog.split_customer_task(selected)
    if not customer:
        customer = N((d or {}).get("customer"))
        project = N(project or selected)
    return N(customer), N(project)


def _sample_with_suffix(value):
    value = N(value)
    if not value:
        return ""
    if re.search(r"(?:샘플|sample)\s*$", value, re.I):
        return value
    return value + " 샘플"


def _detail_title_user_format(d, g):
    d = d or {}
    g = g or {}
    customer, project = _title_customer_project(d, g)
    sample = N(g.get("sample") or d.get("sample") or d.get("occurrence_sample"))
    stage = N(
        g.get("stage")
        or d.get("development_stage")
        or d.get("occurrence_stage")
        or d.get("stage")
    )
    site = N(
        g.get("occurrence_site")
        or d.get("_origin_occurrence_site")
        or d.get("occurrence_site")
    )

    if all((customer, project, sample, stage, site)):
        return f"{customer}_{project}_{_sample_with_suffix(sample)}_{stage}_{site} 이슈 발생"

    return _original_detail_title_text(d, g)


s4._detail_title_text = _detail_title_user_format


# ---------------------------------------------------------------------------
# 3) Keep the native 7D circle + 수평전개 unit below the dynamic 6D region.
#    Recompute an absolute target each run, so repeated updates do not drift.
# ---------------------------------------------------------------------------
_original_update_page2 = s12._update_page2_step12


def _adjust_7d_after_6d(sl, d):
    try:
        imgs = (d or {}).get("_section_images", {}) or {}
        zones, _, _ = s12.step11._adaptive_cascade_layout(d or {}, imgs)
        z6 = zones.get("6D")
        if not z6:
            return

        marker, parent = v310.find_marker(sl, "7D")
        if marker is None:
            return
        unit = parent if parent is not None else marker
        x7, _, _, h7 = v310.box(unit)

        try:
            slide_h = float(
                sl.part.package.presentation_part.presentation.slide_height
            ) / v310.EMU
        except Exception:
            slide_h = 7.5

        desired_y = float(z6["y"]) + float(z6["h"]) + 0.28
        max_y = max(0.10, slide_h - float(h7) / v310.EMU - 0.06)
        desired_y = min(max(0.10, desired_y), max_y)
        v310.move_marker_unit(sl, "7D", float(x7) / v310.EMU, desired_y)
    except Exception:
        pass


def _update_page2_with_7d(sl, d, g, mode):
    _original_update_page2(sl, d, g, mode)
    _adjust_7d_after_6d(sl, d)


s12._update_page2_step12 = _update_page2_with_7d


# ---------------------------------------------------------------------------
# 5) When a summary/detail page is CLONED, strip copied image/media objects.
#    Tables, text, rectangles/frames, lines and other template structure stay.
# ---------------------------------------------------------------------------
_VISUAL_TYPES = {
    x
    for x in (
        getattr(MSO_SHAPE_TYPE, "PICTURE", None),
        getattr(MSO_SHAPE_TYPE, "LINKED_PICTURE", None),
        getattr(MSO_SHAPE_TYPE, "CHART", None),
        getattr(MSO_SHAPE_TYPE, "MEDIA", None),
        getattr(MSO_SHAPE_TYPE, "OLE_OBJECT", None),
    )
    if x is not None
}


def _remove_shape(sh):
    try:
        el = sh._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            return True
    except Exception:
        pass
    return False


def _clean_group_visuals(group):
    for child in list(getattr(group, "shapes", ())):
        st = getattr(child, "shape_type", None)
        if st in _VISUAL_TYPES:
            _remove_shape(child)
        elif st == MSO_SHAPE_TYPE.GROUP:
            _clean_group_visuals(child)


def _remove_copied_visuals(sl):
    for sh in list(sl.shapes):
        if getattr(sh, "has_table", False):
            continue
        st = getattr(sh, "shape_type", None)
        if st in _VISUAL_TYPES:
            _remove_shape(sh)
        elif st == MSO_SHAPE_TYPE.GROUP:
            _clean_group_visuals(sh)


_original_prepare_new_summary_page = s14._prepare_new_summary_page


def _slide_id_value(sl):
    try:
        return int(sl.slide_id)
    except Exception:
        return None


def _index_by_slide_id(prs, slide_id):
    if slide_id is None:
        return None
    for i, sl in enumerate(prs.slides):
        if _slide_id_value(sl) == slide_id:
            return i
    return None


def _move_slide_by_id(prs, slide_id, index):
    current = _index_by_slide_id(prs, slide_id)
    if current is None:
        return None
    sld_id_lst = prs.slides._sldIdLst
    node = sld_id_lst[current]
    sld_id_lst.remove(node)
    index = max(0, min(int(index), len(sld_id_lst)))
    sld_id_lst.insert(index, node)
    return index


def _find_summary_table_on_slide(sl, allow_partial=False):
    """Rediscover the summary table on a cloned page.

    Generic new-summary pages stay strict. A page already confirmed by an exact
    task-name hit may use a slightly older header variant, so task + two known
    summary columns is enough there.
    """
    for sh in v310.walk(sl):
        if not getattr(sh, "has_table", False):
            continue
        tb = sh.table
        for hr in range(min(8, len(tb.rows))):
            hm = _summary_header_map(tb, hr)
            if _is_real_summary_header(hm):
                return tb, hr
            if allow_partial and "task" in hm:
                data_fields = sum(
                    1 for k in ("issue", "problem", "progress", "signal") if k in hm
                )
                if data_fields >= 2:
                    return tb, hr
    return None, None


def _prepare_new_summary_page_stable(prs, pages, g, template_index=None):
    """Clone a summary page and rediscover its table with the SAME flexible rule.

    The legacy helper required an exact 'Signal' header after cloning, which
    made valid weekly templates fail even though the source summary page had
    already been found successfully.
    """
    if not pages:
        raise ValueError(
            "주간회의 PPT에서 과제명 요약 양식 페이지를 찾지 못했습니다."
        )

    # For an overflow continuation, clone the actual last matching page and
    # insert immediately after it.  For a brand-new project, preserve the
    # stable leading-summary-block placement.
    if template_index is None:
        template_index = pages[0][0]
        try:
            desired_index = s14._first_summary_section_insert_index(prs, pages)
        except Exception:
            desired_index = template_index + 1
    else:
        template_index = int(template_index)
        desired_index = template_index + 1

    if template_index < 0 or template_index >= len(prs.slides):
        raise ValueError("복제할 주간회의 요약 양식 페이지 위치가 올바르지 않습니다.")

    source_id = _slide_id_value(prs.slides[template_index])

    # Clone directly instead of calling the legacy helper, because that helper
    # is exactly where the strict '과제명 + Signal' re-scan raises the error.
    s13._clone_slide_with_rels(prs, template_index)

    created_id = None
    for sl in reversed(list(prs.slides)):
        sid = _slide_id_value(sl)
        if sid != source_id:
            # The clone is appended at the end by _clone_slide_with_rels.
            created_id = sid
            break

    if created_id is None:
        raise ValueError("주간회의 요약 페이지 복제본을 확인하지 못했습니다.")

    moved = _move_slide_by_id(prs, created_id, desired_index)
    if moved is None:
        raise ValueError("복제한 주간회의 요약 페이지의 위치를 조정하지 못했습니다.")

    final_index = moved
    sl = prs.slides[final_index]
    tb, hr = _find_summary_table_on_slide(
        sl, allow_partial=(template_index is not None)
    )
    if tb is None:
        raise ValueError(
            "복제한 주간회의 요약 페이지에서 과제명 표를 다시 찾지 못했습니다."
        )

    row = s14._clear_summary_data(tb, hr)

    # Keep the shared page style and only replace the team token when present.
    team = N((g or {}).get("team"))
    if team:
        for sh in v310.walk(sl):
            if not hasattr(sh, "text_frame"):
                continue
            old = N(getattr(sh, "text", ""))
            if "000팀" in old:
                try:
                    sh.text = old.replace("000팀", team + "팀")
                except Exception:
                    pass

    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass

    return final_index, tb, hr, row


s14._prepare_new_summary_page = _prepare_new_summary_page_stable



_original_static_cloned_detail_shape = core._static_cloned_detail_shape


def _static_cloned_detail_shape_preserve_layout(sh):
    """Never delete native D-marker/title units from a copied detail template."""
    try:
        texts = core._shape_texts(sh)
    except Exception:
        texts = [N(getattr(sh, "text", ""))]
    joined = "".join(s13._k(t) for t in texts if N(t))

    # Real templates often keep e.g. "2D 현상" or "5D 개선대책" in one shape
    # or group. The older exact-label test could mistakenly delete these.
    marker_words = (
        ("2d", ("현상", "문제")),
        ("3d", ("임시", "조치", "대응")),
        ("4d", ("원인", "발생", "유출", "시스템")),
        ("5d", ("개선", "대책")),
        ("6d", ("효과", "검증", "유효")),
        ("7d", ("수평", "전개")),
    )
    for marker, words in marker_words:
        if marker in joined and any(word in joined for word in words):
            return True

    # Short marker-only groups/circles must also survive.
    if any(marker in joined for marker in ("2d","3d","4d","5d","6d","7d")) and len(joined) <= 28:
        return True

    return _original_static_cloned_detail_shape(sh)


core._static_cloned_detail_shape = _static_cloned_detail_shape_preserve_layout

_original_clone_detail_shell = core._clone_detail_shell


def _clone_detail_shell_clean(prs, d, insert_at, matched_section=None):
    result = _original_clone_detail_shell(prs, d, insert_at, matched_section)
    sl, template_index, section = result

    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass

    # Do not allow an "attachments only" result. The cloned page itself must
    # still contain a recognizable 2D~6D detail structure before continuing.
    if not core._is_detail_like(sl):
        raise RuntimeError(
            "복제한 상세 양식의 2D~6D 구조가 유지되지 않았습니다. "
            f"(양식 원본 slide {template_index + 1})"
        )

    # If there is no safe native detail section, this newly created detail must
    # become the first slide of a NEW section. Record the request here rather
    # than relying only on the earlier section-selection branch.
    if section is None:
        global _pending_user_native_section
        try:
            slide_id = int(sl.slide_id)
        except Exception:
            slide_id = None
        _pending_user_native_section = {
            "name": N(core._new_section_name(d)) or "신규 과제",
            "slide_index": int(insert_at),
            "slide_id": slide_id,
        }

    return result


core._clone_detail_shell = _clone_detail_shell_clean


# ---------------------------------------------------------------------------
# 4) Native-section resolution:
#    1. customer + project strict
#    2. project-only strict
#    3. similar section only as a USER-CONFIRMABLE suggestion
# ---------------------------------------------------------------------------
def _catalog_separator_alias(value):
    """Resolve catalog names ignoring separators.

    Example: MBAG_EB565M == MBAGEB565M == MBAG EB565M.
    """
    q = s13._k(value)
    if not q:
        return ""
    for values in catalog.PROJECTS.values():
        for candidate in values:
            if s13._k(candidate) == q:
                return N(candidate)
    return ""


def _identity_parts_from_label(value, customer_hint=""):
    """Return separator-insensitive (full, project, customer) identity.

    Catalog aliases win because the first underscore in catalog names is semantic
    customer/project metadata, while its visual presence is NOT required in the
    weekly PPT label.
    """
    raw = N(value)
    q = s13._k(raw)
    hint = s13._k(customer_hint)
    if not q:
        return "", "", hint

    canonical = _catalog_separator_alias(raw)
    if canonical:
        c0, p0 = catalog.split_customer_task(canonical)
        ck = s13._k(c0)
        pk = s13._k(p0 or canonical)
        fk = s13._k(canonical)
        return fk, pk, ck

    # For non-catalog labels, use the known customer only as a prefix hint.
    # This makes GM_MBAG / GM MBAG / GMMBAG identical without treating every
    # underscore inside a project name as a mandatory split point.
    if hint:
        if q.startswith(hint) and len(q) > len(hint):
            return q, q[len(hint):], hint
        return hint + q, q, hint

    # Last-resort canonical A_B behavior when no customer hint exists.
    c0, p0 = catalog.split_customer_task(raw)
    if c0 and p0:
        return q, s13._k(p0), s13._k(c0)
    return q, q, ""


def _project_parts(d):
    d = d or {}
    selected = N(d.get("task_name")) or N(s13._customer_task(d))
    return _identity_parts_from_label(selected, N(d.get("customer")))


def _project_key_for_match(value, customer_key=""):
    """Project identity for matching only; separators never affect equality."""
    q = s13._k(value)
    if not q:
        return ""

    canonical = _catalog_separator_alias(value)
    if canonical:
        _c0, p0 = catalog.split_customer_task(canonical)
        return s13._k(p0 or canonical)

    ck = s13._k(customer_key)
    if ck and q.startswith(ck) and len(q) > len(ck):
        return q[len(ck):]
    return q


def _section_match_level(name, d):
    q = s13._k(name)
    full, project, customer = _project_parts(d)
    if not q:
        return 0

    # Exact customer+project always wins. Visual separators are ignored.
    if full and q == full:
        return 3

    # Then exact project. This fallback is intentionally independent from
    # customer matching: even if customer metadata is different, "GM MBAG"
    # still exactly matches project "MBAG" after separator normalization.
    section_project = _project_key_for_match(name, customer)
    if project and (
        section_project == project
        or (len(project) >= 3 and q.endswith(project))
    ):
        return 2

    # Similarity is suggestion-only and can never override levels 3/2.
    keys = [x for x in (full, project) if x]
    candidates = [q]
    if section_project and section_project != q:
        candidates.append(section_project)
    for key in keys:
        for candidate in candidates:
            if min(len(key), len(candidate)) >= 3:
                if key in candidate or candidate in key:
                    return 1
    return 0


def _matching_native_section_strict(prs, d):
    for level in (3, 2):
        matches = [
            sec
            for sec in core._native_sections(prs)
            if _section_match_level(sec.get("name"), d) == level
        ]
        if matches:
            matches.sort(
                key=lambda x: len(x.get("indices", [])), reverse=True
            )
            return matches[0]
    return None


def _section_resolution_strict(prs, d):
    try:
        s13._clear_slide_text_cache()
    except Exception:
        pass

    sections = core._native_sections(prs)

    for level, reason in (
        (3, "고객사와 과제명이 모두 일치하는 구역을 찾았습니다."),
        (2, "과제명이 정확히 일치하는 구역을 찾았습니다."),
    ):
        matches = [
            sec
            for sec in sections
            if _section_match_level(sec.get("name"), d) == level
        ]
        if matches:
            matches.sort(
                key=lambda x: len(x.get("indices", [])), reverse=True
            )
            sec = matches[0]
            return {
                "mode": "exact",
                "name": N(sec.get("name")),
                "reason": reason,
                "section": sec,
            }

    similar = [
        sec
        for sec in sections
        if _section_match_level(sec.get("name"), d) == 1
    ]
    if similar:
        similar.sort(
            key=lambda x: len(x.get("indices", [])), reverse=True
        )
        sec = similar[0]
        return {
            "mode": "suggest",
            "name": N(sec.get("name")),
            "reason": "과제명과 유사한 기존 구역이 확인되었습니다.",
            "section": sec,
        }

    return {
        "mode": "new",
        "name": N(s13._customer_task(d))
        or N((d or {}).get("task_name"))
        or N((d or {}).get("customer"))
        or "신규 과제",
        "reason": (
            "일치하거나 유사한 기존 구역을 찾지 못했습니다. "
            "신규 구역을 생성합니다."
        ),
        "section": None,
    }


core._project_parts = _project_parts
core._section_match_level = _section_match_level
core._matching_native_section = _matching_native_section_strict
core.section_resolution = _section_resolution_strict


# ---------------------------------------------------------------------------
# Mixed summary/detail native-section recovery.
# Some legacy weekly decks keep summary pages and 2D~6D detail pages inside
# one PowerPoint section.  That section is NOT a safe target for a newly cloned
# detail page.  Treat it as template/source material only and create a fresh
# native section for the new detail + attachments.
# ---------------------------------------------------------------------------
_original_selected_section = core._selected_section
_original_template_detail_index = core._template_detail_index


def _section_contains_summary(prs, section):
    if not section:
        return False
    summary_indices = set(core._summary_indices(prs))
    return any(i in summary_indices for i in section.get("indices", []))


def _selected_section_without_mixed_summary(prs, d, g):
    section = _original_selected_section(prs, d, g)
    if section and _section_contains_summary(prs, section):
        # This is the legacy layout that caused detail creation to disappear:
        # summary and detail share one native section.  Do not append the clone
        # to that mixed section.  The active weekly writer reads this flag
        # immediately after _selected_section() and creates a new section.
        g["_weekly_create_new_section"] = "1"
        g["_weekly_mixed_section_recovered"] = N(section.get("name"))
        return None
    return section


def _detail_template_score(sl, d):
    """Score a reusable detail template independently from native sections."""
    if not core._is_detail_like(sl):
        return -1

    text = s13._slide_text(sl)
    q = s13._k(text)
    score = 0

    # Prefer the complete shared 8D detail skeleton.
    for token in ("2d", "3d", "4d", "5d", "6d"):
        if token in q:
            score += 25
    if "7d" in q or "수평전개" in q:
        score += 10
    for token in ("signal", "이슈기인", "발생단계"):
        if token in q:
            score += 12

    # Same-project detail is the closest visual/template source, but section
    # membership itself never blocks template reuse.
    full, project, _customer = _project_parts(d)
    if full and full in q:
        score += 80
    elif project and project in q:
        score += 55

    return score


def _template_detail_index_any_section(prs, d):
    """Find the best 2D~6D detail shell across the entire deck.

    Summary-page section membership is deliberately ignored.  This recovers
    legacy decks where valid detail templates live in the same section as the
    summary pages.
    """
    summaries = set(core._summary_indices(prs))
    candidates = []
    for i, sl in enumerate(prs.slides):
        if i in summaries:
            continue
        score = _detail_template_score(sl, d)
        if score >= 0:
            candidates.append((score, i))

    if not candidates:
        # Preserve the stable baseline error/fallback behavior if no reusable
        # detail-like slide exists anywhere.
        return _original_template_detail_index(prs, d)

    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][1]


core._selected_section = _selected_section_without_mixed_summary
core._template_detail_index = _template_detail_index_any_section


# ---------------------------------------------------------------------------
# Real weekly-summary table compatibility.
# Existing decks may merge the project cell vertically or leave the project cell
# blank on continuation issue rows.  Also, some continuation pages use slightly
# different header wording.  Detect those pages and treat blank task rows as
# belonging to the previous visible project until the next project label.
# ---------------------------------------------------------------------------
_original_summary_pages_user = s14._summary_pages


def _summary_header_map(tb, hr):
    """Map the REAL weekly-summary header without accepting unrelated tables."""
    hm = {}
    for col in range(len(tb.columns)):
        q = v310.C(tb.cell(hr, col).text)
        if not q:
            continue
        if "과제명" in q:
            hm["task"] = col
        elif q == "이슈" or "이슈명" in q:
            hm["issue"] = col
        elif "현상" in q or "문제" in q:
            hm["problem"] = col
        elif "진행사항" in q or "진행현황" in q or "진행내용" in q:
            hm["progress"] = col
        elif "signal" in q:
            # Accept Signal / Signal상태 / Signal 등, but only as one field in
            # a structurally complete summary table.
            hm["signal"] = col
    return hm


def _is_real_summary_header(hm):
    """A summary table must have 과제명 plus most of the known summary columns."""
    if "task" not in hm:
        return False
    data_fields = sum(
        1 for key in ("issue", "problem", "progress", "signal") if key in hm
    )
    # Requiring 3 prevents unrelated '과제명 + 이슈명' tables from being used
    # as the clone template while still allowing one optional column to differ.
    return data_fields >= 3


# Make every downstream summary operation use the same header interpretation.
s13._summary_map = _summary_header_map


def _summary_pages_flexible(prs):
    pages = []
    for si, sl in enumerate(prs.slides):
        found = None
        for sh in v310.walk(sl):
            if not getattr(sh, "has_table", False):
                continue
            tb = sh.table
            for hr in range(min(8, len(tb.rows))):
                hm = _summary_header_map(tb, hr)
                if _is_real_summary_header(hm):
                    found = (si, tb, hr)
                    break
            if found:
                break
        if found:
            pages.append(found)

    # Do NOT broaden to a random weaker table here. If this strict structural
    # scan finds nothing, let the caller fail instead of cloning a wrong format.
    return pages


s14._summary_pages = _summary_pages_flexible


def _summary_row_has_payload(tb, r, hm):
    """True when a continuation row actually contains issue-summary content."""
    for key in ("issue", "problem", "progress", "signal"):
        if key not in hm:
            continue
        if N(s13._row_text(tb, r, hm[key])):
            return True
    return False


def _summary_display_for_rows(tb, hr, hm, rows):
    """Recover the visible project spelling for a matched project block."""
    if "task" not in hm or not rows:
        return ""
    start = min(rows)
    # The task can be the merge origin above a continuation row.
    for r in range(start, hr, -1):
        raw = N(s13._row_text(tb, r, hm["task"]))
        if raw:
            return raw
    return ""


# ---------------------------------------------------------------------------
# 2 + 4) Summary placement/matching across ALL summary pages.
#    1. exact customer+project
#    2. exact project
#    3. only if the chosen native section was SIMILAR and the user selected
#       "해당 구역 업데이트", match that section name against real summary labels.
#       If absent, use the normal new-summary path with customer+project.
# ---------------------------------------------------------------------------
def _summary_match_key(value):
    """Comparison-only key: separators such as _, spaces, -, / are ignored."""
    return s13._k(value)


def _summary_project_key_from_row(raw, project_key, customer_key=""):
    """Project-only exact fallback, independent from customer matching.

    Step 2 must still work even when step 1 (customer+project) fails.
    Separators are ignored, and a summary label that includes a customer prefix
    matches when its normalized suffix is the exact project key.
    """
    if not project_key:
        return ""

    q = _summary_match_key(raw)
    if not q:
        return ""

    # Pure project label.
    if q == project_key:
        return project_key

    # Catalog-aware identity (e.g. MBAG_EB565M / MBAGEB565M).
    row_project = _project_key_for_match(raw, customer_key)
    if row_project == project_key:
        return project_key

    # Customer-independent fallback: "GM MBAG" -> "mbag".
    # This is exact suffix matching, not fuzzy containment.
    if len(project_key) >= 3 and q.endswith(project_key):
        return project_key

    return ""


def _summary_direct_keys(d, g=None):
    """Exact visible-name candidates, before any customer/project splitting."""
    values = [
        N((g or {}).get("task_name")),
        N((d or {}).get("task_name")),
        N(s13._customer_task(d or {})),
    ]
    customer = N((d or {}).get("customer"))
    task = N((d or {}).get("task_name"))
    if customer and task:
        values.extend([
            customer + "_" + task,
            customer + " " + task,
        ])

    # Include catalog aliases only as equivalent spellings, not as a prerequisite.
    for value in list(values):
        alias = _catalog_separator_alias(value)
        if alias:
            values.append(alias)

    return {s13._k(v) for v in values if s13._k(v)}


def _summary_hits(prs, d, g=None):
    # Keep the strict detector for generic template creation.
    pages = s14._summary_pages(prs)

    direct_keys = _summary_direct_keys(d, g)
    selected = N((g or {}).get("task_name")) or N((d or {}).get("task_name"))
    if selected:
        _full_key, project_key, customer_key = _identity_parts_from_label(
            selected, N((d or {}).get("customer"))
        )
    else:
        _full_key, project_key, customer_key = _project_parts(d)

    direct_hits = []
    project_hits = []

    # IMPORTANT: finding an EXISTING project is intentionally simpler than
    # choosing a generic summary template. Scan every non-detail slide for a
    # table that has 과제명 plus at least two normal summary columns.
    for si, sl in enumerate(prs.slides):
        if core._is_detail_like(sl):
            continue

        found_tables = []
        for sh in v310.walk(sl):
            if not getattr(sh, "has_table", False):
                continue
            tb = sh.table
            for hr in range(min(8, len(tb.rows))):
                hm = _summary_header_map(tb, hr)
                if "task" not in hm:
                    continue
                data_fields = sum(
                    1 for k in ("issue", "problem", "progress", "signal") if k in hm
                )
                if data_fields >= 2:
                    found_tables.append((tb, hr, hm))
                    break

        for tb, hr, hm in found_tables:
            blocks = []
            current = None
            for r in range(hr + 1, len(tb.rows)):
                raw = N(s13._row_text(tb, r, hm["task"]))
                if raw:
                    if current is not None:
                        blocks.append(current)
                    current = {"label": raw, "rows": [r]}
                elif current is not None and _summary_row_has_payload(tb, r, hm):
                    current["rows"].append(r)
            if current is not None:
                blocks.append(current)

            direct_rows = []
            project_rows = []
            for block in blocks:
                raw = block["label"]
                rows = block["rows"]
                raw_key = s13._k(raw)

                # Old/simple behavior first: visible task text equality after
                # removing separators. No catalog/customer split is required.
                if raw_key and raw_key in direct_keys:
                    direct_rows.extend(rows)
                    continue

                if (
                    project_key
                    and _summary_project_key_from_row(
                        raw, project_key, customer_key
                    ) == project_key
                ):
                    project_rows.extend(rows)

            if direct_rows:
                direct_hits.append((si, tb, hr, hm, direct_rows))
            if project_rows:
                project_hits.append((si, tb, hr, hm, project_rows))

    # Exact direct text anywhere wins globally. Then exact project-only.
    return pages, (direct_hits if direct_hits else project_hits)

def _confirmed_section_summary_hits(pages, d, g):
    name = N((g or {}).get("_weekly_section_override_name"))
    if not name:
        return []

    # Level 1 means this section was only available through explicit user
    # confirmation. Exact section matches never need this fallback.
    if _section_match_level(name, d) != 1:
        return []

    _full, _project, customer = _project_parts(d)
    confirmed_key = s13._k(name)
    confirmed_project = _project_key_for_match(name, customer)
    if not confirmed_key:
        return []

    hits = []
    for si, tb, hr in pages:
        hm = s13._summary_map(tb, hr)
        if "task" not in hm:
            continue
        rows = []
        for r in range(hr + 1, len(tb.rows)):
            raw = s13._row_text(tb, r, hm["task"])
            raw_key = s13._k(raw)
            raw_project = _project_key_for_match(raw, customer)

            # Exact same confirmed label after separator normalization.
            if raw_key and raw_key == confirmed_key:
                rows.append(r)
                continue

            # Or exact same project identity after the same normalization.
            if (
                confirmed_project
                and raw_project
                and raw_project == confirmed_project
            ):
                rows.append(r)
                continue

            # Similarity is allowed only here because the user already selected
            # "해당 구역 업데이트".
            pairs = (
                (raw_key, confirmed_key),
                (raw_project, confirmed_project),
            )
            if any(
                a and b and min(len(a), len(b)) >= 4
                and (a in b or b in a)
                for a, b in pairs
            ):
                rows.append(r)

        if rows:
            hits.append((si, tb, hr, hm, rows))
    return hits


def _existing_summary_display(selected, task_hits):
    if not task_hits:
        return N(selected)
    si, tb, hr, hm, rows = sorted(
        task_hits, key=lambda x: x[0]
    )[-1]
    # In real weekly tables the last issue row often has an empty/spanned task
    # cell. Recover the visible label from the project block instead of falling
    # back to the user's input spelling.
    raw = _summary_display_for_rows(tb, hr, hm, rows)
    return raw or N(selected)


def _update_summary_exact_then_confirmed(prs, d, g, mode):
    pages = s14._summary_pages(prs)
    if not pages:
        raise ValueError(
            "주간회의 PPT에서 과제명/Signal 요약 양식 페이지를 찾지 못했습니다."
        )

    selected = N((g or {}).get("task_name")) or N(s13._customer_task(d))
    issue = s13._issue_display(d)

    _, task_hits = _summary_hits(prs, d, g)
    if not task_hits:
        task_hits = _confirmed_section_summary_hits(pages, d, g)

    display = _existing_summary_display(selected, task_hits)
    original_customer_task = s13._customer_task

    try:
        if display:
            s13._customer_task = lambda _d: display

        if mode == "existing" and task_hits:
            best = None
            for si, tb, hr, hm, rows in task_hits:
                for r in rows:
                    sc = s13._row_match_score(
                        tb, r, hm, display, issue
                    )
                    if best is None or sc > best[0]:
                        best = (sc, si, tb, hr, r)
            if best and best[0] >= 140:
                _, si, tb, hr, row = best
                s14._write_summary_row(tb, row, hr, d, g)
                return si, row, "기존 행 업데이트"

        if task_hits:
            # Continue from the LAST page where this task appears.
            si, tb, hr, hm, rows = sorted(
                task_hits, key=lambda x: x[0]
            )[-1]
            after = max(rows)
            if s14._summary_row_insert_fits(prs, si, tb, after):
                row = s14._insert_row_after(tb, after)
                s14._write_summary_row(tb, row, hr, d, g)
                return (
                    si,
                    row,
                    "확정 과제의 마지막 요약 행 바로 아래 삽입",
                )

            nsi, ntb, nhr, nrow = s14._prepare_new_summary_page(
                prs, pages, g, template_index=si
            )
            s14._write_summary_row(ntb, nrow, nhr, d, g)
            return (
                nsi,
                nrow,
                "확정 과제 요약 공간 초과로 바로 다음 페이지에 추가",
            )

        # No exact or user-confirmed similar summary label: do not force fuzzy.
        if selected:
            s13._customer_task = lambda _d: selected
        si, tb, hr, row = s14._prepare_new_summary_page(prs, pages, g)
        s14._write_summary_row(tb, row, hr, d, g)
        return (
            si,
            row,
            "고객사/과제명 매칭 없음: 신규 요약 페이지 생성",
        )
    finally:
        s13._customer_task = original_customer_task


s14._update_summary_by_task = _update_summary_exact_then_confirmed


def _row_match_score_separator_agnostic(tb, r, hm, task, issue):
    """Existing-summary row scoring with the same separator-insensitive identity."""
    task_key = s13._k(task)
    row_task = (
        s13._k(s13._row_text(tb, r, hm["task"]))
        if "task" in hm else ""
    )
    issue_key = s13._k(issue)
    row_issue = (
        s13._k(s13._row_text(tb, r, hm["issue"]))
        if "issue" in hm else ""
    )

    score = 0
    if task_key and row_task == task_key:
        score += 100
    elif task_key and row_task and (task_key in row_task or row_task in task_key):
        score += 60

    if issue_key and row_issue == issue_key:
        score += 100
    elif issue_key and row_issue and (
        issue_key in row_issue or row_issue in issue_key
    ):
        score += 70
    return score


# Loaded last: override the older project_key-based scorer, which treated
# MBAG_EB565M and MBAGEB565M differently.
s13._row_match_score = _row_match_score_separator_agnostic


# ---------------------------------------------------------------------------
# Detail-page placement: when the project already exists, add after its LAST
# detail page instead of falling through to the end of the whole deck.
# ---------------------------------------------------------------------------
_original_new_detail_position = core._new_detail_position


def _new_detail_position_after_last_project(prs, d):
    summaries = core._summary_indices(prs)
    full, project, _customer = _project_parts(d)
    hits = []
    for i, sl in enumerate(prs.slides):
        if i in summaries or not core._is_detail_like(sl):
            continue
        q = s13._k(s13._slide_text(sl))
        if (full and full in q) or (project and project in q):
            hits.append(i)
    if hits:
        return max(hits) + 1
    return _original_new_detail_position(prs, d)


core._new_detail_position = _new_detail_position_after_last_project


def _find_existing_detail_exact_issue(prs, d):
    """Do not reuse another issue's detail page just because the project matches."""
    summaries = set(core._summary_indices(prs))
    full, project, _customer = _project_parts(d)
    issue = s13._k(s13._issue_display(d))
    if not issue:
        return None

    hits = []
    for i, sl in enumerate(prs.slides):
        if i in summaries or not core._is_detail_like(sl):
            continue
        q = s13._k(s13._slide_text(sl))

        project_ok = (
            (full and full in q)
            or (project and project in q)
        )
        if not project_ok:
            continue

        # Existing-detail update is allowed only when the actual issue identity
        # is present. Otherwise a new detail page must be cloned.
        if issue in q:
            hits.append(i)

    return max(hits) if hits else None


core._find_existing_detail = _find_existing_detail_exact_issue


# ---------------------------------------------------------------------------
# Attachment recovery:
# - keep the existing de-duplication path first;
# - if it reports zero although the source really has attachment slides and no
#   current issue attachment exists in the output, retry the stable direct COM
#   inserter once;
# - if de-duplication raises, retry direct insertion once.
# This also preserves the baseline behavior that replaces AUTO_8D_ATTACH_<key>
# pages for the same issue rather than accumulating repeated copies.
# ---------------------------------------------------------------------------
_original_attachment_safe = core._append_8d_attachments_safe


def _count_source_attachment_slides(path):
    try:
        from pptx import Presentation
        return max(0, len(Presentation(path).slides) - 1)
    except Exception:
        return 0


def _has_current_auto_attachment(out_path, d):
    try:
        from pptx import Presentation
        prefix = "AUTO_8D_ATTACH_" + core._attachment_key(d) + "_"
        return any(
            str(getattr(sl, "name", "") or "").startswith(prefix)
            for sl in Presentation(out_path).slides
        )
    except Exception:
        return False


def _direct_attachment_inserter():
    try:
        import attachment_dedupe_final as dedupe
        return getattr(dedupe, "_original_append", None)
    except Exception:
        return None


def _append_8d_attachments_safe_recover(out_path, src8d_path, detail_index, d):
    expected = _count_source_attachment_slides(src8d_path)
    if expected <= 0:
        return 0, ""

    try:
        added, err = _original_attachment_safe(
            out_path, src8d_path, detail_index, d
        )
    except Exception as exc:
        added, err = 0, str(exc)

    if added or _has_current_auto_attachment(out_path, d):
        return added, err

    direct = _direct_attachment_inserter()
    if callable(direct):
        try:
            added = direct(out_path, src8d_path, detail_index, d)
            return added, ""
        except Exception as exc:
            direct_err = str(exc)
            if err:
                return 0, err + " / direct retry: " + direct_err
            return 0, direct_err

    return added, err or "8D 유첨 페이지가 확인되었지만 출력 PPT에 추가되지 않았습니다."


core._append_8d_attachments_safe = _append_8d_attachments_safe_recover


# ---------------------------------------------------------------------------
# New-project section creation:
# python-pptx cannot create a real PowerPoint native section by writing a fake
# p:sectionLst.  Preserve the freshly cloned detail slide in the pptx first,
# then create the real section through PowerPoint COM after the file is saved.
# ---------------------------------------------------------------------------
_pending_user_native_section = None


def _create_native_section_pending(prs, name, slide_index):
    global _pending_user_native_section
    idx = int(slide_index)
    slide_id = None
    try:
        slide_id = int(prs.slides[idx].slide_id)
    except Exception:
        pass

    _pending_user_native_section = {
        "name": N(name) or "신규 과제",
        "slide_index": idx,
        # Numeric position can change after save/attachment insertion.  The
        # PowerPoint slide-id is the stable identity of the cloned detail page.
        "slide_id": slide_id,
    }
    return {
        "name": N(name) or "신규 과제",
        "element": None,
        "slide_ids": [slide_id] if slide_id is not None else [],
        "indices": [idx],
        "_pending_com": True,
    }


def _ps_quote(value):
    return str(value).replace("'", "''")


def _create_native_section_com_saved(ppt_path, name, slide_index):
    """Create a real PowerPoint section before the already-saved detail slide."""
    path_q = _ps_quote(Path(ppt_path).resolve())
    name_q = _ps_quote(N(name) or "신규 과제")
    slide_no = int(slide_index) + 1
    script = f"""
$ErrorActionPreference = 'Stop'
$ppt = $null
$pres = $null
try {{
    $ppt = New-Object -ComObject PowerPoint.Application
    try {{ $ppt.DisplayAlerts = 1 }} catch {{}}
    $pres = $ppt.Presentations.Open('{path_q}', 0, 0, 0)
    if ({slide_no} -lt 1 -or {slide_no} -gt $pres.Slides.Count) {{
        throw '상세페이지 위치가 PowerPoint 범위를 벗어났습니다.'
    }}
    [void]$pres.SectionProperties.AddBeforeSlide({slide_no}, '{name_q}')
    $pres.Save()
}}
finally {{
    if ($pres -ne $null) {{ try {{ $pres.Close() }} catch {{}} }}
    if ($ppt -ne $null) {{ try {{ $ppt.Quit() }} catch {{}} }}
}}
"""
    try:
        proc = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Sta",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            capture_output=True,
            timeout=45,
        )
    except FileNotFoundError as exc:
        raise RuntimeError(
            "Windows PowerShell을 찾지 못해 신규 PowerPoint 구역을 만들지 못했습니다."
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "PowerPoint 신규 구역 생성이 45초를 초과했습니다."
        ) from exc

    if proc.returncode != 0:
        raw = (proc.stderr or proc.stdout or b"")[-1200:]
        try:
            detail = raw.decode("cp949", errors="replace")
        except Exception:
            detail = raw.decode("utf-8", errors="replace")
        raise RuntimeError("PowerPoint 신규 구역 생성 실패: " + detail.strip())
    return True


def _generated_detail_shape_count(sl):
    count = 0
    for sh in v310.walk(sl):
        name = str(getattr(sh, "name", "") or "")
        if name.startswith("AUTO_8D_TEXT_"):
            count += 1
    return count


def _filled_detail_slide(sl):
    return core._is_detail_like(sl) and _generated_detail_shape_count(sl) >= 4


def _current_attachment_first_index(prs, d):
    prefix = "AUTO_8D_ATTACH_" + core._attachment_key(d) + "_"
    for i, sl in enumerate(prs.slides):
        if str(getattr(sl, "name", "") or "").startswith(prefix):
            return i
    return None


def _ensure_saved_detail(saved, d, g, mode, request=None):
    """Guarantee a visible generated detail page before native-section creation.

    Attachments may already have been inserted by the stable flow. If the cloned
    detail shell is missing or contains only template markers, repair it in place;
    if no detail slide exists, clone a known 2D~6D template immediately before
    the current issue attachments and fill it.
    """
    from pptx import Presentation

    prs = Presentation(saved)
    candidate = None
    wanted_id = (request or {}).get("slide_id")

    if wanted_id is not None:
        for i, sl in enumerate(prs.slides):
            try:
                if int(sl.slide_id) == int(wanted_id):
                    candidate = i
                    break
            except Exception:
                pass

    if candidate is not None:
        sl = prs.slides[candidate]
        if not _filled_detail_slide(sl):
            s13._update_detail_slide(sl, d, g, mode)
            try:
                s13._clear_slide_text_cache(sl)
            except Exception:
                pass
            prs.save(saved)
            if _filled_detail_slide(sl):
                return candidate, {
                    "name": N(core._new_section_name(d)) or "신규 과제",
                    "slide_index": candidate,
                    "slide_id": int(sl.slide_id),
                }

    # Look for an already-generated current detail before creating another one.
    full, project, _customer = _project_parts(d)
    issue = s13._k(s13._issue_display(d))
    generated = []
    for i, sl in enumerate(prs.slides):
        if not _filled_detail_slide(sl):
            continue
        q = s13._k(s13._slide_text(sl))
        score = 0
        if full and full in q:
            score += 100
        elif project and project in q:
            score += 70
        if issue and issue in q:
            score += 100
        generated.append((score, i))
    if generated:
        generated.sort(reverse=True)
        if generated[0][0] >= 70:
            i = generated[0][1]
            sl = prs.slides[i]
            return i, {
                "name": N(core._new_section_name(d)) or "신규 과제",
                "slide_index": i,
                "slide_id": int(sl.slide_id),
            }

    # No usable detail survived. Recreate it immediately before this issue's
    # first attachment so the final order is Detail -> attachments.
    insert_at = _current_attachment_first_index(prs, d)
    if insert_at is None:
        insert_at = len(prs.slides)

    sl, template_index, _ = _original_clone_detail_shell(
        prs, d, insert_at, None
    )
    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass
    s13._update_detail_slide(sl, d, g, mode)
    try:
        s13._clear_slide_text_cache(sl)
    except Exception:
        pass

    if not _filled_detail_slide(sl):
        raise RuntimeError(
            "상세페이지를 재생성했지만 2D~6D 내용이 작성되지 않았습니다. "
            f"(양식 원본 slide {template_index + 1})"
        )

    prs.save(saved)
    return insert_at, {
        "name": N(core._new_section_name(d)) or "신규 과제",
        "slide_index": insert_at,
        "slide_id": int(sl.slide_id),
    }


def _find_saved_pending_detail(saved, request, d=None):
    """Return the CURRENT index of the cloned detail page after save/COM edits.

    Do not trust the pre-save numeric position. Summary insertion and PowerPoint
    attachment insertion can shift slide positions. Prefer the persistent
    PowerPoint slide-id, then use a tightly-scoped detail-content fallback.
    """
    if not request:
        return None

    from pptx import Presentation

    prs = Presentation(saved)
    wanted_id = request.get("slide_id")

    if wanted_id is not None:
        for i, sl in enumerate(prs.slides):
            try:
                if int(sl.slide_id) == int(wanted_id):
                    if core._is_detail_like(sl):
                        return i
                    break
            except Exception:
                pass

    # Old-index fallback for files where the slide-id was unavailable.
    old_idx = int(request.get("slide_index", -1))
    if 0 <= old_idx < len(prs.slides) and core._is_detail_like(prs.slides[old_idx]):
        return old_idx

    # Final recovery: find a detail-like slide carrying the current project/title.
    full, project, _customer = _project_parts(d or {})
    candidates = []
    for i, sl in enumerate(prs.slides):
        if not core._is_detail_like(sl):
            continue
        q = s13._k(s13._slide_text(sl))
        score = 0
        if full and full in q:
            score += 100
        if project and project in q:
            score += 70
        # Prefer slides closest to the originally requested location.
        distance = abs(i - old_idx) if old_idx >= 0 else 0
        candidates.append((score, -distance, i))

    if candidates:
        candidates.sort(reverse=True)
        best = candidates[0]
        # A project/title hit is safest. If there is exactly one detail-like slide
        # in the whole output, it is also unambiguous enough to recover.
        if best[0] > 0 or len(candidates) == 1:
            return best[2]

    return None


def _verify_pending_detail_slide(saved, request, d=None):
    """Return the saved detail index; fail only when no generated detail exists."""
    idx = _find_saved_pending_detail(saved, request, d)
    if idx is None:
        from pptx import Presentation
        prs = Presentation(saved)
        old_idx = int(request.get("slide_index", -1))
        raise RuntimeError(
            "신규 구역용 2D~6D 상세페이지를 저장 결과에서 찾지 못했습니다. "
            f"(생성 당시 위치: {old_idx + 1 if old_idx >= 0 else '-'}페이지, "
            f"저장 후 전체: {len(prs.slides)}페이지)"
        )
    return idx


# Replace only the invalid native-section writer.  weekly_fix5 still clones and
# fills the detail slide exactly as before.
core._create_native_section = _create_native_section_pending

_active_weekly_before_native_section = core.base.weekly


def _weekly_with_real_native_section(src, out, d, g, mode):
    global _pending_user_native_section
    _pending_user_native_section = None

    msg, saved = _active_weekly_before_native_section(src, out, d, g, mode)
    request = _pending_user_native_section
    force_new = str((g or {}).get("_weekly_create_new_section") or "").lower() in (
        "1",
        "true",
        "yes",
    )

    if force_new or request is not None:
        # First guarantee that a real, FILLED detail page exists. This also
        # repairs the previous "section + attachments, but no detail" result.
        saved_detail_index, request = _ensure_saved_detail(
            saved, d, g, mode, request
        )
        _create_native_section_com_saved(
            saved, request["name"], saved_detail_index
        )
        msg += (
            f" / 신규 구역 [{request['name']}] 생성 완료"
            f" / 상세 page {saved_detail_index + 1} 생성 확인"
        )

    return msg, saved


core.base.weekly = _weekly_with_real_native_section


# ---------------------------------------------------------------------------
# Final required-input gate.
# Keep the original warning wording/flow for classification confirmation and
# PMS/PLM.  This last-loaded guard only restores missing text-field blocking,
# then delegates to the original UI chain unchanged.
# ---------------------------------------------------------------------------
_REQUIRED_TEXT_FIELDS = (
    ("team", "담당팀", "담당팀을 입력해 주세요."),
    ("task_name", "고객사/과제명", "고객사/과제명을 입력해 주세요."),
    ("owner", "담당자", "담당자를 입력해 주세요."),
    ("sample", "발생 샘플", "발생 샘플을 입력해 주세요."),
)

_REQUIRED_CLASSIFICATION_VALUES = (
    ("form_factor", "폼팩터"),
    ("product_type", "제품 타입"),
    ("occurrence_site", "발생처"),
    ("stage", "개발 단계"),
)


def _required_user_input_error(g):
    g = g or {}

    # Restore the old one-field-at-a-time "입력 확인" behavior for 담당 정보.
    for key, _label, message in _REQUIRED_TEXT_FIELDS:
        if not N(g.get(key)):
            return "입력 확인", message

    # If a classification value itself is empty, use the same wording already
    # used by EnterpriseAppV2 for classification confirmation.
    missing_classification = [
        label
        for key, label in _REQUIRED_CLASSIFICATION_VALUES
        if not N(g.get(key))
    ]
    if missing_classification:
        return (
            "분류 정보 확인",
            "아래 분류값을 확인해 주세요.\n\n"
            + " / ".join(missing_classification)
            + "\n\n값을 직접 선택하거나, 현재 기본값이 맞으면 오른쪽 ○를 클릭해 ✓로 확인해 주세요.",
        )

    # PMS/PLM is intentionally NOT checked here. The original EnterpriseAppV2
    # warning remains authoritative and runs only when Issue DB Excel is selected.
    return None


_original_enterprise_run_required_gate = enterprise_v3.EnterpriseAppV3.run


def _run_with_required_inputs(self):
    g = self.gui()
    error = _required_user_input_error(g)
    if error:
        title, message = error
        try:
            self.status_var.set("READY · 입력 정보 확인이 필요합니다.")
        except Exception:
            pass
        return ui.warning(self, title, message)

    # Delegate classification ○→✓ and Issue-DB-only PMS/PLM checks to the
    # original V2 flow so their existing popup wording remains unchanged.
    return _original_enterprise_run_required_gate(self)


enterprise_v3.EnterpriseAppV3.run = _run_with_required_inputs


# ===========================================================================
# FINAL single-path weekly writer
# ===========================================================================
# Repeated wrappers made real-file behavior diverge from the isolated helpers.
# From here on, the GUI uses ONE deterministic weekly path:
#   1) find existing project page directly
#   2) update/clone that page while preserving its visible project label
#   3) create and verify the detail page
#   4) save
#   5) append attachments after detail
#   6) create native section before the verified detail page
# ===========================================================================

def _simple_summary_table_candidates(sl):
    out = []
    for sh in v310.walk(sl):
        if not getattr(sh, "has_table", False):
            continue
        tb = sh.table
        for hr in range(min(8, len(tb.rows))):
            hm = _summary_header_map(tb, hr)
            if "task" not in hm:
                continue
            # Existing-project lookup is intentionally permissive. The exact
            # task text on the page is the primary identity, not header version.
            out.append((tb, hr, hm))
            break
    return out


def _summary_identity_keys(d, g=None):
    """Return separator/newline-insensitive full/project identities.

    Matching contract:
      A_B == A B == A\\nB
      full customer+project exact wins
      exact project alone is also the same project
    """
    d = d or {}
    g = g or {}

    customer_raw = N(d.get("customer"))
    selected_raw = N(g.get("task_name")) or N(d.get("task_name"))
    canonical_raw = N(s13._customer_task(d))

    full_keys = set()
    project_keys = set()

    def add_full(value):
        q = s13._k(value)
        if q:
            full_keys.add(q)

    def add_project(value):
        q = s13._k(value)
        if q:
            project_keys.add(q)

    # Canonical customer_task is always a full identity when available.
    add_full(canonical_raw)

    # Catalog-aware split first.
    alias = _catalog_separator_alias(selected_raw)
    split_source = alias or selected_raw
    split_customer, split_project = catalog.split_customer_task(split_source)

    # If the selected value already contains customer_project, it is a full
    # identity. Otherwise it is PROJECT-ONLY and must not be promoted to full.
    if split_customer and split_project:
        add_full(split_source)
        add_project(split_project)
        if not customer_raw:
            customer_raw = split_customer
    else:
        add_project(selected_raw)

    # Canonical customer_task may contain the full identity even when selected
    # task is project-only.
    can_customer, can_project = catalog.split_customer_task(canonical_raw)
    if can_customer and can_project:
        add_full(canonical_raw)
        add_project(can_project)

    # Explicit customer + project combinations. Spaces/newlines/underscores all
    # collapse to the same _k() value.
    for project_q in list(project_keys):
        customer_q = s13._k(customer_raw)
        if customer_q:
            full_keys.add(customer_q + project_q)

    return full_keys, project_keys


def _simple_task_match_rank(raw, full_keys, project_keys):
    """Rank visible task/cell/sentence text using the user's old simple contract."""
    q = s13._k(raw)
    if not q:
        return 0

    # 1) Exact customer+project, regardless of _, space or newline.
    if q in full_keys:
        return 400

    # 2) Exact project name alone is the same project.
    if q in project_keys:
        return 360

    # 3) A sentence containing the exact customer+project identity.
    for key in full_keys:
        if len(key) >= 4 and key in q:
            return 320

    # 4) A sentence containing the exact project identity.
    # Keep a minimum length so very short generic fragments do not collide.
    for key in project_keys:
        if len(key) >= 4 and key in q:
            return 280

    return 0

def _find_existing_summary_project(prs, d, g):
    full_keys, project_keys = _summary_identity_keys(d, g)
    selected = N((g or {}).get("task_name")) or N((d or {}).get("task_name"))

    hits = []
    for si, sl in enumerate(prs.slides):
        if core._is_detail_like(sl):
            continue
        for tb, hr, hm in _simple_summary_table_candidates(sl):
            rows = []
            best_rank = 0
            display = ""
            for r in range(hr + 1, len(tb.rows)):
                raw = N(s13._row_text(tb, r, hm["task"]))
                rank = _simple_task_match_rank(raw, full_keys, project_keys)
                if rank:
                    rows.append(r)
                    if rank > best_rank:
                        best_rank = rank
                        display = raw

            if rows:
                # Include blank/merged continuation issue rows belonging to this
                # visible project until the next non-empty project label.
                last = max(rows)
                for r in range(last + 1, len(tb.rows)):
                    raw = N(s13._row_text(tb, r, hm["task"]))
                    if raw:
                        break
                    if _summary_row_has_payload(tb, r, hm):
                        rows.append(r)
                hits.append((best_rank, si, tb, hr, hm, rows, display))
                continue

            # Last fallback mirrors the old "find the project page" behavior:
            # exact normalized project text somewhere on this slide + a task
            # column is enough to choose the page. This covers merged/template
            # layouts where python-pptx cannot expose the visible task cell.
            slide_text = s13._slide_text(sl)
            slide_rank = _simple_task_match_rank(
                slide_text, full_keys, project_keys
            )
            if slide_rank:
                payload_rows = [
                    r for r in range(hr + 1, len(tb.rows))
                    if _summary_row_has_payload(tb, r, hm)
                ]
                if payload_rows:
                    # Recover an existing visible label if possible; otherwise
                    # preserve the user-selected spelling.
                    visible = ""
                    for r in range(hr + 1, len(tb.rows)):
                        raw = N(s13._row_text(tb, r, hm["task"]))
                        if raw:
                            visible = raw
                            break
                    hits.append(
                        (slide_rank, si, tb, hr, hm, payload_rows, visible or selected)
                    )

    if not hits:
        return None

    # Highest exactness first; among equal matches use the LAST page.
    hits.sort(key=lambda x: (x[0], x[1]))
    best_rank = hits[-1][0]
    same_rank = [x for x in hits if x[0] == best_rank]
    return max(same_rank, key=lambda x: x[1])


def _clone_matched_summary_page(prs, source_index, display, g):
    before = {_slide_id_value(sl) for sl in prs.slides}
    s13._clone_slide_with_rels(prs, source_index)
    created = next(
        (_slide_id_value(sl) for sl in prs.slides if _slide_id_value(sl) not in before),
        None,
    )
    if created is None:
        raise RuntimeError("일치 과제의 요약페이지 복제본을 확인하지 못했습니다.")

    moved = _move_slide_by_id(prs, created, source_index + 1)
    if moved is None:
        raise RuntimeError("일치 과제의 요약페이지를 바로 다음 위치로 이동하지 못했습니다.")

    sl = prs.slides[moved]
    found = _simple_summary_table_candidates(sl)
    if not found:
        raise RuntimeError("복제한 일치 과제 페이지에서 과제명 표를 찾지 못했습니다.")

    # Prefer the table that still contains the source project's visible label.
    chosen = None
    display_key = s13._k(display)
    for tb, hr, hm in found:
        for r in range(hr + 1, len(tb.rows)):
            if display_key and s13._k(s13._row_text(tb, r, hm["task"])) == display_key:
                chosen = (tb, hr, hm)
                break
        if chosen:
            break
    if chosen is None:
        chosen = found[0]

    tb, hr, hm = chosen
    row = s14._clear_summary_data(tb, hr)
    # _clear_summary_data intentionally clears values; restore the exact visible
    # project label from the source page before writing the new issue.
    if "task" in hm:
        try:
            base.set_cell_text(tb.cell(row, hm["task"]), display, 8)
        except Exception:
            tb.cell(row, hm["task"]).text = display

    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass
    return moved, tb, hr, row


def _update_summary_old_style(prs, d, g, mode):
    hit = _find_existing_summary_project(prs, d, g)
    issue = s13._issue_display(d)
    original_customer_task = s13._customer_task

    if hit is not None:
        _rank, si, tb, hr, hm, rows, display = hit
        display = N(display) or N((g or {}).get("task_name")) or N((d or {}).get("task_name"))
        try:
            s13._customer_task = lambda _d: display

            if mode == "existing":
                best = None
                for r in rows:
                    sc = s13._row_match_score(tb, r, hm, display, issue)
                    if best is None or sc > best[0]:
                        best = (sc, r)
                if best and best[0] >= 140:
                    s14._write_summary_row(tb, best[1], hr, d, g)
                    return si, best[1], f"기존 과제 [{display}] 행 업데이트"

            after = max(rows)
            if s14._summary_row_insert_fits(prs, si, tb, after):
                row = s14._insert_row_after(tb, after)
                s14._write_summary_row(tb, row, hr, d, g)
                return si, row, f"기존 과제 [{display}] 마지막 행 아래 삽입"

            nsi, ntb, nhr, nrow = _clone_matched_summary_page(
                prs, si, display, g
            )
            s14._write_summary_row(ntb, nrow, nhr, d, g)
            return nsi, nrow, f"기존 과제 [{display}] 마지막 페이지 복제 후 업데이트"
        finally:
            s13._customer_task = original_customer_task

    # No existing project at all: use only the strict generic summary template.
    pages = _summary_pages_flexible(prs)
    if not pages:
        raise RuntimeError(
            "기존 과제명을 찾지 못했고, 신규 과제용 주간회의 요약 양식도 찾지 못했습니다."
        )
    selected = N((g or {}).get("task_name")) or N(s13._customer_task(d))
    try:
        s13._customer_task = lambda _d: selected
        si, tb, hr, row = _prepare_new_summary_page_stable(prs, pages, g)
        s14._write_summary_row(tb, row, hr, d, g)
        return si, row, f"과제 신규 생성 [{selected}]"
    finally:
        s13._customer_task = original_customer_task


def _detail_slide_marker(d):
    return "AUTO_8D_DETAIL_" + core._attachment_key(d)


def _mark_detail_slide(sl, d):
    marker = _detail_slide_marker(d)
    try:
        sl.name = marker
    except Exception:
        try:
            sl._element.cSld.set("name", marker)
        except Exception:
            pass
    return marker


def _is_marked_detail(sl):
    return str(getattr(sl, "name", "") or "").startswith("AUTO_8D_DETAIL_")


def _verified_generated_detail(sl):
    names = {
        str(getattr(sh, "name", "") or "")
        for sh in v310.walk(sl)
    }
    required = (
        "AUTO_8D_TEXT_2D",
        "AUTO_8D_TEXT_3D",
        "AUTO_8D_TEXT_4D_CAUSE",
        "AUTO_8D_TEXT_5D",
        "AUTO_8D_TEXT_6D",
    )
    generated = sum(1 for key in required if key in names)
    return core._is_detail_like(sl) and generated >= 4


def _create_detail_direct(prs, d, g, mode, matched_section, force_new):
    target = None
    if mode == "existing" and not force_new:
        target = _find_existing_detail_exact_issue(prs, d)

    if target is not None:
        s13._update_detail_slide(prs.slides[target], d, g, mode)
        _mark_detail_slide(prs.slides[target], d)
        s13._clear_slide_text_cache(prs.slides[target])
        if not _verified_generated_detail(prs.slides[target]):
            raise RuntimeError("기존 상세페이지에 2D~6D 내용을 작성하지 못했습니다.")
        return target, _slide_id_value(prs.slides[target]), "기존 상세 업데이트"

    if matched_section and matched_section.get("indices") and not force_new:
        insert_at = max(matched_section["indices"]) + 1
    else:
        # New native section: keep the detail+attachments together at the end.
        insert_at = len(prs.slides)

    template_index = core._template_detail_index(prs, d)
    s13._clone_slide_with_rels(prs, template_index)
    s13._move_last_slide_to(prs, insert_at)
    sl = prs.slides[insert_at]

    # Use the stable cleanup, but never remove the D marker/title skeleton.
    core._clear_cloned_issue_content(sl)
    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass

    s13._update_detail_slide(sl, d, g, mode)
    _mark_detail_slide(sl, d)
    s13._clear_slide_text_cache(sl)

    if not _verified_generated_detail(sl):
        raise RuntimeError(
            "상세페이지 복제 후 2D~6D 실제 내용 생성 확인에 실패했습니다. "
            f"(복제 양식 slide {template_index + 1})"
        )

    if matched_section and not force_new:
        try:
            core._add_slide_to_native_section(prs, matched_section, insert_at)
        except Exception:
            pass

    return insert_at, _slide_id_value(sl), f"상세 신규 생성 (양식 slide {template_index + 1})"


def _find_saved_slide_index_by_id(saved, slide_id):
    from pptx import Presentation
    prs = Presentation(saved)
    if slide_id is not None:
        for i, sl in enumerate(prs.slides):
            try:
                if int(sl.slide_id) == int(slide_id):
                    return i
            except Exception:
                pass
    return None


def _weekly_single_path(src, out, d, g, mode):
    from pptx import Presentation

    try:
        s13._clear_slide_text_cache()
    except Exception:
        pass

    prs = Presentation(src)
    summary_index, _row, summary_action = _update_summary_old_style(
        prs, d, g, mode
    )

    # Resolve native-section intent only AFTER the summary has been settled.
    matched_section = core._selected_section(prs, d, g)
    force_new = str((g or {}).get("_weekly_create_new_section") or "").lower() in (
        "1", "true", "yes"
    )
    if matched_section is None:
        force_new = True

    detail_index, detail_id, detail_action = _create_detail_direct(
        prs, d, g, mode, matched_section, force_new
    )

    # Version naming remains identical to final_output_polish.
    versioned = final_polish._version_path(out)
    Path(versioned).parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(versioned)
        saved = str(versioned)
    except PermissionError:
        p = Path(versioned)
        saved = str(
            p.with_name(
                p.stem
                + "_"
                + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                + p.suffix
            )
        )
        prs.save(saved)

    # The DETAIL must exist before attachments are allowed to run.
    current_detail_index = _find_saved_slide_index_by_id(saved, detail_id)
    if current_detail_index is None:
        raise RuntimeError("저장 후 상세페이지 ID를 찾지 못했습니다. 유첨 추가를 중단했습니다.")

    check = Presentation(saved)
    if not _verified_generated_detail(check.slides[current_detail_index]):
        raise RuntimeError("저장된 상세페이지의 2D~6D 내용 확인에 실패했습니다. 유첨 추가를 중단했습니다.")

    attached, attach_error = core._append_8d_attachments_safe(
        saved, (g or {}).get("ppt8d", ""), current_detail_index, d
    )

    # Attachment insertion can shift numeric indices; resolve the SAME detail by ID.
    final_detail_index = _find_saved_slide_index_by_id(saved, detail_id)
    if final_detail_index is None:
        raise RuntimeError("유첨 추가 후 상세페이지가 사라졌습니다. 신규 구역 생성을 중단했습니다.")

    final_check = Presentation(saved)
    if (
        not _verified_generated_detail(final_check.slides[final_detail_index])
        or not _is_marked_detail(final_check.slides[final_detail_index])
    ):
        raise RuntimeError(
            "유첨 추가 후 상세페이지 보존 확인에 실패했습니다. "
            "상세 없이 유첨만 저장하는 결과는 생성하지 않았습니다."
        )

    if force_new:
        _create_native_section_com_saved(
            saved, N(core._new_section_name(d)) or "신규 과제", final_detail_index
        )
        section_action = (
            f"신규 구역 [{N(core._new_section_name(d)) or '신규 과제'}] 생성"
        )
    else:
        section_action = (
            f"기존 구역 [{N((matched_section or {}).get('name'))}] 사용"
        )

    try:
        final_polish._clean_summary_placeholders(saved)
    except Exception:
        pass

    attach_text = (
        f"유첨 {attached}페이지"
        if attached
        else ("유첨 없음" if not attach_error else "유첨 실패: " + N(attach_error)[:180])
    )
    return (
        "주간회의 PPT: "
        + summary_action
        + " / "
        + detail_action
        + " / "
        + section_action
        + " / "
        + attach_text
        + f" / 상세 page {final_detail_index + 1}"
        + f" / 버전 저장={Path(saved).name}",
        saved,
    )


# Absolute last assignment: no earlier wrapper can bypass this writer.
core.base.weekly = _weekly_single_path


# ---------------------------------------------------------------------------
# Reduced-output protection.
# The update-only generator previously removed a newly cloned detail page when
# its XML signature matched a source/template slide. Explicitly marked current
# detail slides must always survive the reduced PPT.
# ---------------------------------------------------------------------------
_original_ppt_update_only_user = output_variant._ppt_update_only


def _ppt_update_only_keep_current_detail(source, saved):
    from collections import Counter
    from pptx import Presentation

    src = Presentation(source)
    dst = Presentation(saved)
    source_counts = Counter(output_variant._slide_signature(sl) for sl in src.slides)

    keep = []
    for i, slide in enumerate(dst.slides):
        name = str(getattr(slide, "name", "") or "")

        # Current generated detail and current AUTO attachments are always output
        # artifacts, even when their content resembles an existing source page.
        if name.startswith("AUTO_8D_DETAIL_") or name.startswith("AUTO_8D_ATTACH_"):
            keep.append(i)
            continue

        sig = output_variant._slide_signature(slide)
        if source_counts.get(sig, 0) > 0:
            source_counts[sig] -= 1
        else:
            keep.append(i)

    if not keep:
        return None, 0

    keep_set = set(keep)
    for i in range(len(dst.slides) - 1, -1, -1):
        if i not in keep_set:
            output_variant._remove_slide(dst, i)

    target = Path(saved).with_name(
        Path(saved).stem + "_업데이트사항만" + Path(saved).suffix
    )
    dst.save(target)

    # Never hand the user a reduced PPT that contains attachments for this run
    # but has no generated detail page.
    check = Presentation(target)
    has_detail = any(_is_marked_detail(sl) for sl in check.slides)
    has_attach = any(
        str(getattr(sl, "name", "") or "").startswith("AUTO_8D_ATTACH_")
        for sl in check.slides
    )
    if has_attach and not has_detail:
        try:
            Path(target).unlink()
        except Exception:
            pass
        raise RuntimeError(
            "축약본에서 상세페이지가 누락되어 저장을 중단했습니다. "
            "전체 업데이트 본은 유지합니다."
        )

    return target, len(keep)


output_variant._ppt_update_only = _ppt_update_only_keep_current_detail
