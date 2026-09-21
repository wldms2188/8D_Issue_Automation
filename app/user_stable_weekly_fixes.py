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
import tkinter as tk
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
import main_enterprise as enterprise_main
import main_final as legacy_final
import ui_enterprise as ui
import final_output_polish as final_polish
import output_variant_final as output_variant
import weekly_style_final as weekly_style

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


def _shape_rgb(value):
    try:
        rgb = value.rgb
        if rgb is None:
            return None
        return tuple(int(x) for x in rgb)
    except Exception:
        return None


def _is_red_annotation_shape(sh):
    """Detect copied red circles/underlines/callouts used as old-issue marks."""
    # Never remove a text-bearing fixed D marker/title just from its color.
    text = N(getattr(sh, "text", ""))
    if text and _static_cloned_detail_shape_preserve_layout(sh):
        return False

    colors = []
    try:
        colors.append(_shape_rgb(sh.line.color))
    except Exception:
        pass
    try:
        colors.append(_shape_rgb(sh.fill.fore_color))
    except Exception:
        pass

    for rgb in colors:
        if not rgb:
            continue
        r, g, b = rgb
        # Strong red only. Yellow/orange 7D markers and green D markers survive.
        if r >= 170 and g <= 95 and b <= 95 and r >= g + 70 and r >= b + 70:
            return True
    return False


def _is_static_detail_child(sh):
    """True only for fixed template marker/title/header objects."""
    if getattr(sh, "has_table", False):
        return False
    st = getattr(sh, "shape_type", None)
    if st == MSO_SHAPE_TYPE.GROUP:
        return False
    return _static_cloned_detail_shape_preserve_layout(sh)


_FIXED_DETAIL_TEXT_KEYS = {
    "2d", "3d", "4d", "5d", "6d", "7d",
    "현상", "문제현상",
    "임시조치", "임시대응", "임시대책", "임시대책필요시",
    "원인분석", "발생원인", "유출원인", "시스템원인",
    "개선대책", "효과검증", "유효성점검", "수평전개",
    "signal", "이슈기인", "발생단계",
}

_FIXED_DETAIL_PREFIXES = (
    ("2d", ("현상", "문제현상")),
    ("3d", ("임시조치", "임시대응", "임시대책", "임시대책필요시")),
    ("4d", ("원인분석", "발생원인", "유출원인", "시스템원인")),
    ("5d", ("개선대책",)),
    ("6d", ("효과검증", "유효성점검")),
    ("7d", ("수평전개",)),
)


def _fixed_detail_text_only(text):
    """Return only fixed template title text; drop copied issue body text."""
    raw = N(text)
    if not raw:
        return ""

    # Keep exact static lines first.
    kept = []
    for line in str(raw).splitlines():
        line_n = N(line)
        q = s13._k(line_n)
        if q in _FIXED_DETAIL_TEXT_KEYS:
            kept.append(line_n)

    if kept:
        return "\n".join(kept)

    # Combined title/body in a single line/paragraph, e.g.
    # "4D 발생원인 체결 토크 부족..." -> preserve only "4D 발생원인".
    q = s13._k(raw)
    for marker, titles in _FIXED_DETAIL_PREFIXES:
        if marker not in q:
            continue
        for title in titles:
            tq = s13._k(title)
            pos_m = q.find(marker)
            pos_t = q.find(tq)
            if pos_m >= 0 and pos_t >= pos_m:
                return marker.upper() + " " + title

    # Title without D marker.
    for title in (
        "현상", "문제현상", "임시조치", "임시대응", "임시대책",
        "원인분석", "발생원인", "유출원인", "시스템원인",
        "개선대책", "효과검증", "유효성점검", "수평전개",
    ):
        if s13._k(raw) == s13._k(title):
            return title

    return ""


def _set_shape_text_preserve_first_run(sh, value):
    """Change text while preserving the first run's formatting when possible."""
    if not hasattr(sh, "text_frame"):
        return False
    try:
        tf = sh.text_frame
        paragraphs = list(tf.paragraphs)
        first_run = None
        for p in paragraphs:
            runs = list(p.runs)
            if runs and first_run is None:
                first_run = runs[0]

        if first_run is not None:
            # Clear everything first. pptx run proxy identity is not stable
            # enough to safely compare with "is" while iterating again.
            for p in paragraphs:
                for run in list(p.runs):
                    run.text = ""
            first_run.text = N(value)
            return True

        sh.text = N(value)
        return True
    except Exception:
        try:
            sh.text = N(value)
            return True
        except Exception:
            return False


def _shape_hits_detail_content_zone(sh):
    try:
        return any(
            core._overlap_ratio(sh, zone) > 0.005
            for zone in core._content_zones()
        )
    except Exception:
        return False


def _clear_old_detail_text_shape(sh):
    """Clear copied issue text but keep the template shape/frame itself."""
    text = N(getattr(sh, "text", ""))
    if not text:
        return

    fixed = _fixed_detail_text_only(text)
    if fixed:
        # Mixed fixed-title + old issue body -> retain title only.
        if s13._k(text) != s13._k(fixed):
            _set_shape_text_preserve_first_run(sh, fixed)
        return

    _set_shape_text_preserve_first_run(sh, "")


def _clean_preserved_detail_group(group):
    """Recursively strip old issue text/pictures from a preserved template group."""
    for child in list(getattr(group, "shapes", ())):
        if _is_red_annotation_shape(child):
            _remove_shape(child)
            continue

        st = getattr(child, "shape_type", None)
        if st in _VISUAL_TYPES:
            _remove_shape(child)
            continue

        if st == MSO_SHAPE_TYPE.GROUP:
            _clean_preserved_detail_group(child)
            continue

        if getattr(child, "has_table", False):
            try:
                core._clear_cloned_table_content(child)
            except Exception:
                pass
            continue

        text = N(getattr(child, "text", ""))
        if text:
            _clear_old_detail_text_shape(child)


def _purge_cloned_detail_artifacts(sl):
    """Aggressively clean copied issue content while retaining the detail shell.

    - red annotations: remove
    - copied pictures/media: already removed by visual cleanup
    - groups: recurse and keep only fixed marker/title text
    - ungrouped text in 2D~6D content zones: clear copied text
    - combined title+body text box: keep the fixed title only
    """
    for sh in list(sl.shapes):
        if _is_red_annotation_shape(sh):
            _remove_shape(sh)
            continue

        st = getattr(sh, "shape_type", None)
        if st == MSO_SHAPE_TYPE.GROUP:
            _clean_preserved_detail_group(sh)
            continue

        if getattr(sh, "has_table", False):
            # Keep table geometry; the core cleaner removes old value cells.
            try:
                core._clear_cloned_table_content(sh)
            except Exception:
                pass
            continue

        name = str(getattr(sh, "name", "") or "")
        if name.startswith("AUTO_8D_"):
            continue

        text = N(getattr(sh, "text", ""))
        if text and _shape_hits_detail_content_zone(sh):
            _clear_old_detail_text_shape(sh)


def _clone_detail_shell_clean(prs, d, insert_at, matched_section=None):
    result = _original_clone_detail_shell(prs, d, insert_at, matched_section)
    sl, template_index, section = result

    try:
        _remove_copied_visuals(sl)
    except Exception:
        pass
    try:
        _purge_cloned_detail_artifacts(sl)
    except Exception:
        pass

    # Do not allow an "attachments only" result. The cloned page itself must
    # still contain a recognizable 2D~6D detail structure before continuing.
    if not _strict_detail_template_fingerprint(sl)["ok"]:
        raise RuntimeError(
            "복제한 페이지가 실제 2D~6D 상세 양식이 아닙니다. "
            "일정표/게이트/로드맵 페이지는 상세로 사용하지 않습니다. "
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


def _short_shape_tokens(sl):
    """Collect short visible labels, including children of grouped markers."""
    out = []
    for sh in v310.walk(sl):
        text = N(getattr(sh, "text", ""))
        if text and len(text) <= 40:
            out.append((s13._k(text), text))
    return out


def _strict_detail_template_fingerprint(sl):
    """Recognize real 8D detail layouts without requiring separate marker shapes.

    Real company templates can place D labels inside groups/tables/combined
    title boxes, so geometry/shape separation is not mandatory.  Timeline/Gate
    pages are rejected by their schedule vocabulary and weak 8D structure.
    """
    whole_text = s13._slide_text(sl)
    whole = s13._k(whole_text)

    d_tokens = ("2d", "3d", "4d", "5d", "6d")
    d_hits = {token for token in d_tokens if token in whole}

    metadata_hits = sum(
        1 for token in ("signal", "이슈기인", "발생단계") if token in whole
    )
    title_hits = sum(
        1
        for token in (
            "현상", "문제현상",
            "임시조치", "임시대응", "임시대책",
            "원인분석", "발생원인", "유출원인", "시스템원인",
            "개선대책", "효과검증", "유효성점검", "수평전개",
        )
        if token in whole
    )

    # Reuse the legacy score because it already reads text from groups/tables.
    legacy_structure = s13._detail_structure_score(sl)

    schedule_hits = sum(
        whole.count(token)
        for token in ("cv", "dv", "pd", "pv", "sop", "gate", "target", "now")
    )

    # A true detail page normally carries most D stages.  This deliberately
    # accepts legacy/grouped layouts that do not expose separate marker shapes.
    strong_detail = (
        len(d_hits) >= 4
        and legacy_structure >= 5
        and (title_hits >= 2 or metadata_hits >= 1 or len(d_hits) == 5)
    )

    # A schedule/timeline may mention one or two D items, but if schedule terms
    # dominate and the detail titles/meta are weak it must never be a template.
    schedule_dominant = (
        schedule_hits >= 5
        and title_hits < 3
        and metadata_hits < 2
    )

    return {
        "ok": bool(strong_detail and not schedule_dominant),
        "markers": d_hits,
        "metadata_hits": metadata_hits,
        "title_hits": title_hits,
        "schedule_hits": schedule_hits,
        "legacy_structure": legacy_structure,
    }


def _detail_template_score(sl, d):
    """Score only slides proven to have the real 8D detail skeleton."""
    fp = _strict_detail_template_fingerprint(sl)
    if not fp["ok"]:
        return -1

    text = s13._slide_text(sl)
    q = s13._k(text)

    # Structure dominates project similarity. A wrong-layout same-project slide
    # must NEVER beat a true detail template from another project.
    score = (
        len(fp["markers"]) * 100
        + fp["metadata_hits"] * 40
        + fp["title_hits"] * 10
    )

    if "7d" in q or "수평전개" in q:
        score += 30

    full, project, _customer = _project_parts(d)
    if full and full in q:
        score += 80
    elif project and project in q:
        score += 55

    return score


def _template_detail_index_any_section(prs, d):
    """Find a proven detail template anywhere; never fall back to a roadmap."""
    summaries = set(core._summary_indices(prs))
    candidates = []

    for i, sl in enumerate(prs.slides):
        if i in summaries:
            continue
        score = _detail_template_score(sl, d)
        if score >= 0:
            candidates.append((score, i))

    if not candidates:
        raise ValueError(
            "주간회의 PPT에서 실제 2D~6D 상세 양식을 찾지 못했습니다. "
            "일정표/게이트/로드맵 페이지는 상세 양식으로 사용하지 않습니다."
        )

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
        if _strict_detail_template_fingerprint(sl)["ok"]:
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
        if i in summaries or not _strict_detail_template_fingerprint(sl)["ok"]:
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
        if i in summaries or not _strict_detail_template_fingerprint(sl)["ok"]:
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
    return (
        _strict_detail_template_fingerprint(sl)["ok"]
        and _generated_detail_shape_count(sl) >= 4
    )


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

        # Already valid: do NOT keep searching and clone a duplicate detail page.
        if _filled_detail_slide(sl):
            return candidate, {
                "name": N(core._new_section_name(d)) or "신규 과제",
                "slide_index": candidate,
                "slide_id": int(sl.slide_id),
            }

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
                    if _strict_detail_template_fingerprint(sl)["ok"]:
                        return i
                    break
            except Exception:
                pass

    # Old-index fallback for files where the slide-id was unavailable.
    old_idx = int(request.get("slide_index", -1))
    if (
        0 <= old_idx < len(prs.slides)
        and _strict_detail_template_fingerprint(prs.slides[old_idx])["ok"]
    ):
        return old_idx

    # Final recovery: find a detail-like slide carrying the current project/title.
    full, project, _customer = _project_parts(d or {})
    candidates = []
    for i, sl in enumerate(prs.slides):
        if not _strict_detail_template_fingerprint(sl)["ok"]:
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


def _project_parenthetical_parts(value, customer_hint=""):
    """Return only parenthetical parts that belong to the PROJECT identity."""
    raw = N(value)
    alias = _catalog_separator_alias(raw)
    source = alias or raw

    customer, project = catalog.split_customer_task(source)
    if not project:
        project = source

    # If a customer hint was prepended manually, strip only that leading
    # customer token before reading project qualifiers.
    hint = N(customer_hint)
    if hint:
        q_source = s13._k(source)
        q_hint = s13._k(hint)
        if q_hint and q_source.startswith(q_hint) and not customer:
            project = source[len(hint):].lstrip(" _-\n\t")

    return tuple(
        s13._k(x)
        for x in re.findall(r"\(([^()]*)\)", N(project))
        if s13._k(x)
    )


def _required_project_parens(d, g=None):
    """Backward-compatible name for the project-specific parenthesis rule."""
    d = d or {}
    g = g or {}
    selected = N(g.get("task_name")) or N(d.get("task_name"))
    return _project_parenthetical_parts(selected, N(d.get("customer")))


def _project_parentheses_match(raw, expected_parts):
    """Only project-name parentheses are compared; unrelated note parentheses are ignored.

    The normal identity match still decides whether the project text exists.
    This helper is only an additional guard when the INPUT project itself has
    qualifiers such as (EU), (US), (LHD), etc.
    """
    if not expected_parts:
        return True

    q = s13._k(raw)
    # Each expected qualifier must be present as part of the normalized project
    # identity. Do not require every parenthesis found in the whole cell/page
    # to equal the input; e.g. "(2차)" elsewhere is unrelated.
    return all(part and part in q for part in expected_parts)



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


def _simple_task_match_rank(raw, full_keys, project_keys, required_parens=()):
    """Rank visible text by normalized identity.

    s13._k() removes separators and punctuation but KEEPS the text inside
    parentheses, so:
      A_B == A B == A\\nB
      EB-L(EU) != EB-L(US)
    No separate parenthesis-list equality is used because a summary cell/page
    can legitimately contain unrelated parentheses elsewhere in the sentence.
    """
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

def _legacy_summary_target_keys(d, g=None):
    """Simple project-name matching restored from the previously working behavior.

    Separators are ignored by s13._k():
      A_B == A B == A\\nB
    Parenthetical CONTENT is preserved:
      EB-L(EU) != EB-L(US)
    """
    d = d or {}
    g = g or {}

    selected = N(g.get("task_name")) or N(d.get("task_name"))
    customer = N(d.get("customer"))

    # Prefer the user's supplied catalog spelling only when it is an exact
    # separator-insensitive alias. This does not control matching; it only helps
    # us recover the project-only part reliably.
    canonical = _catalog_separator_alias(selected) or selected
    cat_customer, cat_project = catalog.split_customer_task(canonical)

    customer_q = s13._k(cat_customer or customer)
    selected_q = s13._k(canonical)

    if cat_customer and cat_project:
        project_q = s13._k(cat_project)
        full_q = selected_q
    else:
        # No underscore in the visible/input spelling. If it starts with the
        # separately entered customer, the remainder is the project name.
        if customer_q and selected_q.startswith(customer_q) and len(selected_q) > len(customer_q):
            project_q = selected_q[len(customer_q):]
            full_q = selected_q
        else:
            project_q = selected_q
            full_q = customer_q + project_q if customer_q else project_q

    return full_q, project_q


def _legacy_summary_match_rank(raw, full_q, project_q):
    """Old/simple exact-name-first matcher used only for weekly summary lookup."""
    q = s13._k(raw)
    if not q:
        return 0

    if full_q and q == full_q:
        return 400
    if project_q and q == project_q:
        return 380

    # Existing weekly files sometimes contain the project label plus a note in
    # the same cell. Keep this narrow; normalized qualifier text still makes
    # (EU) and (US) different strings.
    if full_q and len(full_q) >= 4 and full_q in q:
        return 340
    if project_q and len(project_q) >= 4 and project_q in q:
        return 320

    return 0



def _find_existing_summary_project(prs, d, g):
    """Restore the simple project-name search that previously worked.

    No catalog-mode split, no parenthesis side-gate, no fuzzy section logic.
    """
    full_q, project_q = _legacy_summary_target_keys(d, g)
    selected = N((g or {}).get("task_name")) or N((d or {}).get("task_name"))

    hits = []
    for si, sl in enumerate(prs.slides):
        if _strict_detail_template_fingerprint(sl)["ok"]:
            continue

        for tb, hr, hm in _simple_summary_table_candidates(sl):
            rows = []
            best_rank = 0
            display = ""

            for r in range(hr + 1, len(tb.rows)):
                raw = N(s13._row_text(tb, r, hm["task"]))
                rank = _legacy_summary_match_rank(raw, full_q, project_q)
                if rank:
                    rows.append(r)
                    if rank > best_rank:
                        best_rank = rank
                        display = raw

            if rows:
                # Blank/merged continuation rows belong to the last visible
                # project until the next non-empty task label.
                last = max(rows)
                for r in range(last + 1, len(tb.rows)):
                    raw = N(s13._row_text(tb, r, hm["task"]))
                    if raw:
                        break
                    if _summary_row_has_payload(tb, r, hm):
                        rows.append(r)

                hits.append((best_rank, si, tb, hr, hm, rows, display))
                continue

            # Some weekly templates expose the visible project name outside the
            # task cell because of merging/grouping. Use the same simple match
            # against the page text as a final fallback.
            slide_rank = _legacy_summary_match_rank(
                s13._slide_text(sl), full_q, project_q
            )
            if slide_rank:
                payload_rows = [
                    r
                    for r in range(hr + 1, len(tb.rows))
                    if _summary_row_has_payload(tb, r, hm)
                ]
                if payload_rows:
                    visible = ""
                    for r in range(hr + 1, len(tb.rows)):
                        raw = N(s13._row_text(tb, r, hm["task"]))
                        if raw and _legacy_summary_match_rank(raw, full_q, project_q):
                            visible = raw
                            break

                    hits.append(
                        (
                            slide_rank,
                            si,
                            tb,
                            hr,
                            hm,
                            payload_rows,
                            visible or selected,
                        )
                    )

    if not hits:
        return None

    # Same project on multiple pages -> use the LAST matching page.
    best_rank = max(x[0] for x in hits)
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
    fp = _strict_detail_template_fingerprint(sl)
    return fp["ok"] and generated >= 4


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


def _final_weekly_manifest(saved):
    from pptx import Presentation

    prs = Presentation(saved)
    details = []
    attachments = []
    summaries = []

    for i, sl in enumerate(prs.slides):
        name = str(getattr(sl, "name", "") or "")
        if _verified_generated_detail(sl):
            details.append(i)
        if name.startswith("AUTO_8D_ATTACH_"):
            attachments.append(i)
        try:
            if _simple_summary_table_candidates(sl):
                summaries.append(i)
        except Exception:
            pass

    return {
        "slide_count": len(prs.slides),
        "details": details,
        "attachments": attachments,
        "summaries": summaries,
    }


def _assert_final_weekly_detail(saved):
    manifest = _final_weekly_manifest(saved)
    if not manifest["details"]:
        raise RuntimeError(
            "최종 전체본 검증 실패: 2D~6D 상세페이지가 없습니다. "
            "유첨만 남는 결과는 완료 처리하지 않았습니다. "
            f"(전체 {manifest['slide_count']}p, 유첨 {len(manifest['attachments'])}p)"
        )
    return manifest


# ---------------------------------------------------------------------------
# Detail change highlighting.
# weekly_style_final treated "existing mode + newly cloned detail page" as
# unchanged because no AUTO_8D_TEXT_* existed in the pre-update snapshot.
# Any newly-created detail content must be blue, regardless of GUI mode.
# ---------------------------------------------------------------------------
_detail_update_before_blue_fix = s13._update_detail_slide


def _detail_auto_snapshot(sl):
    out = {}
    for sh in v310.walk(sl):
        name = str(getattr(sh, "name", "") or "")
        if name.startswith("AUTO_8D_TEXT_"):
            out[name] = N(getattr(sh, "text", ""))
    return out


def _update_detail_slide_force_new_text_blue(sl, d, g, mode):
    before = _detail_auto_snapshot(sl)
    _detail_update_before_blue_fix(sl, d, g, mode)
    after = _detail_auto_snapshot(sl)

    for sh in v310.walk(sl):
        name = str(getattr(sh, "name", "") or "")
        if not name.startswith("AUTO_8D_TEXT_"):
            continue

        new_text = after.get(name, N(getattr(sh, "text", "")))
        old_exists = name in before
        old_text = before.get(name, "")

        # New shape/text => blue. Existing but changed => blue.
        # Only truly unchanged existing content stays black.
        changed = (not old_exists and bool(new_text)) or (old_exists and old_text != new_text)
        try:
            weekly_style._color_shape(
                sh,
                weekly_style.UPDATE_BLUE if changed else weekly_style.BLACK,
            )
        except Exception:
            pass


s13._update_detail_slide = _update_detail_slide_force_new_text_blue


# Use the pre-polish weekly_fix5 writer for detail generation. It is the
# previously working detail creation path captured before final_output_polish
# wrapped base.weekly.
_known_good_detail_writer = final_polish._original_weekly


def _find_current_detail_after_baseline(src, saved, d, request=None):
    from pptx import Presentation

    dst = Presentation(saved)

    # New-section request contains the exact cloned detail slide id.
    if request and request.get("slide_id") is not None:
        wanted = int(request["slide_id"])
        for i, sl in enumerate(dst.slides):
            try:
                if int(sl.slide_id) == wanted and _filled_detail_slide(sl):
                    return i
            except Exception:
                pass

    # For a NEW issue, identify the new non-attachment filled detail slide by
    # comparing slide ids against the original weekly deck.
    try:
        src_prs = Presentation(src)
        source_ids = {int(sl.slide_id) for sl in src_prs.slides}
    except Exception:
        source_ids = set()

    new_details = []
    for i, sl in enumerate(dst.slides):
        name = str(getattr(sl, "name", "") or "")
        if name.startswith("AUTO_8D_ATTACH_"):
            continue
        try:
            is_new_id = int(sl.slide_id) not in source_ids
        except Exception:
            is_new_id = False
        if is_new_id and _filled_detail_slide(sl):
            new_details.append(i)
    if new_details:
        return new_details[-1]

    # Existing-issue update: the detail slide keeps its source slide id.
    try:
        idx = _find_existing_detail_exact_issue(dst, d)
        if idx is not None and _filled_detail_slide(dst.slides[idx]):
            return idx
    except Exception:
        pass

    return None


def _weekly_baseline_detail_path(src, out, d, g, mode):
    """Restore the known-good detail creator; keep only our summary/section fixes."""
    global _pending_user_native_section
    from pptx import Presentation

    _pending_user_native_section = None

    # Version exactly once, then call the original weekly_fix5 detail writer.
    versioned = final_polish._version_path(out)
    msg, saved = _known_good_detail_writer(src, versioned, d, g, mode)

    request = _pending_user_native_section
    force_new = str((g or {}).get("_weekly_create_new_section") or "").lower() in (
        "1", "true", "yes"
    )

    detail_index = _find_current_detail_after_baseline(src, saved, d, request)
    if detail_index is None:
        raise RuntimeError(
            "최종 전체본 검증 실패: 기존 상세 생성 경로를 실행했지만 "
            "현재 이슈의 2D~6D 상세페이지를 찾지 못했습니다. "
            "유첨만 있는 파일은 완료 처리하지 않았습니다."
        )

    # Mark exactly this run's detail so the reduced output cannot remove it.
    prs = Presentation(saved)
    detail_slide = prs.slides[detail_index]
    _mark_detail_slide(detail_slide, d)
    detail_id = int(detail_slide.slide_id)
    prs.save(saved)

    # If a new native section was requested, create it only after the known-good
    # detail and attachments are already saved.
    if force_new or request is not None:
        current_index = _find_saved_slide_index_by_id(saved, detail_id)
        if current_index is None:
            raise RuntimeError("신규 구역 생성 전 상세페이지 위치를 다시 찾지 못했습니다.")
        section_name = (
            N((request or {}).get("name"))
            or N(core._new_section_name(d))
            or "신규 과제"
        )
        _create_native_section_com_saved(saved, section_name, current_index)
        msg += f" / 신규 구역 [{section_name}] 생성"

    # Final whole-file check: the EXACT marked detail must still be present.
    final_prs = Presentation(saved)
    final_detail = None
    for i, sl in enumerate(final_prs.slides):
        if str(getattr(sl, "name", "") or "") == _detail_slide_marker(d):
            final_detail = i
            break

    if final_detail is None or not _filled_detail_slide(final_prs.slides[final_detail]):
        raise RuntimeError(
            "최종 전체본 검증 실패: 저장/유첨/구역 생성 후 상세페이지가 없습니다. "
            "유첨만 있는 결과는 저장 완료로 처리하지 않았습니다."
        )

    attach_pages = [
        i + 1
        for i, sl in enumerate(final_prs.slides)
        if str(getattr(sl, "name", "") or "").startswith("AUTO_8D_ATTACH_")
    ]

    return (
        msg
        + f" / FINAL 상세 page={final_detail + 1}"
        + f" / FINAL 유첨 page={attach_pages}"
        + " / WRITER=BASELINE_DETAIL"
        + f" / 버전 저장={Path(saved).name}",
        saved,
    )


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

    # FINAL whole-file verification after attachments AND native-section COM.
    # This is the file the user actually opens.
    manifest = _assert_final_weekly_detail(saved)

    # Prefer the generated detail belonging to this run. If PowerPoint changed
    # the numeric position, report the actual final page number from the file.
    final_detail_pages = [i + 1 for i in manifest["details"]]
    final_attachment_pages = [i + 1 for i in manifest["attachments"]]

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
        + f" / FINAL 상세 page={final_detail_pages}"
        + f" / FINAL 유첨 page={final_attachment_pages}"
        + f" / FINAL 전체={manifest['slide_count']}p"
        + " / WRITER=SINGLE_PATH"
        + f" / 버전 저장={Path(saved).name}",
        saved,
    )


# Absolute last assignment: every GUI/core alias uses the previously working
# weekly_fix5 detail creator plus our narrow summary/native-section fixes.
for _runtime_base in (
    core.base,
    s14.base,
    s13.base,
    enterprise_main.base,
    legacy_final.base,
    final_polish.base,
):
    _runtime_base.weekly = _weekly_baseline_detail_path


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


# ---------------------------------------------------------------------------
# Parenthesized project confirmation from the ACTUAL weekly-summary page.
#
# User rule:
# - normal projects: keep the old/simple automatic matcher untouched.
# - projects containing (...): before execution, inspect the weekly summary,
#   show same-base candidates, and let the user choose the actual summary label.
# ---------------------------------------------------------------------------
def _parenless_identity(value):
    return s13._k(re.sub(r"\([^()]*\)", "", N(value)))


def _catalog_exact_name_for_summary_label(label, team=""):
    """Map a visible weekly-summary label back to the user's project catalog."""
    raw_q = s13._k(label)
    if not raw_q:
        return ""

    pool = list(catalog.PROJECTS.get(N(team), ()) or ())
    seen = set(pool)
    for item in catalog._all_projects():
        if item not in seen:
            pool.append(item)
            seen.add(item)

    exact = []
    contained = []
    for item in pool:
        iq = s13._k(item)
        pq = s13._k(catalog.project_part(item))
        if iq and raw_q == iq:
            exact.append(item)
        elif pq and raw_q == pq:
            exact.append(item)
        elif iq and len(iq) >= 5 and iq in raw_q:
            contained.append(item)
        elif pq and len(pq) >= 5 and pq in raw_q:
            contained.append(item)

    if exact:
        return exact[0]
    if contained:
        # Prefer the longest identity; it is least likely to be a partial model.
        contained.sort(key=lambda x: len(s13._k(x)), reverse=True)
        return contained[0]
    return ""


def _parenthesized_weekly_candidates(weekly_path, selected, team="", customer=""):
    """Return visible task labels whose BASE project matches when () is ignored."""
    if not selected or not re.search(r"\([^()]+\)", N(selected)):
        return []

    from pptx import Presentation

    selected_catalog = _catalog_separator_alias(selected) or N(selected)
    selected_customer, selected_project = catalog.split_customer_task(selected_catalog)
    if not selected_project:
        selected_project = selected_catalog

    project_base = _parenless_identity(selected_project)
    full_base = _parenless_identity(
        (selected_customer or N(customer)) + "_" + selected_project
    )
    if not project_base:
        return []

    prs = Presentation(weekly_path)
    found = []
    seen = set()

    for sl in prs.slides:
        for tb, hr, hm in _simple_summary_table_candidates(sl):
            task_col = hm.get("task")
            if task_col is None:
                continue
            for r in range(hr + 1, len(tb.rows)):
                raw = N(s13._row_text(tb, r, task_col))
                if not raw:
                    continue

                raw_base = _parenless_identity(raw)
                if not raw_base:
                    continue

                same_base = (
                    raw_base == project_base
                    or raw_base == full_base
                    or (
                        len(project_base) >= 4
                        and (
                            raw_base.endswith(project_base)
                            or project_base in raw_base
                        )
                    )
                )
                if not same_base:
                    continue

                key = s13._k(raw)
                if key in seen:
                    continue
                seen.add(key)

                canonical = _catalog_exact_name_for_summary_label(raw, team)
                found.append(
                    {
                        "display": raw,
                        "canonical": canonical or raw.replace("\n", " ").strip(),
                        "exact": s13._k(raw) == s13._k(selected_catalog),
                    }
                )

    # Exact same qualifier/spelling first; otherwise preserve weekly page order.
    found.sort(key=lambda x: (0 if x["exact"] else 1))
    return found


def _choose_parenthesized_weekly_candidate(self, selected, candidates):
    if not candidates:
        return "keep", selected

    if len(candidates) == 1:
        candidate = candidates[0]
        msg = (
            "주간회의 요약페이지에서 유사한 과제명을 찾았습니다.\n\n"
            f"입력값          : {selected}\n"
            f"요약페이지 과제 : {candidate['display']}\n\n"
            "이 과제로 업데이트하시겠습니까?"
        )
        choice = ui.dialog(
            self,
            "주간회의 과제명 확인",
            msg,
            "question",
            (("입력값 유지", "keep"), ("이 과제로 업데이트", "use")),
            width=650,
            height=350,
            button_width=16,
        )
        if choice is None:
            return "cancel", selected
        if choice == "use":
            return "use", candidate["canonical"]
        return "keep", selected

    # More than one same-base candidate, e.g. EB-L(EU) / EB-L(US).
    win = tk.Toplevel(self)
    win.withdraw()
    win.title("주간회의 과제명 확인")
    win.configure(bg="white")
    win.transient(self)
    result = {"value": None}

    head = tk.Frame(win, bg=ui.NAVY, height=56)
    head.pack(fill="x")
    head.pack_propagate(False)
    tk.Label(
        head,
        text="주간회의 과제명 확인",
        bg=ui.NAVY,
        fg="white",
        font=("Malgun Gothic", 12, "bold"),
    ).pack(side="left", padx=22)

    body = tk.Frame(win, bg="white")
    body.pack(fill="both", expand=True, padx=22, pady=16)
    tk.Label(
        body,
        text=(
            "주간회의 요약페이지에서 유사한 과제명이 여러 개 확인되었습니다.\n"
            "업데이트할 과제를 선택해 주세요.\n\n"
            f"입력값 : {selected}"
        ),
        bg="white",
        fg=ui.TEXT,
        justify="left",
        anchor="w",
        font=("Malgun Gothic", 10),
    ).pack(fill="x", pady=(0, 10))

    lb = tk.Listbox(
        body,
        font=("Malgun Gothic", 10),
        height=min(8, len(candidates)),
        exportselection=False,
    )
    lb.pack(fill="both", expand=True)
    for item in candidates:
        lb.insert("end", item["display"].replace("\n", " / "))
    exact_index = next((i for i, x in enumerate(candidates) if x["exact"]), 0)
    lb.selection_set(exact_index)
    lb.activate(exact_index)

    foot = tk.Frame(win, bg="#F6F8FA", height=76)
    foot.pack(fill="x")
    foot.pack_propagate(False)
    box = tk.Frame(foot, bg="#F6F8FA")
    box.pack(side="right", padx=20, pady=14)

    def finish(value):
        result["value"] = value
        win.destroy()

    tk.Button(
        box,
        text="취소",
        command=lambda: finish(("cancel", selected)),
        bg="#E5EBF0",
        fg=ui.TEXT,
        bd=0,
        font=("Malgun Gothic", 9, "bold"),
        width=10,
        pady=9,
    ).pack(side="left", padx=4)
    tk.Button(
        box,
        text="입력값 유지",
        command=lambda: finish(("keep", selected)),
        bg="#E5EBF0",
        fg=ui.TEXT,
        bd=0,
        font=("Malgun Gothic", 9, "bold"),
        width=12,
        pady=9,
    ).pack(side="left", padx=4)

    def use_selected():
        sel = lb.curselection()
        if not sel:
            return
        item = candidates[int(sel[0])]
        finish(("use", item["canonical"]))

    tk.Button(
        box,
        text="선택한 과제 사용",
        command=use_selected,
        bg=ui.BLUE,
        fg="white",
        bd=0,
        font=("Malgun Gothic", 9, "bold"),
        width=14,
        pady=9,
    ).pack(side="left", padx=4)

    win.protocol("WM_DELETE_WINDOW", lambda: finish(("cancel", selected)))
    ui.center_window(win, self, 680, 430)
    win.deiconify()
    win.grab_set()
    win.focus_force()
    self.wait_window(win)
    return result["value"] or ("cancel", selected)


_original_enterprise_run_parenthetical_weekly = enterprise_main.EnterpriseApp.run


def _run_with_parenthesized_weekly_choice(self):
    g = self.gui()
    weekly = N(g.get("pptweekly"))
    selected = N(g.get("task_name"))

    # Only the small catalog/input subset containing parentheses gets this popup.
    if (
        weekly
        and Path(weekly).exists()
        and selected
        and re.search(r"\([^()]+\)", selected)
    ):
        try:
            candidates = _parenthesized_weekly_candidates(
                weekly,
                selected,
                N(g.get("team")),
                "",
            )
            action, chosen = _choose_parenthesized_weekly_candidate(
                self, selected, candidates
            )
            if action == "cancel":
                try:
                    self.status_var.set(
                        "READY · 주간회의 과제명 확인이 취소되었습니다."
                    )
                except Exception:
                    pass
                return
            if action == "use" and chosen and chosen != selected:
                try:
                    self.vars["task_name"].set(chosen)
                except Exception:
                    pass
        except Exception:
            # Candidate suggestion is optional. Never block the old/simple
            # normal execution merely because the preflight scanner failed.
            pass

    return _original_enterprise_run_parenthetical_weekly(self)


enterprise_main.EnterpriseApp.run = _run_with_parenthesized_weekly_choice
