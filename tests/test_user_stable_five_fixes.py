import io
import tempfile
from pathlib import Path
import unittest

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

import main_recovery_step12 as s12
import main_recovery_step13 as s13
import main_recovery_step14 as s14
import user_stable_weekly_fixes as fix
import main_recovery_step14_fix2 as core


def add_real_detail_slide(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    labels = (
        ("2D", "현상"),
        ("3D", "임시조치"),
        ("4D", "발생원인"),
        ("5D", "개선대책"),
        ("6D", "효과검증"),
    )
    for i, (marker, title) in enumerate(labels):
        y = 1.7 + i * 0.65
        sl.shapes.add_textbox(
            Inches(0.4), Inches(y), Inches(0.55), Inches(0.3)
        ).text = marker
        sl.shapes.add_textbox(
            Inches(1.0), Inches(y), Inches(1.5), Inches(0.3)
        ).text = title
    for i, text in enumerate(("Signal", "이슈기인", "발생단계")):
        sl.shapes.add_textbox(
            Inches(8.0), Inches(0.5 + i * 0.4), Inches(2.0), Inches(0.3)
        ).text = text
    sl.shapes.add_textbox(
        Inches(5.5), Inches(2.8), Inches(1.5), Inches(0.3)
    ).text = "유출원인"
    sl.shapes.add_textbox(
        Inches(5.5), Inches(5.0), Inches(1.5), Inches(0.3)
    ).text = "수평전개"
    return sl


def add_summary_slide(prs, task, issue="old", top=0.7, rows=3):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    tb = sl.shapes.add_table(
        rows, 5, Inches(0.5), Inches(top), Inches(9.5), Inches(1.4)
    ).table
    headers = ["과제명", "이슈명", "현상", "진행사항", "Signal"]
    for c, h in enumerate(headers):
        tb.cell(0, c).text = h
    tb.cell(1, 0).text = task
    tb.cell(1, 1).text = issue
    return sl, tb


class UserStableFiveFixesTest(unittest.TestCase):
    def test_required_owner_and_classification_fields_are_all_mandatory(self):
        complete = {
            "team": "파우치형Pack개발품질1팀",
            "task_name": "MBAG_EB-L(EU)",
            "owner": "홍길동",
            "sample": "B2",
            "plm_no": "",
            "xlsx": "",
            "form_factor": "파우치형",
            "product_type": "EV Pack",
            "occurrence_site": "제품 생산",
            "stage": "DV",
        }
        # Weekly-only: PLM number is optional here and is left to the original
        # Issue-DB-only V2 warning flow.
        self.assertIsNone(fix._required_user_input_error(complete))

        missing_owner = dict(complete)
        missing_owner["owner"] = ""
        self.assertEqual(
            fix._required_user_input_error(missing_owner),
            ("입력 확인", "담당자를 입력해 주세요."),
        )

        missing_stage = dict(complete)
        missing_stage["stage"] = ""
        title, message = fix._required_user_input_error(missing_stage)
        self.assertEqual(title, "분류 정보 확인")
        self.assertIn("개발 단계", message)

    def test_detail_title_exact_user_format(self):
        d = {"customer": "GM", "task_name": "GM_MBAG"}
        g = {
            "task_name": "GM_MBAG",
            "sample": "B2",
            "stage": "DV",
            "occurrence_site": "제품 생산",
        }
        self.assertEqual(
            fix._detail_title_user_format(d, g),
            "GM_MBAG_B2 샘플_DV_제품 생산 이슈 발생",
        )

    def test_summary_uses_last_exact_matching_page(self):
        prs = Presentation()
        add_summary_slide(prs, "GM_MBAG", "old-1")
        prs.slides.add_slide(prs.slide_layouts[6])
        add_summary_slide(prs, "GM_MBAG", "old-2")

        d = {
            "customer": "GM",
            "task_name": "GM_MBAG",
            "issue_name": "new",
        }
        g = {"task_name": "GM_MBAG"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb, row, hr, _d, _g):
                hm = s13._summary_map(tb, hr)
                tb.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb.cell(row, hm["issue"]).text = "NEW"
            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 2)
        self.assertIn("마지막", action)
        self.assertEqual(
            prs.slides[2].shapes[0].table.cell(row, 0).text,
            "GM_MBAG",
        )

    def test_summary_customer_project_ignores_separator_style(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "GM MBAG", "old")

        d = {
            "customer": "GM",
            "task_name": "MBAG",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, _ = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertEqual(tb.cell(row, 0).text, "GM MBAG")

    def test_mbag_underscore_and_no_underscore_are_same_summary_project(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "MBAGEB565M", "old")

        d = {
            "customer": "MBAG",
            "task_name": "MBAG_EB565M",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG_EB565M"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertIn("마지막", action)
        self.assertEqual(tb.cell(row, 0).text, "MBAGEB565M")

    def test_mbag_no_underscore_input_resolves_catalog_identity(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "MBAG_EB565M", "old")

        # Even if extracted customer metadata is stale/different, the explicit
        # project selection resolves against the catalog separator-insensitively.
        d = {
            "customer": "Mercedes",
            "task_name": "MBAGEB565M",
            "issue_name": "new",
        }
        g = {"task_name": "MBAGEB565M"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, _ = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertEqual(tb.cell(row, 0).text, "MBAG_EB565M")

    def test_mbag_separator_variants_are_exact_section_match(self):
        d1 = {"customer": "MBAG", "task_name": "MBAG_EB565M"}
        d2 = {"customer": "Mercedes", "task_name": "MBAGEB565M"}
        self.assertEqual(fix._section_match_level("MBAGEB565M", d1), 3)
        self.assertEqual(fix._section_match_level("MBAG_EB565M", d2), 3)

    def test_project_only_fallback_survives_customer_mismatch(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "GM MBAG", "old")

        # Step 1 must fail because customer metadata is intentionally different.
        # Step 2 must still match the exact project MBAG.
        d = {
            "customer": "OTHER",
            "task_name": "MBAG",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertIn("마지막", action)
        self.assertEqual(tb.cell(row, 0).text, "GM MBAG")

    def test_project_only_section_fallback_survives_customer_mismatch(self):
        d = {"customer": "OTHER", "task_name": "MBAG"}
        self.assertEqual(fix._section_match_level("GM MBAG", d), 2)
        self.assertEqual(fix._section_match_level("GM_MBAG", d), 2)

    def test_confirmed_similar_section_reuses_real_summary_label(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "GM_MBAG E~", "old")

        d = {
            "customer": "GM",
            "task_name": "GM_MBAG",
            "issue_name": "new",
        }
        g = {
            "task_name": "GM_MBAG",
            "_weekly_section_override_name": "GM_MBAG E~",
        }

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"
            s14._write_summary_row = writer
            si, row, _ = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertEqual(tb.cell(row, 0).text, "GM_MBAG E~")

    def test_similarity_without_confirmed_section_does_not_force_match(self):
        prs = Presentation()
        add_summary_slide(prs, "GM_MBAG E~", "old")

        d = {
            "customer": "GM",
            "task_name": "GM_MBAG",
            "issue_name": "new",
        }
        g = {"task_name": "GM_MBAG"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb, row, hr, _d, _g):
                hm = s13._summary_map(tb, hr)
                tb.cell(row, hm["task"]).text = s13._customer_task(_d)
            s14._write_summary_row = writer
            _, _, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertIn("신규 요약 페이지", action)

    def test_summary_overflow_clones_immediately_after_last_match(self):
        prs = Presentation()
        add_summary_slide(prs, "OTHER", "old")
        prs.slides.add_slide(prs.slide_layouts[6])
        add_summary_slide(prs, "GM_MBAG", "old")

        # Add trailing detail pages so "next page" is observably different
        # from "append to the end of the deck".
        tail1 = prs.slides.add_slide(prs.slide_layouts[6])
        tail1.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text = "TAIL-1"
        tail2 = prs.slides.add_slide(prs.slide_layouts[6])
        tail2.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text = "TAIL-2"

        d = {
            "customer": "GM",
            "task_name": "GM_MBAG",
            "issue_name": "new",
        }
        g = {"task_name": "GM_MBAG"}

        old_fit = s14._summary_row_insert_fits
        old_writer = s14._write_summary_row
        try:
            s14._summary_row_insert_fits = lambda *args, **kwargs: False

            def writer(tb, row, hr, _d, _g):
                hm = s13._summary_map(tb, hr)
                tb.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, _, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._summary_row_insert_fits = old_fit
            s14._write_summary_row = old_writer

        self.assertEqual(si, 3)
        self.assertIn("바로 다음 페이지", action)
        self.assertIn("TAIL-1", "\n".join(
            getattr(sh, "text", "") for sh in prs.slides[4].shapes
        ))
        self.assertIn("TAIL-2", "\n".join(
            getattr(sh, "text", "") for sh in prs.slides[5].shapes
        ))

    def test_summary_detector_ignores_decoy_project_table(self):
        prs = Presentation()
        decoy = prs.slides.add_slide(prs.slide_layouts[6])
        dt = decoy.shapes.add_table(
            3, 2, Inches(0.5), Inches(0.5), Inches(5), Inches(1.2)
        ).table
        dt.cell(0, 0).text = "과제명"
        dt.cell(0, 1).text = "이슈명"
        dt.cell(1, 0).text = "WRONG_TEMPLATE"
        dt.cell(1, 1).text = "x"

        real = prs.slides.add_slide(prs.slide_layouts[6])
        rt = real.shapes.add_table(
            3, 5, Inches(0.5), Inches(0.7), Inches(9.5), Inches(1.5)
        ).table
        headers = ("고객사/과제명", "이슈명", "현상", "진행사항", "Signal 상태")
        for col, h in enumerate(headers):
            rt.cell(0, col).text = h
        rt.cell(1, 0).text = "MBAG EB-L(EU)"
        rt.cell(1, 1).text = "old"

        pages = fix._summary_pages_flexible(prs)
        self.assertEqual([x[0] for x in pages], [1])

    def test_existing_task_page_can_match_even_if_header_variant_is_not_generic_template(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        tb = sl.shapes.add_table(
            4, 3, Inches(0.5), Inches(0.7), Inches(8.5), Inches(1.8)
        ).table
        # This older summary variant has only two recognized data columns,
        # so it is intentionally NOT a generic-template candidate.
        for col, h in enumerate(("과제명", "이슈명", "진행사항")):
            tb.cell(0, col).text = h
        tb.cell(1, 0).text = "MBAG EB-L(EU)"
        tb.cell(1, 1).text = "old"
        tb.cell(1, 2).text = "old progress"

        d = {
            "customer": "MBAG",
            "task_name": "MBAG_EB-L(EU)",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG_EB-L(EU)"}

        # Strict generic summary detection may exclude this page.
        self.assertEqual(fix._summary_pages_flexible(prs), [])

        _pages, hits = fix._summary_hits(prs, d, g)
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0][0], 0)
        self.assertEqual(hits[0][4], [1])

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"
            s14._write_summary_row = writer

            # _update_summary_by_task itself still needs at least one strict page
            # for generic new-page fallback. Add a different valid template page;
            # exact task matching must still choose slide 0, not that template.
            add_summary_slide(prs, "OTHER_PROJECT", "x")
            si, row, action = s14._update_summary_by_task(prs, d, g, "new")
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertIn("마지막", action)
        self.assertEqual(tb.cell(row, 0).text, "MBAG EB-L(EU)")

    def test_saved_blank_detail_shell_is_filled_before_section_creation(self):
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "weekly.pptx"
            prs = Presentation()
            sl = add_real_detail_slide(prs)
            request = {
                "name": "MBAG_EB-L(EU)",
                "slide_index": 0,
                "slide_id": int(sl.slide_id),
            }
            prs.save(out)

            old_update = s13._update_detail_slide
            try:
                def fake_update(slide, d, g, mode):
                    for i, key in enumerate(("2D", "3D", "4D_CAUSE", "5D")):
                        sh = slide.shapes.add_textbox(
                            Inches(3),
                            Inches(0.5 + i * 0.4),
                            Inches(2),
                            Inches(0.3),
                        )
                        sh.name = "AUTO_8D_TEXT_" + key
                        sh.text = key + " NEW"
                s13._update_detail_slide = fake_update

                idx, repaired = fix._ensure_saved_detail(
                    str(out),
                    {"customer": "MBAG", "task_name": "EB-L(EU)", "issue_name": "NEW"},
                    {},
                    "new",
                    request,
                )
            finally:
                s13._update_detail_slide = old_update

            check = Presentation(out)
            self.assertEqual(idx, 0)
            self.assertTrue(fix._filled_detail_slide(check.slides[0]))
            self.assertEqual(repaired["slide_index"], 0)

    def test_catalog_split_matcher_keeps_normal_projects_on_legacy_path(self):
        d = {"customer": "MBAG", "task_name": "MBAG_EB565M"}
        g = {"task_name": "MBAG_EB565M", "team": "파우치형Pack개발품질1팀"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)
        spec = fix._catalog_project_match_spec(d, g)

        self.assertEqual(spec["mode"], "normal")
        self.assertGreater(
            fix._weekly_task_match_rank(
                "MBAG\nEB565M", d, g, full_keys, project_keys, spec
            ),
            0,
        )
        self.assertGreater(
            fix._weekly_task_match_rank(
                "EB565M", d, g, full_keys, project_keys, spec
            ),
            0,
        )

    def test_catalog_parenthesis_projects_use_strict_qualified_lane(self):
        d = {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"}
        g = {"task_name": "MBAG_EB-L(EU)", "team": "원통형Pack개발품질팀"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)
        spec = fix._catalog_project_match_spec(d, g)

        self.assertEqual(spec["mode"], "qualified")
        self.assertEqual(spec["project"], "EB-L(EU)")
        self.assertGreater(
            fix._weekly_task_match_rank(
                "MBAG\nEB-L(EU)\n검토(2차)",
                d, g, full_keys, project_keys, spec
            ),
            0,
        )
        self.assertEqual(
            fix._weekly_task_match_rank(
                "MBAG\nEB-L(US)\n검토(2차)",
                d, g, full_keys, project_keys, spec
            ),
            0,
        )

    def test_summary_finder_uses_normal_and_parenthesis_lanes_separately(self):
        # Normal project: old matching behavior.
        prs = Presentation()
        _, tb = add_summary_slide(prs, "MBAG\nEB565M", "old")
        hit = fix._find_existing_summary_project(
            prs,
            {"customer": "MBAG", "task_name": "MBAG_EB565M"},
            {"task_name": "MBAG_EB565M", "team": "파우치형Pack개발품질1팀"},
        )
        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 0)

        # Parenthesized variants: only the matching qualifier may win.
        prs2 = Presentation()
        add_summary_slide(prs2, "MBAG EB-L(US)", "wrong")
        add_summary_slide(prs2, "MBAG\nEB-L(EU)", "right")
        hit2 = fix._find_existing_summary_project(
            prs2,
            {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
            {"task_name": "MBAG_EB-L(EU)", "team": "원통형Pack개발품질팀"},
        )
        self.assertIsNotNone(hit2)
        self.assertEqual(hit2[1], 1)

    def test_project_parentheses_rule_is_exact_for_input_project(self):
        expected = fix._project_parenthetical_parts(
            "MBAG_EB-L(EU)", "MBAG"
        )
        self.assertEqual(expected, ("eu",))

        self.assertTrue(
            fix._project_parentheses_match(
                "MBAG\nEB-L(EU)\n검토사항(2차)",
                expected,
            )
        )
        self.assertFalse(
            fix._project_parentheses_match(
                "MBAG\nEB-L(US)\n검토사항(2차)",
                expected,
            )
        )

    def test_summary_project_parentheses_rule_uses_project_not_other_notes(self):
        prs = Presentation()
        sl1 = prs.slides.add_slide(prs.slide_layouts[6])
        tb1 = sl1.shapes.add_table(
            3, 4, Inches(0.5), Inches(0.7), Inches(9), Inches(1.5)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항")):
            tb1.cell(0, col).text = h
        tb1.cell(1, 0).text = "MBAG\nEB-L(US)\n검토(2차)"
        tb1.cell(1, 1).text = "wrong"

        sl2 = prs.slides.add_slide(prs.slide_layouts[6])
        tb2 = sl2.shapes.add_table(
            3, 4, Inches(0.5), Inches(0.7), Inches(9), Inches(1.5)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항")):
            tb2.cell(0, col).text = h
        tb2.cell(1, 0).text = "MBAG EB-L(EU) (3차)"
        tb2.cell(1, 1).text = "right"

        hit = fix._find_existing_summary_project(
            prs,
            {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
            {"task_name": "MBAG_EB-L(EU)"},
        )
        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 1)

    def test_parenthetical_match_ignores_unrelated_parentheses_in_same_sentence(self):
        d = {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"}
        g = {"task_name": "MBAG_EB-L(EU)"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)

        # The project qualifier is (EU), but the same cell/page may contain an
        # unrelated note like (2차). That must not block the project match.
        raw = "과제 : MBAG\nEB-L(EU)\n검토사항(2차)"
        self.assertGreater(
            fix._simple_task_match_rank(raw, full_keys, project_keys),
            0,
        )

        # The inner qualifier text still distinguishes variants because _k()
        # keeps EU/US while removing only punctuation/separators.
        wrong = "과제 : MBAG\nEB-L(US)\n검토사항(2차)"
        self.assertEqual(
            fix._simple_task_match_rank(wrong, full_keys, project_keys),
            0,
        )

    def test_existing_summary_project_finds_matching_qualifier_with_extra_parentheses(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        tb = sl.shapes.add_table(
            3, 4, Inches(0.5), Inches(0.7), Inches(9), Inches(1.5)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항")):
            tb.cell(0, col).text = h
        tb.cell(1, 0).text = "MBAG\nEB-L(EU)\n(2차)"
        tb.cell(1, 1).text = "old"

        hit = fix._find_existing_summary_project(
            prs,
            {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
            {"task_name": "MBAG_EB-L(EU)"},
        )
        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 0)

    def test_parenthetical_project_variants_never_match_each_other(self):
        d = {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"}
        g = {"task_name": "MBAG_EB-L(EU)"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)
        required = fix._required_project_parens(d, g)

        self.assertGreater(
            fix._simple_task_match_rank(
                "MBAG EB-L(EU)", full_keys, project_keys, required
            ),
            0,
        )
        self.assertEqual(
            fix._simple_task_match_rank(
                "MBAG EB-L(US)", full_keys, project_keys, required
            ),
            0,
        )
        self.assertEqual(
            fix._simple_task_match_rank(
                "MBAG EB-L", full_keys, project_keys, required
            ),
            0,
        )

    def test_summary_search_chooses_exact_parenthetical_variant(self):
        prs = Presentation()
        add_summary_slide(prs, "MBAG EB-L(US)", "us-old")
        add_summary_slide(prs, "MBAG\nEB-L(EU)", "eu-old")

        hit = fix._find_existing_summary_project(
            prs,
            {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
            {"task_name": "MBAG_EB-L(EU)"},
        )

        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 1)
        self.assertEqual(hit[6], "MBAG\nEB-L(EU)")

    def test_summary_identity_matches_underscore_space_and_newline(self):
        d = {"customer": "A", "task_name": "B"}
        g = {"task_name": "B"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)

        self.assertEqual(
            fix._simple_task_match_rank("A_B", full_keys, project_keys),
            400,
        )
        self.assertEqual(
            fix._simple_task_match_rank("A B", full_keys, project_keys),
            400,
        )
        self.assertEqual(
            fix._simple_task_match_rank("A\nB", full_keys, project_keys),
            400,
        )
        self.assertEqual(
            fix._simple_task_match_rank("B", full_keys, project_keys),
            360,
        )

    def test_summary_project_only_sentence_match_is_same_project(self):
        d = {
            "customer": "MBAG",
            "task_name": "MBAG_EB-L(EU)",
        }
        g = {"task_name": "MBAG_EB-L(EU)"}
        full_keys, project_keys = fix._summary_identity_keys(d, g)

        self.assertGreaterEqual(
            fix._simple_task_match_rank(
                "주간 현황\nEB-L(EU)\n신규 이슈",
                full_keys,
                project_keys,
            ),
            280,
        )
        self.assertGreaterEqual(
            fix._simple_task_match_rank(
                "MBAG\nEB-L(EU)",
                full_keys,
                project_keys,
            ),
            320,
        )

    def test_project_only_cell_can_drive_existing_summary_page(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "EB-L(EU)", "old")
        d = {
            "customer": "MBAG",
            "task_name": "MBAG_EB-L(EU)",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG_EB-L(EU)"}

        hit = fix._find_existing_summary_project(prs, d, g)
        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 0)
        self.assertEqual(hit[6], "EB-L(EU)")

    def test_page_level_project_match_finds_old_summary_even_when_task_cell_is_blank(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        tb = sl.shapes.add_table(
            4, 4, Inches(0.5), Inches(0.7), Inches(9), Inches(1.7)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항")):
            tb.cell(0, col).text = h
        # Simulate a legacy merged/template layout where python-pptx exposes
        # the project text as a separate visible shape instead of the row cell.
        tb.cell(1, 1).text = "old issue"
        tb.cell(1, 2).text = "old problem"
        sl.shapes.add_textbox(
            Inches(0.6), Inches(0.2), Inches(3), Inches(0.35)
        ).text = "MBAG EB-L(EU)"

        hit = fix._find_existing_summary_project(
            prs,
            {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
            {"task_name": "MBAG_EB-L(EU)"},
        )
        self.assertIsNotNone(hit)
        self.assertEqual(hit[1], 0)

    def test_schedule_gate_page_is_never_detail_template(self):
        prs = Presentation()
        schedule = prs.slides.add_slide(prs.slide_layouts[6])
        for i, text in enumerate(
            (
                "CV 25/06/25",
                "DV 25/10/25",
                "PD/PV gate 일정 조정",
                "SOP Target Now",
                "4D 원인 분석",
                "5D 개선 일정",
            )
        ):
            schedule.shapes.add_textbox(
                Inches(0.5), Inches(0.5 + i * 0.45), Inches(4.5), Inches(0.3)
            ).text = text

        detail = add_real_detail_slide(prs)

        self.assertFalse(
            fix._strict_detail_template_fingerprint(schedule)["ok"]
        )
        self.assertTrue(
            fix._strict_detail_template_fingerprint(detail)["ok"]
        )
        self.assertEqual(
            fix._template_detail_index_any_section(
                prs, {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"}
            ),
            1,
        )

    def test_detail_clone_cleanup_removes_red_annotations_and_grouped_old_text(self):
        prs = Presentation()
        sl = add_real_detail_slide(prs)

        # Old red selection circle outside the D content zones.
        red = sl.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(8.2),
            Inches(0.6),
            Inches(0.7),
            Inches(0.35),
        )
        red.fill.background()
        red.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)

        # Simulate a preserved 4D group that also carries old issue body text.
        group = sl.shapes.add_group_shape()
        fixed = group.shapes.add_textbox(
            Inches(0.2), Inches(0.2), Inches(1.4), Inches(0.35)
        )
        fixed.text = "4D 발생원인"
        old = group.shapes.add_textbox(
            Inches(1.7), Inches(0.2), Inches(3.0), Inches(0.5)
        )
        old.text = "기존 이슈 검정 본문은 삭제되어야 함"

        fix._purge_cloned_detail_artifacts(sl)

        all_text = "\n".join(
            str(getattr(sh, "text", "") or "")
            for sh in fix.v310.walk(sl)
        )
        self.assertIn("4D 발생원인", all_text)
        self.assertNotIn("기존 이슈 검정 본문은 삭제되어야 함", all_text)

        # No strong-red old annotation should remain.
        self.assertFalse(
            any(fix._is_red_annotation_shape(sh) for sh in fix.v310.walk(sl))
        )

    def test_detail_cleanup_keeps_fixed_markers_titles(self):
        prs = Presentation()
        sl = add_real_detail_slide(prs)
        before = fix._strict_detail_template_fingerprint(sl)
        self.assertTrue(before["ok"])

        fix._purge_cloned_detail_artifacts(sl)

        after = fix._strict_detail_template_fingerprint(sl)
        self.assertTrue(after["ok"])
        self.assertGreaterEqual(len(after["markers"]), 4)

    def test_new_detail_text_is_blue_even_in_existing_mode(self):
        prs = Presentation()
        sl = add_real_detail_slide(prs)

        old_inner = fix._detail_update_before_blue_fix
        try:
            def fake_inner(slide, d, g, mode):
                sh = slide.shapes.add_textbox(
                    Inches(3), Inches(1.0), Inches(3), Inches(0.5)
                )
                sh.name = "AUTO_8D_TEXT_2D"
                sh.text = "새 상세 내용"

            fix._detail_update_before_blue_fix = fake_inner
            fix._update_detail_slide_force_new_text_blue(
                sl, {}, {}, "existing"
            )
        finally:
            fix._detail_update_before_blue_fix = old_inner

        auto = next(
            sh for sh in sl.shapes
            if str(getattr(sh, "name", "")) == "AUTO_8D_TEXT_2D"
        )
        colors = [
            run.font.color.rgb
            for p in auto.text_frame.paragraphs
            for run in p.runs
            if run.text
        ]
        self.assertTrue(colors)
        self.assertTrue(all(x == RGBColor(0x00, 0x33, 0xFF) for x in colors))

    def test_baseline_detail_writer_is_pre_polish_weekly_writer(self):
        self.assertIs(fix._known_good_detail_writer, fix.final_polish._original_weekly)

    def test_final_full_manifest_requires_real_detail(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            p = td / "final.pptx"

            prs = Presentation()
            detail = add_real_detail_slide(prs)
            for i, key in enumerate(("2D", "3D", "4D_CAUSE", "5D")):
                sh = detail.shapes.add_textbox(
                    Inches(3), Inches(0.5 + i * 0.4), Inches(2), Inches(0.3)
                )
                sh.name = "AUTO_8D_TEXT_" + key
                sh.text = "generated " + key
            prs.save(p)

            manifest = fix._assert_final_weekly_detail(str(p))
            self.assertEqual(manifest["details"], [0])

            p2 = td / "attach_only.pptx"
            bad = Presentation()
            sl = bad.slides.add_slide(bad.slide_layouts[6])
            try:
                sl.name = "AUTO_8D_ATTACH_TEST_1"
            except Exception:
                sl._element.cSld.set("name", "AUTO_8D_ATTACH_TEST_1")
            bad.save(p2)

            with self.assertRaises(RuntimeError):
                fix._assert_final_weekly_detail(str(p2))

    def test_all_runtime_weekly_aliases_point_to_baseline_detail_path(self):
        self.assertIs(core.base.weekly, fix._weekly_baseline_detail_path)
        self.assertIs(s14.base.weekly, fix._weekly_baseline_detail_path)
        self.assertIs(s13.base.weekly, fix._weekly_baseline_detail_path)
        self.assertIs(fix.enterprise_main.base.weekly, fix._weekly_baseline_detail_path)
        self.assertIs(fix.legacy_final.base.weekly, fix._weekly_baseline_detail_path)

    def test_reduced_output_never_drops_marked_detail_before_attachments(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            source = td / "source.pptx"
            saved = td / "saved.pptx"

            src = Presentation()
            base_slide = src.slides.add_slide(src.slide_layouts[6])
            base_slide.shapes.add_textbox(
                Inches(0.5), Inches(0.5), Inches(2), Inches(0.4)
            ).text = "UNCHANGED"
            src.save(source)

            dst = Presentation(source)
            detail = dst.slides.add_slide(dst.slide_layouts[6])
            detail.shapes.add_textbox(
                Inches(0.5), Inches(0.5), Inches(2), Inches(0.4)
            ).text = "DETAIL"
            fix._mark_detail_slide(
                detail, {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)", "issue_name": "X"}
            )

            attach = dst.slides.add_slide(dst.slide_layouts[6])
            attach.shapes.add_textbox(
                Inches(0.5), Inches(0.5), Inches(2), Inches(0.4)
            ).text = "ATTACH"
            try:
                attach.name = "AUTO_8D_ATTACH_TEST_1"
            except Exception:
                attach._element.cSld.set("name", "AUTO_8D_ATTACH_TEST_1")
            dst.save(saved)

            reduced, count = fix._ppt_update_only_keep_current_detail(
                str(source), str(saved)
            )
            check = Presentation(reduced)

            self.assertTrue(any(fix._is_marked_detail(sl) for sl in check.slides))
            self.assertTrue(
                any(
                    str(getattr(sl, "name", "") or "").startswith("AUTO_8D_ATTACH_")
                    for sl in check.slides
                )
            )
            self.assertGreaterEqual(count, 2)

    def test_single_path_always_creates_filled_detail_before_attachments(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            weekly = td / "weekly.pptx"
            source8d = td / "source8d.pptx"
            out = td / "weekly_update.pptx"

            prs = Presentation()
            add_summary_slide(prs, "MBAG EB-L(EU)", "old")
            detail = prs.slides.add_slide(prs.slide_layouts[6])
            for i, token in enumerate(("2D", "3D", "4D", "5D", "6D", "7D")):
                detail.shapes.add_textbox(
                    Inches(0.4), Inches(2.0 + i * 0.55), Inches(0.7), Inches(0.3)
                ).text = token
                detail.shapes.add_textbox(
                    Inches(1.1), Inches(2.0 + i * 0.55), Inches(1.6), Inches(0.3)
                ).text = token + " 항목"
            detail.shapes.add_textbox(
                Inches(8.5), Inches(0.5), Inches(2), Inches(0.3)
            ).text = "Signal"
            detail.shapes.add_textbox(
                Inches(8.5), Inches(0.9), Inches(2), Inches(0.3)
            ).text = "이슈기인"
            detail.shapes.add_textbox(
                Inches(8.5), Inches(1.3), Inches(2), Inches(0.3)
            ).text = "발생단계"
            prs.save(weekly)

            src = Presentation()
            src.slides.add_slide(src.slide_layouts[6])
            src.save(source8d)

            d = {
                "customer": "MBAG",
                "task_name": "MBAG_EB-L(EU)",
                "issue_name": "신규 시험 이슈",
                "problem": "문제 현상",
                "temporary_action": "임시조치",
                "cause_4d": "발생원인",
                "leak_cause": "유출원인",
                "action_5d": "개선대책",
                "verification_6d": "효과검증",
            }
            g = {
                "ppt8d": str(source8d),
                "task_name": "MBAG_EB-L(EU)",
                "team": "원통형Pack개발품질팀",
                "owner": "홍길동",
                "sample": "B2",
                "stage": "DV",
                "occurrence_site": "제품 생산",
                "_weekly_create_new_section": "1",
            }

            old_section = fix._create_native_section_com_saved
            try:
                fix._create_native_section_com_saved = lambda *args, **kwargs: True
                msg, saved = fix._weekly_single_path(
                    str(weekly), str(out), d, g, "new"
                )
            finally:
                fix._create_native_section_com_saved = old_section

            result = Presentation(saved)
            filled = [
                i for i, slide in enumerate(result.slides)
                if fix._verified_generated_detail(slide)
            ]
            self.assertTrue(filled)
            self.assertIn("기존 과제 [MBAG EB-L(EU)]", msg)
            self.assertIn("상세 신규 생성", msg)

    def test_summary_direct_match_beats_all_fallback_logic(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "MBAG EB-L(EU)", "old")
        d = {
            "customer": "WRONG_CUSTOMER",
            "task_name": "MBAG_EB-L(EU)",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG_EB-L(EU)"}

        pages, hits = fix._summary_hits(prs, d, g)
        self.assertEqual(len(pages), 1)
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0][0], 0)
        self.assertEqual(hits[0][4], [1])

    def test_combined_detail_marker_title_is_preserved(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        sh = sl.shapes.add_textbox(
            Inches(0.4), Inches(2.2), Inches(2.5), Inches(0.4)
        )
        sh.text = "2D 현상"
        self.assertTrue(fix._static_cloned_detail_shape_preserve_layout(sh))

    def test_clone_without_safe_section_records_new_section_request(self):
        prs = Presentation()
        src = add_real_detail_slide(prs)

        old_clone = fix._original_clone_detail_shell
        old_pending = fix._pending_user_native_section
        try:
            def fake_clone(_prs, _d, insert_at, matched_section=None):
                s13._clone_slide_with_rels(_prs, 0)
                s13._move_last_slide_to(_prs, insert_at)
                return _prs.slides[insert_at], 0, matched_section

            fix._original_clone_detail_shell = fake_clone
            fix._pending_user_native_section = None
            fix._clone_detail_shell_clean(
                prs,
                {"customer": "MBAG", "task_name": "MBAG_EB-L(EU)"},
                1,
                None,
            )
            req = fix._pending_user_native_section
        finally:
            fix._original_clone_detail_shell = old_clone
            fix._pending_user_native_section = old_pending

        self.assertIsNotNone(req)
        self.assertEqual(req["slide_index"], 1)
        self.assertEqual(req["name"], "MBAG_EB-L(EU)")

    def test_mbag_ebl_space_and_underscore_match_exact_summary(self):
        prs = Presentation()
        _, tb = add_summary_slide(prs, "MBAG EB-L(EU)", "old")
        d = {
            "customer": "MBAG",
            "task_name": "MBAG_EB-L(EU)",
            "issue_name": "new",
        }
        g = {"task_name": "MBAG_EB-L(EU)"}

        old_writer = s14._write_summary_row
        try:
            def writer(tb0, row, hr, _d, _g):
                hm = s13._summary_map(tb0, hr)
                tb0.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb0.cell(row, hm["issue"]).text = "NEW"
            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._write_summary_row = old_writer

        self.assertEqual(si, 0)
        self.assertIn("마지막", action)
        self.assertEqual(tb.cell(row, 0).text, "MBAG EB-L(EU)")

    def test_cloned_summary_does_not_require_exact_signal_header(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        tb = sl.shapes.add_table(
            3, 5, Inches(0.5), Inches(0.7), Inches(9.5), Inches(1.5)
        ).table
        headers = ("과제명", "이슈명", "현상", "진행사항", "Signal 상태")
        for col, h in enumerate(headers):
            tb.cell(0, col).text = h
        tb.cell(1, 0).text = "MBAG EB-L(EU)"
        tb.cell(1, 1).text = "old"

        pages = fix._summary_pages_flexible(prs)
        self.assertEqual(len(pages), 1)

        si, new_tb, hr, row = fix._prepare_new_summary_page_stable(
            prs, pages, {"team": "Pack개발품질1"}, template_index=0
        )

        self.assertEqual(si, 1)
        self.assertEqual(hr, 0)
        self.assertGreaterEqual(row, 1)
        hm = s13._summary_map(new_tb, hr)
        self.assertIn("task", hm)
        self.assertIn("issue", hm)

    def test_summary_project_block_with_blank_task_continuation(self):
        prs = Presentation()

        sl1 = prs.slides.add_slide(prs.slide_layouts[6])
        tb1 = sl1.shapes.add_table(
            4, 5, Inches(0.5), Inches(0.7), Inches(9.5), Inches(1.8)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항", "Signal")):
            tb1.cell(0, col).text = h
        tb1.cell(1, 0).text = "MBAG_EB565M"
        tb1.cell(1, 1).text = "old-1"
        tb1.cell(2, 0).text = ""
        tb1.cell(2, 1).text = "old-2"

        sl2 = prs.slides.add_slide(prs.slide_layouts[6])
        tb2 = sl2.shapes.add_table(
            4, 5, Inches(0.5), Inches(0.7), Inches(9.5), Inches(1.8)
        ).table
        for col, h in enumerate(("과제명", "이슈명", "현상", "진행사항", "Signal")):
            tb2.cell(0, col).text = h
        tb2.cell(1, 0).text = "MBAG EB565M"
        tb2.cell(1, 1).text = "old-last-1"
        tb2.cell(2, 0).text = ""
        tb2.cell(2, 1).text = "old-last-2"

        tail = prs.slides.add_slide(prs.slide_layouts[6])
        tail.shapes.add_textbox(
            Inches(1), Inches(1), Inches(2), Inches(1)
        ).text = "TAIL"

        d = {
            "customer": "MBAG",
            "task_name": "MBAGEB565M",
            "issue_name": "new",
        }
        g = {"task_name": "MBAGEB565M"}

        old_fit = s14._summary_row_insert_fits
        old_writer = s14._write_summary_row
        try:
            s14._summary_row_insert_fits = lambda *args, **kwargs: False

            def writer(tb, row, hr, _d, _g):
                hm = s13._summary_map(tb, hr)
                tb.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._summary_row_insert_fits = old_fit
            s14._write_summary_row = old_writer

        self.assertEqual(si, 2)
        self.assertIn("바로 다음 페이지", action)
        new_tb = next(
            sh.table for sh in prs.slides[2].shapes
            if getattr(sh, "has_table", False)
        )
        hm = s13._summary_map(new_tb, 0)
        self.assertEqual(new_tb.cell(row, hm["task"]).text, "MBAG EB565M")
        self.assertIn(
            "TAIL",
            "\n".join(getattr(sh, "text", "") for sh in prs.slides[3].shapes),
        )

    def test_mixed_separator_last_page_overflow_preserves_existing_label(self):
        prs = Presentation()
        add_summary_slide(prs, "MBAG_EB565M", "old-1")
        tail0 = prs.slides.add_slide(prs.slide_layouts[6])
        tail0.shapes.add_textbox(
            Inches(1), Inches(1), Inches(2), Inches(1)
        ).text = "BETWEEN"
        _, last_tb = add_summary_slide(prs, "MBAG EB565M", "old-2")
        tail1 = prs.slides.add_slide(prs.slide_layouts[6])
        tail1.shapes.add_textbox(
            Inches(1), Inches(1), Inches(2), Inches(1)
        ).text = "TAIL"

        d = {
            "customer": "MBAG",
            "task_name": "MBAGEB565M",
            "issue_name": "new",
        }
        g = {"task_name": "MBAGEB565M"}

        old_fit = s14._summary_row_insert_fits
        old_writer = s14._write_summary_row
        try:
            s14._summary_row_insert_fits = lambda *args, **kwargs: False

            def writer(tb, row, hr, _d, _g):
                hm = s13._summary_map(tb, hr)
                tb.cell(row, hm["task"]).text = s13._customer_task(_d)
                tb.cell(row, hm["issue"]).text = "NEW"

            s14._write_summary_row = writer
            si, row, action = s14._update_summary_by_task(
                prs, d, g, "new"
            )
        finally:
            s14._summary_row_insert_fits = old_fit
            s14._write_summary_row = old_writer

        # The last matching summary page was index 2, so continuation must be 3.
        self.assertEqual(si, 3)
        self.assertIn("바로 다음 페이지", action)

        # The copied continuation must keep the exact label spelling/style from
        # the last existing occurrence, not rewrite it as the input spelling.
        new_tb = next(
            sh.table for sh in prs.slides[3].shapes
            if getattr(sh, "has_table", False)
        )
        hm = s13._summary_map(new_tb, 0)
        self.assertEqual(new_tb.cell(row, hm["task"]).text, "MBAG EB565M")

        # The unrelated trailing page must remain after the inserted continuation.
        self.assertIn(
            "TAIL",
            "\n".join(getattr(sh, "text", "") for sh in prs.slides[4].shapes),
        )

    def test_existing_detail_requires_same_issue_not_only_same_project(self):
        prs = Presentation()
        sl = add_real_detail_slide(prs)
        sl.shapes.add_textbox(
            Inches(3), Inches(0.5), Inches(4), Inches(0.4)
        ).text = "GM_MBAG old issue"

        d = {
            "customer": "GM",
            "task_name": "GM_MBAG",
            "issue_name": "completely new issue",
        }
        self.assertIsNone(fix._find_existing_detail_exact_issue(prs, d))

        sl.shapes.add_textbox(
            Inches(3), Inches(1.0), Inches(4), Inches(0.4)
        ).text = "completely new issue"
        s13._clear_slide_text_cache()
        self.assertEqual(fix._find_existing_detail_exact_issue(prs, d), 0)

    def test_detail_position_uses_last_existing_project_page(self):
        prs = Presentation()
        for label in ("OTHER", "GM_MBAG first", "GM_MBAG last", "TAIL"):
            sl = add_real_detail_slide(prs)
            sl.shapes.add_textbox(
                Inches(3), Inches(0.5), Inches(4), Inches(0.5)
            ).text = label

        d = {"customer": "GM", "task_name": "GM_MBAG"}
        self.assertEqual(core._new_detail_position(prs, d), 3)

    def test_attachment_safe_retries_direct_when_dedupe_adds_nothing(self):
        old_count = fix._count_source_attachment_slides
        old_safe = fix._original_attachment_safe
        old_has = fix._has_current_auto_attachment
        old_direct = fix._direct_attachment_inserter
        called = {"direct": 0}
        try:
            fix._count_source_attachment_slides = lambda _p: 2
            fix._original_attachment_safe = lambda *args, **kwargs: (0, "")
            fix._has_current_auto_attachment = lambda *args, **kwargs: False

            def direct(*args, **kwargs):
                called["direct"] += 1
                return 2

            fix._direct_attachment_inserter = lambda: direct
            added, err = core._append_8d_attachments_safe(
                "out.pptx", "src.pptx", 4, {"task_name": "GM_MBAG"}
            )
        finally:
            fix._count_source_attachment_slides = old_count
            fix._original_attachment_safe = old_safe
            fix._has_current_auto_attachment = old_has
            fix._direct_attachment_inserter = old_direct

        self.assertEqual(added, 2)
        self.assertEqual(err, "")
        self.assertEqual(called["direct"], 1)

    def test_mixed_summary_section_forces_new_detail_section(self):
        prs = Presentation()
        add_summary_slide(prs, "GM_MBAG", "old")
        detail = prs.slides.add_slide(prs.slide_layouts[6])
        for i, token in enumerate(("2D", "3D", "4D", "5D", "6D")):
            detail.shapes.add_textbox(
                Inches(0.5), Inches(0.5 + i * 0.4), Inches(2), Inches(0.3)
            ).text = token

        mixed = {"name": "기존 혼합구역", "indices": [0, 1]}
        old_selected = fix._original_selected_section
        try:
            fix._original_selected_section = lambda _prs, _d, _g: mixed
            g = {}
            selected = fix._selected_section_without_mixed_summary(
                prs, {"customer": "GM", "task_name": "GM_MBAG"}, g
            )
        finally:
            fix._original_selected_section = old_selected

        self.assertIsNone(selected)
        self.assertEqual(g.get("_weekly_create_new_section"), "1")
        self.assertEqual(
            g.get("_weekly_mixed_section_recovered"), "기존 혼합구역"
        )

    def test_detail_template_is_found_even_when_it_shares_summary_section(self):
        prs = Presentation()
        add_summary_slide(prs, "GM_MBAG", "old")
        detail = prs.slides.add_slide(prs.slide_layouts[6])
        for i, token in enumerate(
            ("2D", "3D", "4D", "5D", "6D", "7D", "Signal", "이슈기인", "발생단계")
        ):
            detail.shapes.add_textbox(
                Inches(0.5),
                Inches(0.4 + i * 0.35),
                Inches(2.5),
                Inches(0.3),
            ).text = token
        detail.shapes.add_textbox(
            Inches(3.5), Inches(0.5), Inches(3), Inches(0.4)
        ).text = "GM_MBAG"

        self.assertEqual(
            fix._template_detail_index_any_section(
                prs, {"customer": "GM", "task_name": "GM_MBAG"}
            ),
            1,
        )

    def test_pending_native_section_does_not_write_fake_xml(self):
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        before = list(core._native_sections(prs))
        sec = fix._create_native_section_pending(prs, "GM_MBAG", 0)
        after = list(core._native_sections(prs))

        self.assertEqual(before, after)
        self.assertTrue(sec["_pending_com"])
        self.assertIsNone(sec["element"])
        self.assertEqual(sec["indices"], [0])

    def test_pending_detail_is_relocated_by_slide_id_after_position_shift(self):
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "shifted.pptx"
            prs = Presentation()

            summary = prs.slides.add_slide(prs.slide_layouts[6])
            summary.shapes.add_textbox(
                Inches(0.5), Inches(0.5), Inches(2), Inches(0.4)
            ).text = "SUMMARY"

            detail = add_real_detail_slide(prs)
            detail.shapes.add_textbox(
                Inches(3), Inches(0.5), Inches(3), Inches(0.4)
            ).text = "GM_MBAG"

            request = {
                "name": "GM_MBAG",
                "slide_index": 1,
                "slide_id": int(detail.slide_id),
            }

            # Simulate a post-create slide movement: old numeric index now points
            # to another slide, but PowerPoint slide-id remains the same.
            tail = prs.slides.add_slide(prs.slide_layouts[6])
            tail.shapes.add_textbox(
                Inches(0.5), Inches(0.5), Inches(2), Inches(0.4)
            ).text = "TAIL"
            fix._move_slide_by_id(prs, int(detail.slide_id), 2)
            prs.save(out)

            found = fix._verify_pending_detail_slide(
                str(out), request, {"customer": "GM", "task_name": "GM_MBAG"}
            )

        self.assertEqual(found, 2)

    def test_new_section_wrapper_requires_and_verifies_detail_slide(self):
        old_active = fix._active_weekly_before_native_section
        old_com = fix._create_native_section_com_saved
        calls = {"com": 0}
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "weekly.pptx"

            def fake_active(src, out_path, d, g, mode):
                prs = Presentation()
                sl = add_real_detail_slide(prs)
                for i, key in enumerate(("2D", "3D", "4D_CAUSE", "5D")):
                    sh = sl.shapes.add_textbox(
                        Inches(3.0),
                        Inches(0.8 + i * 0.35),
                        Inches(2.0),
                        Inches(0.3),
                    )
                    sh.name = "AUTO_8D_TEXT_" + key
                    sh.text = "generated " + key
                prs.save(out_path)
                fix._create_native_section_pending(prs, "GM_MBAG", 0)
                return "ok", str(out_path)

            def fake_com(saved, name, slide_index):
                calls["com"] += 1
                self.assertEqual(name, "GM_MBAG")
                self.assertEqual(slide_index, 0)
                return True

            try:
                fix._active_weekly_before_native_section = fake_active
                fix._create_native_section_com_saved = fake_com
                msg, saved = fix._weekly_with_real_native_section(
                    "src.pptx",
                    str(out),
                    {"task_name": "GM_MBAG"},
                    {"_weekly_create_new_section": "1"},
                    "new",
                )
            finally:
                fix._active_weekly_before_native_section = old_active
                fix._create_native_section_com_saved = old_com

        self.assertEqual(calls["com"], 1)
        self.assertIn("상세 page 1 생성 확인", msg)

    def test_section_match_levels_are_strict(self):
        d = {"customer": "GM", "task_name": "GM_MBAG"}
        self.assertEqual(fix._section_match_level("GM_MBAG", d), 3)
        self.assertEqual(fix._section_match_level("MBAG", d), 2)
        self.assertEqual(fix._section_match_level("GM_MBAG E~", d), 1)

    def test_clone_visual_cleanup_preserves_table_and_frame(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        sl.shapes.add_table(
            2, 2, Inches(0.5), Inches(0.5), Inches(3), Inches(1)
        )
        sl.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.3),
            Inches(0.3),
            Inches(9),
            Inches(5),
        )

        image = Image.new("RGB", (20, 20), "white")
        bio = io.BytesIO()
        image.save(bio, format="PNG")
        bio.seek(0)
        sl.shapes.add_picture(
            bio, Inches(1), Inches(2), Inches(1), Inches(1)
        )

        before = len(sl.shapes)
        fix._remove_copied_visuals(sl)
        after = len(sl.shapes)

        self.assertEqual(after, before - 1)
        self.assertTrue(
            any(getattr(sh, "has_table", False) for sh in sl.shapes)
        )

    def test_7d_moves_below_6d_absolute_target(self):
        prs = Presentation()
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        marker = sl.shapes.add_textbox(
            Inches(0.5), Inches(2.0), Inches(0.5), Inches(0.3)
        )
        marker.text = "7D"
        title = sl.shapes.add_textbox(
            Inches(1.1), Inches(2.0), Inches(1.2), Inches(0.3)
        )
        title.text = "수평전개"

        old_layout = s12.step11._adaptive_cascade_layout
        try:
            s12.step11._adaptive_cascade_layout = lambda d, imgs: (
                {"6D": {"x": 5.5, "y": 4.0, "w": 5.0, "h": 1.0}},
                {},
                {},
            )
            fix._adjust_7d_after_6d(sl, {})
        finally:
            s12.step11._adaptive_cascade_layout = old_layout

        self.assertGreater(float(marker.top) / fix.v310.EMU, 5.0)
        self.assertAlmostEqual(
            float(marker.top) / fix.v310.EMU,
            float(title.top) / fix.v310.EMU,
            places=2,
        )


if __name__ == "__main__":
    unittest.main()
