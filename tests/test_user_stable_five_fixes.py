import io
import tempfile
from pathlib import Path
import unittest

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

import main_recovery_step12 as s12
import main_recovery_step13 as s13
import main_recovery_step14 as s14
import user_stable_weekly_fixes as fix
import main_recovery_step14_fix2 as core


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

    def test_detail_position_uses_last_existing_project_page(self):
        prs = Presentation()
        for label in ("OTHER", "GM_MBAG first", "GM_MBAG last", "TAIL"):
            sl = prs.slides.add_slide(prs.slide_layouts[6])
            for i, token in enumerate(("2D", "3D", "4D", "5D", "6D")):
                box = sl.shapes.add_textbox(
                    Inches(0.5), Inches(0.5 + i * 0.4), Inches(2), Inches(0.3)
                )
                box.text = token
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

    def test_new_section_wrapper_requires_and_verifies_detail_slide(self):
        old_active = fix._active_weekly_before_native_section
        old_com = fix._create_native_section_com_saved
        calls = {"com": 0}
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "weekly.pptx"

            def fake_active(src, out_path, d, g, mode):
                prs = Presentation()
                sl = prs.slides.add_slide(prs.slide_layouts[6])
                for i, token in enumerate(("2D", "3D", "4D", "5D", "6D")):
                    sh = sl.shapes.add_textbox(
                        Inches(0.5),
                        Inches(0.5 + i * 0.4),
                        Inches(2),
                        Inches(0.3),
                    )
                    sh.text = token
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
