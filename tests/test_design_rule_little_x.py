"""Regression tests for automatic Issue DB design-rule Little x classification."""
import unittest

from openpyxl import Workbook

import design_rule_little_x as dr


class DesignRuleLittleXTests(unittest.TestCase):
    def test_all_36_rule_titles_map_to_their_own_code(self):
        self.assertEqual(len(dr.RULES), 36)
        for rule in dr.RULES:
            with self.subTest(code=rule["code"], name=rule["name"]):
                self.assertIn(rule["code"], dr.classify_little_x(rule["name"]))

    def test_representative_pack_cause_examples(self):
        cases = {
            "가스켓 간섭에 의한 캔 크랙 발생": "1.1.1",
            "레이저 오조사로 실링부 손상": "1.1.3",
            "전극 돌출 확인": "1.1.2",
            "볼트 풀림으로 버스바 체결력 불량": "3.3.1",
            "접촉면 Air Gap 발생으로 접촉 불량": "3.3.2",
            "접지 누락으로 접지 연속성 불만족": "5.3.1",
            "LQC 공정에서 오조립 불량 다발": "6.1.2",
            "EOL 전압 측정 불량 다발": "6.1.3",
        }
        for cause, code in cases.items():
            with self.subTest(cause=cause):
                self.assertIn(code, dr.classify_little_x(cause))

    def test_multiple_matching_rules_are_preserved(self):
        codes = dr.classify_little_x(
            "볼트 풀림으로 체결력 불량 발생, 접촉면 Air Gap으로 접촉 불량 동시 확인"
        )
        self.assertIn("3.3.1", codes)
        self.assertIn("3.3.2", codes)

    def test_unrelated_cause_returns_no_rule(self):
        self.assertEqual(
            dr.classify_little_x("원인 분석 진행 중이며 현재 명확한 원인 미확인"),
            [],
        )

    def test_ambiguous_match_is_recommendation_not_final_multi_write(self):
        rec = dr.recommendation(
            "볼트 풀림으로 체결력 불량 발생, 접촉면 Air Gap으로 접촉 불량 동시 확인"
        )
        self.assertTrue(rec["ambiguous"])
        self.assertIn("3.3.1", rec["candidates"])
        self.assertIn("3.3.2", rec["candidates"])

        wb = Workbook()
        ws = wb.active
        ws.cell(6, 19).value = "규칙 위반 여부"
        ws.cell(6, 20).value = "Little x"
        result = dr.write_confirmed_rule_classification(
            ws, 8, "3.3.2", confirmed=True
        )
        self.assertEqual(result["violation"], "Y")
        self.assertEqual(result["codes"], ["3.3.2"])
        self.assertEqual(ws.cell(8, 19).value, "Y")
        self.assertEqual(ws.cell(8, 20).value, "3.3.2")
        self.assertNotIn("3.3.1", ws.cell(8, 20).value)

    def test_single_match_is_still_only_a_recommendation_until_user_confirms(self):
        rec = dr.recommendation("볼트 풀림으로 체결력 불량 발생")
        self.assertFalse(rec["ambiguous"])
        self.assertEqual(rec["recommended"], "3.3.1")

        wb = Workbook()
        ws = wb.active
        ws.cell(6, 19).value = "규칙 위반 여부"
        ws.cell(6, 20).value = "Little x"
        ws.cell(8, 19).value = "기존값"
        ws.cell(8, 20).value = "기존규칙"
        result = dr.write_confirmed_rule_classification(
            ws, 8, rec["recommended"], confirmed=False
        )
        self.assertIsNone(result["violation"])
        self.assertEqual(ws.cell(8, 19).value, "기존값")
        self.assertEqual(ws.cell(8, 20).value, "기존규칙")

    def test_user_can_confirm_no_violation(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(6, 19).value = "설계 규칙 위반 여부"
        ws.cell(6, 20).value = "Little x"
        result = dr.write_confirmed_rule_classification(
            ws, 8, dr.NONE_CODE, confirmed=True
        )
        self.assertEqual(result["violation"], "N")
        self.assertEqual(result["codes"], [])
        self.assertEqual(ws.cell(8, 19).value, "N")
        self.assertEqual(ws.cell(8, 20).value, "")

    def test_header_detection_and_y_code_write(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(6, 19).value = "설계 규칙 위반 여부"
        ws.cell(6, 20).value = "Little x"
        result = dr.write_rule_classification(
            ws, 8, "볼트 풀림으로 체결력 불량 발생"
        )
        self.assertEqual((result["violation_col"], result["little_x_col"]), (19, 20))
        self.assertEqual(ws.cell(8, 19).value, "Y")
        self.assertIn("3.3.1", ws.cell(8, 20).value)

    def test_no_match_writes_n_and_blank_little_x(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(5, 22).value = "규칙 위반 여부 Y/N"
        ws.cell(5, 23).value = "위반 규칙 (Little X)"
        dr.write_rule_classification(ws, 8, "원인 미확인")
        self.assertEqual(ws.cell(8, 22).value, "N")
        self.assertEqual(ws.cell(8, 23).value, "")


if __name__ == "__main__":
    unittest.main()
