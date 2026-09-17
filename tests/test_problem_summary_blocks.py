"""Regression tests for Issue DB summary block boundaries.
Run from repository root:
    py -m unittest tests.test_problem_summary_blocks -v
"""
import re
import unittest

MAX_LEN = 150


def structured_blocks(src):
    marker = re.compile(r"^\s*(?:[-•·]|\(?\d+\s*[.)])\s*")
    lines = [x.strip() for x in str(src).replace("\r\n", "\n").replace("\r", "\n").split("\n") if x.strip()]
    blocks, current = [], []
    for line in lines:
        if marker.match(line) and current:
            blocks.append(" ".join(current).strip())
            current = [line]
        else:
            current.append(line)
    if current:
        blocks.append(" ".join(current).strip())
    return blocks


def pack_blocks(src, header=""):
    blocks = structured_blocks(src)
    selected = []
    prefix = (header + "\n") if header else ""
    used = len(prefix)
    for block in blocks:
        sep = 3 if selected else 0
        if used + sep + len(block) > MAX_LEN:
            break
        selected.append(block)
        used += sep + len(block)
    return (prefix + " / ".join(selected)).strip()


class SummaryBlockBoundaryTests(unittest.TestCase):
    def assert_no_partial(self, marker):
        first = f"{marker} 현상 " + "가" * 35
        second = "- 시험조건 " + "나" * 110
        third = "3. 뒤의 짧은 항목"
        out = pack_blocks("\n".join([first, second, third]), "• 시험명 : 진동 시험")
        self.assertIn("현상", out)
        self.assertNotIn("시험조건", out)
        self.assertNotIn("뒤의 짧은 항목", out)
        self.assertLessEqual(len(out), MAX_LEN)

    def test_dash_boundary(self): self.assert_no_partial("-")
    def test_bullet_boundary(self): self.assert_no_partial("•")
    def test_middle_dot_boundary(self): self.assert_no_partial("·")
    def test_number_dot_boundary(self): self.assert_no_partial("1.")
    def test_number_paren_boundary(self): self.assert_no_partial("1)")
    def test_parenthesized_number_boundary(self): self.assert_no_partial("(1)")

    def test_complete_next_block_is_kept(self):
        src = "1. 현상 통신 불가\n- DUT3 재현\n2. 시험조건 DV 진동"
        out = pack_blocks(src)
        self.assertIn("2. 시험조건 DV 진동", out)
        self.assertLessEqual(len(out), MAX_LEN)

    def test_overflow_stops_in_source_order(self):
        src = "1. 현상 " + "가" * 40 + "\n2. 시험조건 " + "나" * 120 + "\n3. 짧음"
        out = pack_blocks(src)
        self.assertNotIn("2. 시험조건", out)
        self.assertNotIn("3. 짧음", out)

    def test_continuation_lines_belong_to_same_block(self):
        src = "1. 현상\n첫째 설명\n둘째 설명\n2. 시험조건\n조건 설명"
        blocks = structured_blocks(src)
        self.assertEqual(blocks[0], "1. 현상 첫째 설명 둘째 설명")
        self.assertEqual(blocks[1], "2. 시험조건 조건 설명")


if __name__ == "__main__":
    unittest.main()
