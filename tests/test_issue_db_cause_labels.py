"""Regression tests for Issue DB occurrence-cause label cleanup."""
import unittest

import main_recovery_step5 as step5


class IssueDbCauseLabelTests(unittest.TestCase):
    def test_root_cause_heading_line_is_not_numbered_as_content(self):
        out = step5._cause_db_text({
            "cause_4d": "Root cause\nBolt loosening\nTorque insufficient",
            "leak_cause": "",
            "system_cause": "",
        })
        self.assertIn("1. 발생원인", out)
        self.assertNotIn("Root cause", out)
        self.assertIn("1) Bolt loosening", out)
        self.assertIn("2) Torque insufficient", out)

    def test_root_cause_prefix_keeps_real_content(self):
        out = step5._cause_db_text({
            "cause_4d": "Root cause : Bolt loosening",
            "leak_cause": "",
            "system_cause": "",
        })
        self.assertNotIn("Root cause", out)
        self.assertIn("1) Bolt loosening", out)

    def test_korean_occurrence_cause_heading_is_also_not_duplicated(self):
        out = step5._cause_db_text({
            "cause_4d": "발생원인\n체결 토크 부족",
            "leak_cause": "",
            "system_cause": "",
        })
        self.assertEqual(out, "1. 발생원인\n1) 체결 토크 부족")


if __name__ == "__main__":
    unittest.main()
