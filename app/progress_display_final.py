"""Final progress display polish.

The v3 header already has a dedicated percent label. Therefore the status/progress
text must not repeat a leading numeric percentage (e.g. '85% · 85...').
"""
import re
import main_enterprise_v3 as v3

_original_set_progress = v3.EnterpriseAppV3._set_progress


def clean_set_progress(self, pct, label):
    text = str(label or '').strip()
    # Remove accidental repeated leading percentage(s) from inherited status text.
    text = re.sub(r'^(?:\s*\d{1,3}\s*%?\s*[·|:\-]?\s*)+', '', text).strip()
    if not text:
        text = '처리 중' if int(pct) < 100 else '업데이트 완료'
    return _original_set_progress(self, pct, text)


v3.EnterpriseAppV3._set_progress = clean_set_progress
