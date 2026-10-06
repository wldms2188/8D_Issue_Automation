# Recovery STEP5: Issue DB text-format refinements only.
# Preserve STEP4 weekly PPT behavior, extraction, GUI, images, and layout.

import main_recovery_step4 as step4
import main_v310 as v310
import re

base = step4.base
N = v310.N
_EXCEL_ILLEGAL_CONTROL_RE=re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def _excel_safe(text):
    return _EXCEL_ILLEGAL_CONTROL_RE.sub('',N(text))

_old_write_row = base.write_row


def _numbered_lines(text):
    out=[]
    for raw in N(text).split('\n'):
        s=_excel_safe(raw).strip()
        if not s:
            continue
        # remove existing simple bullets/numbering to avoid double numbering
        s=re.sub(r'^[-•·▪◦]\s*','',s)
        s=re.sub(r'^\(?\d+\)?[.)]\s*','',s)
        if s:
            out.append(s)
    return out


def _occurrence_cause_lines(text):
    """Return actual occurrence-cause content without duplicated heading labels."""
    out=[]
    for s in _numbered_lines(text):
        # 8D 4D often contains a structural row/heading such as "Root cause".
        # It is equivalent to the DB's own "발생원인" heading, not cause content.
        if re.fullmatch(r'(?:root\s*cause|발생\s*원인)',s,re.I):
            continue

        # If the heading and content were extracted on one line, keep the
        # content after the heading instead of dropping the whole line.
        stripped=re.sub(
            r'^(?:root\s*cause|발생\s*원인)\s*[:：\-–—]?\s*',
            '',
            s,
            count=1,
            flags=re.I,
        ).strip()
        if stripped:
            out.append(stripped)
    return out


def _cause_db_text(d):
    parts=[]
    cause=_occurrence_cause_lines(d.get('cause_4d'))
    leak=_numbered_lines(d.get('leak_cause'))
    system=_numbered_lines(d.get('system_cause'))

    # DB format requested by user: occurrence cause / escape cause.
    if cause:
        parts.append('1. 발생원인')
        parts.extend(f'{i}) {x}' for i,x in enumerate(cause,1))

    # System cause belongs to the leak-side DB block, not a separate DB heading.
    leak_all=leak+system
    if leak_all:
        parts.append('2. 유출원인')
        parts.extend(f'{i}) {x}' for i,x in enumerate(leak_all,1))

    return _excel_safe('\n'.join(parts))


def _month_label_from_date(d):
    try:
        _,m=base.parse_year_month(d.get('occurrence_date'))
        if m is None or str(m).strip()=='':
            return ''
        return f'{int(str(m).strip())}월'
    except Exception:
        return ''


def write_row_step5(ws,r,d,g):
    # Let the known-good STEP4/legacy writer fill the row first.
    result=_old_write_row(ws,r,d,g)

    # Then change only the two requested Issue DB fields.
    ws.cell(r,7).value=_month_label_from_date(d)   # 발생월: 9월, not 09
    ws.cell(r,16).value=_cause_db_text(d)          # 원인 numbered format
    return result


base.write_row = write_row_step5


class RecoveryStep5App(step4.RecoveryStep4App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY STEP5')


if __name__=='__main__':
    RecoveryStep5App().mainloop()
