# 8D Issue Automation v3.2.3
# Shortened weekly problem keeps the actual symptom + test/build name,
# while removing test conditions/details.
import re

import main_v322 as v322
import main_v321 as v321

base = v322.base
N = v321.N
C = v321.C


# Lines/tokens that are normally test/build conditions rather than the issue symptom.
_CONDITION_WORDS = (
    '시험조건','시험 조건','상세시험조건','상세 시험 조건','조건','환경조건','환경 조건',
    'soc','dod','soh','온도','습도','전류','전압조건','전압 조건','충전조건','충전 조건',
    '방전조건','방전 조건','충방전조건','충방전 조건','시간','cycle','사이클','샘플수','sample 수',
    '샘플 수','시험방법','시험 방법','시험절차','시험 절차','프로파일','profile','c-rate','crate'
)

# Words that strongly indicate an actual issue/symptom. Preserve these lines even if
# they also contain numeric values or some condition wording.
_SYMPTOM_WORDS = (
    '발생','불량','이상','파손','파단','누액','누수','누설','변형','부풀','스웰','swelling',
    '전압저하','전압 저하','voltage drop','drop','단선','단락','쇼트','short','통신불가','통신 불가',
    '미동작','오동작','미점등','소손','탄화','크랙','crack','이탈','탈락','변색','과열','발열',
    '편차','오차','미검출','검출불가','검출 불가','정지','fail','failure','ng'
)


def _clean_problem_line(line):
    s=N(line)
    s=re.sub(r'^[-•·▪◦]\s*','',s)
    s=re.sub(r'^\(?\d+\)?[.)]\s*','',s)
    return s.strip()


def _looks_like_symptom(line):
    q=C(line)
    return bool(q) and any(C(w) in q for w in _SYMPTOM_WORDS)


def _looks_like_condition(line):
    q=C(line)
    if not q:
        return False
    if any(C(w) in q for w in _CONDITION_WORDS):
        return True
    # Numeric/unit-heavy standalone lines are usually detailed conditions.
    units=('℃','°c','%','a','ma','v','mv','w','kw','min','sec','hr','hour','h','c-rate','c')
    digit_count=sum(ch.isdigit() for ch in line)
    if digit_count>=2 and any(u in line.lower() for u in units):
        return True
    return False


def _symptom_lines(problem, event_name=''):
    kept=[]
    for raw in N(problem).split('\n'):
        s=_clean_problem_line(raw)
        if not s:
            continue
        q=C(s)
        # Event-name-only lines are represented separately as 시험명/빌드명.
        if event_name and C(event_name)==q:
            continue
        # Actual symptom has priority over condition filtering.
        if _looks_like_symptom(s):
            kept.append(s)
            continue
        if _looks_like_condition(s):
            continue
        # Preserve short descriptive lines that are not obviously metadata/conditions.
        if not any(C(x) in q for x in ('시험일자','시험 일자','시험장소','시험 장소','장비','설비','lot','로트','시료','sample')):
            kept.append(s)

    # De-duplicate while preserving source order.
    out=[]; seen=set()
    for s in kept:
        k=C(s)
        if k and k not in seen:
            seen.add(k); out.append(s)
    return out


def _short_problem(d):
    kind,name=v321._event_name(d)
    lines=[]
    if kind=='시험' and name:
        lines.append(f'시험명: {name}')
    elif kind=='빌드' and name:
        lines.append(f'빌드명: {name}')

    symptoms=_symptom_lines(d.get('problem'),name)
    lines.extend(symptoms)

    # If filtering could not isolate a symptom, do not erase the actual problem.
    # Keep the first non-condition source line as a safe fallback.
    if len(lines)<=1:
        for raw in N(d.get('problem')).split('\n'):
            s=_clean_problem_line(raw)
            if s and not _looks_like_condition(s):
                if not name or C(s)!=C(name):
                    lines.append(s)
                    break

    return '\n'.join(lines).strip() or N(d.get('problem'))


# Replace only the shortening rule. All v3.2.2 status confirmation and v3.2.1
# GUI/Excel/PPT behavior remain unchanged.
v321._short_problem=_short_problem


class App(v322.App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.3')


if __name__=='__main__':
    App().mainloop()
