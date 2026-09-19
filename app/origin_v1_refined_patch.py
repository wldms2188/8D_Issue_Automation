"""Stable V1-style issue-origin recommender with refined rule analysis.
No numeric score is exposed. It stays human-in-the-loop and returns discussion-needed on mixed evidence.
"""
import re
import main_recovery_step8 as step8

RULES = {
    '부품': (
        '부품','자재','소재','원재료','협력사','입고','supplier','component',
        'lot편차','부품불량','셀불량','재료불량','수입검사','부품편차',
        'part','parts','material defect','component defect','supplier issue','incoming defect','part variation',
    ),
    '설계': (
        '설계','design','도면','공차','사양','spec','구조','치수','강성','간섭',
        '설계마진','공차설계','구조간섭','사양미흡','설계변경','도면변경',
        'design margin','design issue','drawing','tolerance','dimension','structural interference','specification',
    ),
    '공정': (
        '공정','작업','조립','체결','토크','용접','설비','가공','도포','압착',
        '공정조건','작업조건','작업표준','공정산포','토크산포','용접조건','도포조건',
        'process','process condition','work condition','assembly','fastening','torque','welding','equipment','coating','pressing',
    ),
    '기타': (
        '운송','보관','취급','고객사용','외부충격','환경조건','사용조건','취급부주의',
        'transport','storage','handling','customer use','external impact','environmental condition','usage condition',
    ),
}
NEGATIONS=('문제없음','이상없음','정상','원인아님','무관','영향없음','no issue','no problem','no abnormality','normal','not cause','not a cause','not related','no impact')
MIXERS=('동시에','복합','및','그리고')


def _norm(s):
    return re.sub(r'\s+','',str(s or '').lower())


def _negated(q, term):
    """True only when every occurrence of term is locally negated."""
    t=_norm(term)
    if not t:
        return False
    positions=[m.start() for m in re.finditer(re.escape(t),q)]
    if not positions:
        return False
    for pos in positions:
        # Negation normally follows the subject phrase (e.g. 공정조건 문제없음).
        after=q[pos+len(t):pos+len(t)+48]
        if not any(n in after for n in NEGATIONS):
            return False
    return True


def recommend_origin(d):
    text=step8._cause_text(d)
    if not text:
        return 'TBD','4D 원인 내용이 아직 없어 TBD로 표시합니다.'
    q=_norm(text)
    evidence={}
    for cat,terms in RULES.items():
        hits=[]
        for term in terms:
            t=_norm(term)
            if t and t in q and not _negated(q,t):
                # Avoid counting a short token again when a more specific phrase already explains it.
                if not any(t in _norm(h) or _norm(h) in t for h in hits):
                    hits.append(term)
        if hits:
            evidence[cat]=hits

    if not evidence:
        return '논의 중','원인 내용은 있으나 부품/설계/공정/기타로 명확히 분류할 근거가 부족합니다.'
    if len(evidence)==1:
        cat=next(iter(evidence))
        return cat, f"8D 원인 내용에서 {', '.join(evidence[cat][:3])} 관련 표현이 확인되어 {cat}을(를) 추천합니다."

    # Multiple categories: only auto-pick when one category has clearly richer independent evidence.
    ranked=sorted(evidence.items(),key=lambda kv:len(kv[1]),reverse=True)
    top_cat,top_hits=ranked[0]
    second_hits=ranked[1][1]
    mixed=any(m in q for m in MIXERS)
    if mixed or len(top_hits) <= len(second_hits)+1:
        cats=', '.join(evidence)
        return '논의 중',f'원인 내용에 {cats} 범주의 단서가 함께 있어 추가 논의가 필요합니다.'
    return top_cat, f"8D 원인 내용에서 {', '.join(top_hits[:3])} 관련 표현이 상대적으로 명확하여 {top_cat}을(를) 추천합니다."


step8._recommend_origin = recommend_origin
