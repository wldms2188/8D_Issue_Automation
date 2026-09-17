"""Evidence-weighted issue-origin recommender.

Keeps the existing five user-facing categories while improving recommendation
quality.  It is deliberately conservative: ambiguous cases are sent to
'논의 중' rather than forcing a category.
"""
import re
from dataclasses import dataclass

CATEGORIES=('부품','설계','공정','기타')

# Phrase weights are intentionally strongest for causal/action phrases and weaker
# for generic nouns that can appear in many battery quality issues.
LEXICON={
 '부품':{
  '협력사':3.0,'supplier':3.0,'수입검사':2.5,'입고불량':3.5,'원자재':3.0,'원재료':3.0,
  '소재불량':4.0,'부품불량':4.0,'자재불량':4.0,'lot불량':3.0,'lot편차':3.0,'부품편차':3.0,
  'cell불량':3.5,'셀불량':3.5,'component':2.0,'부품교체':2.5,'업체개선':3.0,
 },
 '설계':{
  '설계변경':4.0,'도면변경':4.0,'공차변경':4.0,'사양변경':3.5,'구조변경':4.0,'치수변경':3.5,
  '설계마진':4.0,'공차':2.5,'간섭':3.0,'강성':2.5,'구조':2.0,'치수':1.5,'도면':1.5,
  'spec변경':3.5,'design':2.0,'설계오류':4.5,'설계미흡':4.0,'설계반영':3.0,
 },
 '공정':{
  '공정조건':3.5,'조건변경':3.0,'작업조건':3.0,'작업표준':3.0,'작업자':2.0,'조립불량':4.0,
  '체결토크':4.0,'토크산포':4.0,'용접조건':4.0,'용접불량':4.0,'도포조건':4.0,'압착조건':4.0,
  '설비조건':3.5,'설비이상':3.5,'가공불량':3.5,'검사누락':3.0,'검출누락':3.0,'공정산포':4.0,
  '조립':1.5,'체결':2.0,'용접':2.0,'도포':2.0,'압착':2.0,'생산공정':2.5,
 },
 '기타':{
  '운송중':3.5,'운송조건':3.5,'보관조건':3.5,'취급부주의':4.0,'고객사용':3.5,'사용조건':3.0,
  '외부충격':4.0,'외부환경':3.5,'환경조건':2.5,'보관':2.0,'운송':2.0,'취급':2.0,
 },
}

FIELD_WEIGHT={
 'cause_4d':1.60, 'leak_cause':1.35, 'system_cause':1.25,
 'action_5d':1.15, 'verification_6d':0.45, 'problem':0.35,
 'occurrence_site':0.35, 'task_name':0.20, 'issue_name':0.20,
}

# A phrase in an explicit negation context should not become positive evidence.
NEGATION=('아님','아니다','무관','문제없','이상없','원인아님','해당없')
UNKNOWN=('tbd','미확인','확인중','검토중','원인미상','분석중','조사중')

@dataclass
class OriginResult:
    recommendation:str
    confidence:str
    score:float
    runner_up:str
    runner_score:float
    evidence:list
    scores:dict
    reason:str


def _n(s): return re.sub(r'\s+','',str(s or '').lower())

def _negated(text,start,end):
    window=text[max(0,start-10):min(len(text),end+10)]
    return any(x in window for x in NEGATION)

def _field_hits(value,category):
    q=_n(value); hits=[]
    if not q:return hits
    for phrase,weight in LEXICON[category].items():
        p=_n(phrase); pos=q.find(p)
        if pos>=0 and not _negated(q,pos,pos+len(p)):
            hits.append((phrase,weight))
    # Avoid double-counting a generic token when a more specific phrase contains it.
    kept=[]
    for phrase,weight in sorted(hits,key=lambda x:(len(_n(x[0])),x[1]),reverse=True):
        if any(_n(phrase) in _n(old[0]) for old in kept):continue
        kept.append((phrase,weight))
    return kept

def classify(d):
    cause=' '.join(str(d.get(k) or '') for k in ('cause_4d','leak_cause','system_cause'))
    cq=_n(cause)
    if not cq or all(x in cq for x in ('미확인',)) or any(cq==x for x in UNKNOWN):
        return OriginResult('TBD','미확정',0.0,'',0.0,[],{c:0.0 for c in CATEGORIES},'4D 원인이 확정되지 않아 자동 분류하지 않습니다.')

    scores={c:0.0 for c in CATEGORIES}; evidence={c:[] for c in CATEGORIES}
    for field,fw in FIELD_WEIGHT.items():
        value=d.get(field)
        for cat in CATEGORIES:
            for phrase,pw in _field_hits(value,cat):
                pts=pw*fw; scores[cat]+=pts
                evidence[cat].append((pts,field,phrase))

    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    best,bscore=ranked[0]; second,sscore=ranked[1]
    margin=bscore-sscore
    # Conservative thresholds: generic single-token evidence should not auto-win.
    if bscore<3.0:
        rec='논의 중'; conf='낮음'
    elif margin<1.5 or (sscore>0 and bscore/sscore<1.35):
        rec='논의 중'; conf='경합'
    elif bscore>=8.0 and margin>=3.0:
        rec=best; conf='높음'
    else:
        rec=best; conf='보통'

    ev=sorted(evidence[best],reverse=True)[:5]
    ev_text=[f'{field}:{phrase}' for _,field,phrase in ev]
    if rec=='논의 중':
        reason=f'상위 후보가 {best} {bscore:.1f}점, {second} {sscore:.1f}점으로 자동 확정 기준을 충족하지 않아 논의 중을 추천합니다.'
    else:
        reason=f'{best} {bscore:.1f}점 / 2순위 {second} {sscore:.1f}점. 주요 근거: '+(', '.join(ev_text) if ev_text else '명시 근거 부족')
    return OriginResult(rec,conf,bscore,second,sscore,ev_text,scores,reason)


def recommend_origin(d):
    r=classify(d)
    label=r.recommendation
    return label, f'신뢰도 {r.confidence} · {r.reason}'
