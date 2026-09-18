"""Evidence-weighted issue-origin recommender for battery Pack/Module quality."""
import re
from dataclasses import dataclass

CATEGORIES=('부품','설계','공정','기타')

LEXICON={
 '부품':{
  '협력사':3.0,'supplier':3.0,'수입검사':2.5,'입고불량':4.0,'원자재':3.0,'원재료':3.0,'소재불량':4.5,
  '부품불량':4.5,'자재불량':4.5,'lot불량':4.0,'lot편차':3.5,'부품편차':3.5,'cell불량':4.0,'셀불량':4.0,
  '부품교체':2.5,'업체개선':3.0,'커넥터불량':4.0,'하네스불량':4.0,'pcb불량':4.0,'bms불량':4.0,
  '센서불량':4.0,'thermistor불량':4.0,'써미스터불량':4.0,'퓨즈불량':4.0,'릴레이불량':4.0,
  'contactor불량':4.0,'컨택터불량':4.0,'busbar불량':4.0,'버스바불량':4.0,'절연재불량':4.0,
  'seal불량':4.0,'gasket불량':4.0,'가스켓불량':4.0,'소재편차':3.5,'도금불량':3.5,
 },
 '설계':{
  '설계변경':4.5,'도면변경':4.5,'공차변경':4.5,'사양변경':4.0,'구조변경':4.5,'치수변경':4.0,
  '설계마진':4.5,'설계오류':5.0,'설계미흡':4.5,'설계반영':3.5,'공차부족':4.5,'공차':2.5,
  '간섭':3.5,'강성부족':4.5,'강성':2.5,'구조':2.0,'도면':1.5,'spec변경':4.0,'design':2.0,
  '절연거리':4.0,'연면거리':4.0,'방수구조':4.0,'실링구조':4.0,'열설계':4.0,'냉각설계':4.0,
  '유로설계':4.0,'열전달':3.0,'접촉저항':3.0,'진동공진':4.0,'공진':3.0,'내진동':3.0,
  '하네스라우팅':4.0,'배선경로':4.0,'클리어런스':4.0,'체결구조':3.5,'sw로직':4.0,'제어로직':4.0,
  '알고리즘':3.5,'센싱로직':4.0,'진단로직':4.0,'보호로직':4.0,
 },
 '공정':{
  '공정조건':4.0,'조건변경':3.0,'작업조건':3.5,'작업표준':3.5,'작업자':2.5,'조립불량':4.5,
  '체결토크':4.5,'토크산포':4.5,'미체결':4.5,'과체결':4.5,'용접조건':4.5,'용접불량':4.5,
  '미용접':4.5,'과용접':4.5,'스패터':3.5,'도포조건':4.5,'도포불량':4.5,'미도포':4.5,
  '압착조건':4.5,'압착불량':4.5,'crimp불량':4.5,'설비조건':4.0,'설비이상':4.0,'가공불량':4.0,
  '검사누락':3.5,'검출누락':3.5,'공정산포':4.5,'조립':1.5,'체결':2.0,'용접':2.0,'도포':2.0,
  '압착':2.0,'생산공정':2.5,'오조립':4.5,'역조립':4.5,'이물':3.5,'오염':3.0,'세척불량':4.0,
  '실링불량':4.0,'기밀불량':4.0,'누설검사':3.0,'검사조건':3.0,'검사기준':3.0,'치공구':3.5,
  '지그':3.0,'설비편차':4.0,'공정능력':3.5,'관리한계':3.0,
 },
 '기타':{
  '운송중':4.0,'운송조건':4.0,'보관조건':4.0,'취급부주의':4.5,'고객사용':4.0,'사용조건':3.5,
  '외부충격':4.5,'외부환경':4.0,'환경조건':3.0,'보관':2.0,'운송':2.0,'취급':2.0,
  '침수':3.5,'염수':3.5,'결로':3.5,'과충전':3.0,'과방전':3.0,'비정상사용':4.0,'오사용':4.0,
  '차량충돌':4.5,'낙하충격':4.0,
 },
}

FIELD_WEIGHT={
 'cause_4d':1.70,'leak_cause':0.75,'system_cause':0.70,
 'action_5d':1.10,'verification_6d':0.35,'problem':0.30,
 'occurrence_site':0.30,'task_name':0.15,'issue_name':0.20,
}
NEGATION=('아님','아니다','무관','문제없','이상없','원인아님','해당없','영향없','관련없')
UNKNOWN=('tbd','미확인','확인중','검토중','원인미상','분석중','조사중','원인분석중')

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

def _cause_unknown(cause):
    q=_n(cause)
    if not q:return True
    # Do not mark a long causal statement unknown merely because it also says
    # that one sub-cause is still under review.
    return q in UNKNOWN or (len(q)<24 and any(x in q for x in UNKNOWN))

def _negated(text,start,end):
    # Negation is checked locally, not across the whole sentence.
    before=text[max(0,start-8):start]
    after=text[end:min(len(text),end+8)]
    return any(x in before or x in after for x in NEGATION)

def _field_hits(value,category):
    q=_n(value); hits=[]
    if not q:return hits
    for phrase,weight in LEXICON[category].items():
        p=_n(phrase)
        for m in re.finditer(re.escape(p),q):
            if not _negated(q,m.start(),m.end()):
                hits.append((phrase,weight)); break
    kept=[]
    for phrase,weight in sorted(hits,key=lambda x:(len(_n(x[0])),x[1]),reverse=True):
        if any(_n(phrase) in _n(old[0]) for old in kept):continue
        kept.append((phrase,weight))
    return kept

def classify(d):
    cause=' '.join(str(d.get(k) or '') for k in ('cause_4d','leak_cause','system_cause'))
    if _cause_unknown(cause):
        return OriginResult('TBD','미확정',0.0,'',0.0,[],{c:0.0 for c in CATEGORIES},'4D 원인이 확정되지 않아 자동 분류하지 않습니다.')

    scores={c:0.0 for c in CATEGORIES}; evidence={c:[] for c in CATEGORIES}
    for field,fw in FIELD_WEIGHT.items():
        value=d.get(field)
        for cat in CATEGORIES:
            for phrase,pw in _field_hits(value,cat):
                pts=pw*fw; scores[cat]+=pts; evidence[cat].append((pts,field,phrase))

    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    best,bscore=ranked[0]; second,sscore=ranked[1]; margin=bscore-sscore
    # When the confirmed cause itself contains strong evidence for two technical
    # origins, keep it for human discussion even if the action plan adds more
    # weight to one side.  Corrective actions must not erase causal ambiguity.
    cause_scores={c:0.0 for c in CATEGORIES}
    for field in ('cause_4d','leak_cause','system_cause'):
        fw=FIELD_WEIGHT[field]
        for cat in CATEGORIES:
            for _,pw in _field_hits(d.get(field),cat):
                cause_scores[cat]+=pw*fw
    cause_ranked=sorted(cause_scores.items(),key=lambda x:x[1],reverse=True)
    cause_ambiguous=(cause_ranked[1][1]>=3.2 and\n                     cause_ranked[0][1]/cause_ranked[1][1]<1.80)
    if cause_ambiguous:
        rec='논의 중'; conf='경합'
    elif bscore<3.2:
        rec='논의 중'; conf='낮음'
    elif margin<1.8 or (sscore>0 and bscore/sscore<1.40):
        rec='논의 중'; conf='경합'
    elif bscore>=8.5 and margin>=3.2:
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
    return r.recommendation, f'신뢰도 {r.confidence} · {r.reason}'
