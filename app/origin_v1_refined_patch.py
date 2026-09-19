"""Stable V1-style issue-origin recommender with refined rule analysis.
No numeric score is exposed. It stays human-in-the-loop and returns discussion-needed on mixed evidence.
"""
import re
import main_recovery_step8 as step8

RULES = {
    '부품': (
        # Korean
        '부품','자재','소재','원재료','협력사','입고','lot편차','부품불량','셀불량',
        '재료불량','수입검사','부품편차',
        # English: supplier / part / material / cell origin
        'supplier','supplier part','supplier component','supplier defect','supplier variation',
        'part','parts','part defect','defective part','part variation','component','components',
        'component defect','component variation','material','raw material','material defect',
        'material variation','material property','cell','cell defect','cell variation',
        'incoming part','incoming material','lot variation','batch variation','vendor defect',
    ),
    '설계': (
        # Korean
        '설계','도면','공차','사양','규격','스펙','구조','치수','강성','간섭','설계마진','공차설계',
        '구조간섭','사양미흡','규격 미흡','규격미흡','스펙 미흡','스펙미흡','설계변경','도면변경',
        '설계 요구사항','설계요구사항','치수 부적합','치수부적합','공차 부적정','공차부적정',
        # English: design / drawing / dimensional / specification origin
        'design','design issue','design defect','design margin','insufficient design margin',
        'design review','design change','drawing','drawing error','drawing change',
        'tolerance','tolerance stack','tolerance stack up','tolerance stack-up',
        'clearance','insufficient clearance','interference','structural interference',
        'geometry','geometric','spec','specification','spec limit','specification limit',
        'design spec','design specification','requirement','design requirement',
        'specification issue','structural','stiffness','strength margin','design selection',
        'design requirement','design criteria','cad','layout interference',
    ),
    '공정': (
        # Korean
        '공정','작업','조립','체결','토크','용접','설비','가공','도포','압착','공정조건',
        '작업조건','작업표준','공정산포','토크산포','용접조건','도포조건',
        '볼트 풀림','볼트풀림','체결 풀림','체결풀림','체결불량','체결 불량','체결 미흡','체결미흡',
        '체결력 부족','체결력부족','토크 부족','토크부족','볼트 체결','볼트체결','너트 풀림','너트풀림',
        # English: manufacturing / assembly / equipment / parameter origin
        'process','manufacturing process','process condition','process parameter',
        'process variation','manufacturing variation','work condition','work instruction',
        'assembly','assembly condition','assembly error','fastening','fastening condition','fastening failure',
        'fastener loosening','loosened fastener','bolt loosening','loose bolt','fixing bolt',
        'bolt fastening','bolt torque','insufficient torque','under torque','under-torque',
        'torque','torque variation',
        'welding','weld','welding condition','coating','coating condition','dispensing',
        'adhesive application','pressing','press fit','crimp','crimping','riveting','curing',
        'equipment','machine','machine setting','equipment setting','fixture','jig',
        'tooling','operator','operator error','setup','alignment','calibration',
        'production condition','manufacturing condition',
    ),
    '기타': (
        # Korean
        '운송','보관','취급','고객사용','외부충격','환경조건','사용조건','취급부주의',
        # English: logistics / environment / customer-use origin
        'transport','transportation','shipping','shipment damage','storage','warehousing',
        'handling','mishandling','customer use','customer usage','customer handling',
        'misuse','external impact','external force','environmental condition',
        'environment','usage condition','use condition','field condition',
        'road condition','packaging damage','logistics damage',
    ),
}

NEGATIONS=(
    # Korean
    '문제없음','이상없음','정상','원인아님','무관','영향없음',
    # English after/before-term negation and exclusion
    'no issue','no problem','no abnormality','normal','not cause','not a cause',
    'not caused by','not due to','not related','not related to','no impact',
    'no effect','without issue','within spec','within specification','ruled out',
    'excluded as cause','not contributing','not contributory','not responsible',
)
MIXERS=('동시에','복합','및','그리고','combined','combination','both','together','and')

def _norm(s):
    return re.sub(r'\s+','',str(s or '').lower())

def _plain(s):
    return re.sub(r'\s+',' ',str(s or '').lower()).strip()

def _term_positions(text,term):
    q=_plain(text); t=_plain(term)
    if not t:
        return []
    # English terms use word-ish boundaries to avoid short-token false matches.
    if re.search(r'[a-z]',t):
        pat=r'(?<![a-z0-9])'+re.escape(t)+r'(?![a-z0-9])'
        return [(m.start(),m.end()) for m in re.finditer(pat,q,re.I)]
    return [(m.start(),m.end()) for m in re.finditer(re.escape(t),q,re.I)]

def _negated(text,term):
    """True only when every occurrence is negated inside its own sentence/clause."""
    q=_plain(text)
    positions=_term_positions(text,term)
    if not positions:
        return False
    negs=tuple(_plain(n) for n in NEGATIONS)
    separators='.;\n'
    for a,b in positions:
        left=max([q.rfind(sep,0,a) for sep in separators]+[-1])+1
        rights=[q.find(sep,b) for sep in separators]
        rights=[x for x in rights if x>=0]
        right=min(rights) if rights else len(q)
        clause=q[left:right].strip()
        if not any(n in clause for n in negs):
            return False
    return True

def _hits_for_text(text):
    q=_norm(text)
    evidence={}
    for cat,terms in RULES.items():
        hits=[]
        for term in terms:
            t=_norm(term)
            if not t:
                continue
            # English terms require a real word/phrase match. This prevents short
            # tokens such as "spec" from matching unrelated words like "specific".
            if re.search(r'[A-Za-z]',str(term)):
                if not _term_positions(text,term):
                    continue
            elif t not in q:
                continue
            if _negated(text,term):
                continue
            # Avoid double-counting generic tokens contained in a specific phrase.
            normalized=[_norm(h) for h in hits]
            if any(t in h or h in t for h in normalized):
                # Prefer the more specific/longer phrase as evidence.
                replace=None
                for i,h in enumerate(normalized):
                    if h in t and len(t)>len(h):
                        replace=i; break
                if replace is not None:
                    hits[replace]=term
                continue
            hits.append(term)
        if hits:
            evidence[cat]=hits
    return evidence

def _occurrence_site_prior(d):
    """Return a weak prior from occurrence site; screen selection wins over PPT value.

    _origin_occurrence_site is the value selected in the GUI.  Only when that
    value is missing do we fall back to the occurrence site extracted from PPT.
    """
    selected=_plain((d or {}).get('_origin_occurrence_site'))
    ppt_site=_plain((d or {}).get('occurrence_site'))
    site=selected if selected else ppt_site
    compact=_norm(site)
    if not site:
        return None,None

    part_sites=(
        '부품 생산','부품생산','component production','part production',
        'component manufacturing','part manufacturing','supplier production'
    )
    product_sites=(
        '제품 생산','제품생산','product production','product manufacturing',
        'pack production','pack manufacturing','assembly production'
    )
    if any(_norm(x) in compact for x in part_sites):
        return '부품',f'"{site}" 중 발생한 이슈이며, 추가 원인 분류 단서가 부족해 부품 기인 가능성을 보조 근거로 반영했습니다.'
    if any(_norm(x) in compact for x in product_sites):
        return '공정',f'"{site}" 중 발생한 이슈이며, 추가 원인 분류 단서가 부족해 공정 기인 가능성을 보조 근거로 반영했습니다.'
    return None,None

def _decide_with_site_prior(evidence,text,d):
    """Use occurrence-site only as a tie-break/support signal for ambiguous evidence."""
    if not evidence:
        prior,reason=_occurrence_site_prior(d)
        return (prior,reason) if prior else (None,None)

    if len(evidence)==1:
        return _decide_from_evidence(evidence,text)

    rec,reason=_decide_from_evidence(evidence,text)
    if rec!='논의 중':
        return rec,reason

    prior,site_reason=_occurrence_site_prior(d)
    if prior and prior in evidence:
        hits=evidence.get(prior) or []
        return prior,(
            f'4D 원인 내용에 여러 범주의 단서가 있으나 '
            f'{", ".join(hits[:3]) or prior} 관련 표현이 확인되고, '
            f'{site_reason}'
        )
    return rec,reason

def _decide_from_evidence(evidence,text):
    if not evidence:
        return None,None
    if len(evidence)==1:
        cat=next(iter(evidence))
        return cat, f"8D 원인 내용에서 {', '.join(evidence[cat][:3])} 관련 표현이 확인되어 {cat}을(를) 추천합니다."

    ranked=sorted(evidence.items(),key=lambda kv:len(kv[1]),reverse=True)
    top_cat,top_hits=ranked[0]
    second_hits=ranked[1][1]
    q=_plain(text)
    mixed=any(re.search(r'(?<![a-z])'+re.escape(m)+r'(?![a-z])',q) if re.search(r'[a-z]',m) else m in _norm(text) for m in MIXERS)

    # Mixed cause statements should stay human-in-the-loop unless one category
    # has clearly richer independent evidence.
    if mixed or len(top_hits)<=len(second_hits)+1:
        cats=', '.join(evidence)
        return '논의 중',f'원인 내용에 {cats} 범주의 단서가 함께 있어 추가 논의가 필요합니다.'
    return top_cat, f"8D 원인 내용에서 {', '.join(top_hits[:3])} 관련 표현이 상대적으로 명확하여 {top_cat}을(를) 추천합니다."

def recommend_origin(d):
    """Bilingual issue-origin recommendation.

    Occurrence/root cause is primary. Escape/system cause can support a decision
    only when occurrence cause has no usable category evidence, so phrases such as
    "inspection missed" or "control plan gap" cannot override the real root cause.
    """
    occurrence=str(d.get('cause_4d') or '').strip()
    leak=str(d.get('leak_cause') or '').strip()
    system=str(d.get('system_cause') or '').strip()
    if not (occurrence or leak or system):
        prior,site_reason=_occurrence_site_prior(d)
        if prior:
            return prior,site_reason
        return 'TBD','4D 원인 내용이 아직 없어 TBD로 표시합니다.'

    primary=_hits_for_text(occurrence)
    if primary:
        rec,reason=_decide_with_site_prior(primary,occurrence,d)
        if rec:
            return rec,reason

    # Fallback only: when occurrence cause has no categorizable evidence.
    supporting='\n'.join(x for x in (leak,system) if x)
    support=_hits_for_text(supporting)
    if support:
        rec,reason=_decide_with_site_prior(support,supporting,d)
        if rec:
            return rec,reason

    # Occurrence site is a weak final fallback, not a replacement for explicit cause text.
    prior,site_reason=_occurrence_site_prior(d)
    if prior:
        return prior,site_reason

    return '논의 중','원인 내용은 있으나 부품/설계/공정/기타로 명확히 분류할 근거가 부족합니다.'


step8._recommend_origin = recommend_origin
