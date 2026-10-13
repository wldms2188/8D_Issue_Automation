"""Automatic Issue DB design-rule violation / Y.y.x classification.

The rule catalog below is transcribed from the Pack/Module Design Rule sheet
provided by the user (Big Y 1~6, Y.y.x 36 items).

Matching is intentionally conservative:
- explicit Y.y.x title/strong phrase wins;
- example-style matches require all terms in a group;
- only cause text is evaluated, not the symptom/action text.
"""
import re


def _text(value):
    return "" if value is None else str(value).strip()


def _key(value):
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", _text(value)).casefold()


def _rule(code, name, phrases=(), groups=()):
    return {
        "code": code,
        "name": name,
        "phrases": tuple(phrases),
        "groups": tuple(tuple(x) for x in groups),
    }


ENGLISH_RULE_ALIASES = {
    "1.1.1": {
        "phrases": (
            "excessive interference or external force",
            "gasket interference crack",
            "mechanical interference crack",
            "external force induced crack",
        ),
        "groups": (
            ("interference", "crack"),
            ("external force", "crack"),
            ("mechanical test", "crack"),
            ("gasket", "interference", "crack"),
        ),
    },
    "1.1.2": {
        "phrases": (
            "excessive internal electrode movement",
            "electrode movement",
            "electrode protrusion",
            "electrode displacement",
        ),
        "groups": (
            ("electrode", "protrusion"),
            ("electrode", "displacement"),
        ),
    },
    "1.1.3": {
        "phrases": (
            "laser welding damage",
            "laser misalignment",
            "laser mis-shot",
            "laser over penetration",
            "excessive laser penetration",
        ),
        "groups": (
            ("laser", "penetration"),
            ("laser", "seal", "damage"),
            ("laser", "overheat"),
            ("laser", "burn through"),
        ),
    },
    "1.2.1": {
        "phrases": (
            "anode connection",
            "lead mis-insertion",
            "lead displacement",
            "cell to cell contact",
            "cell contact failure",
        ),
        "groups": (
            ("anode", "connection"),
            ("lead", "misinsert"),
        ),
    },
    "1.2.2": {
        "phrases": (
            "weld joint corrosion",
            "joint corrosion",
            "busbar corrosion",
            "lead corrosion",
            "wire corrosion",
            "immersion coolant corrosion",
            "insulating oil corrosion",
            "potting resin corrosion",
        ),
        "groups": (
            ("weld", "corrosion"),
            ("joint", "corrosion"),
            ("busbar", "corrosion"),
        ),
    },
    "1.3.1": {
        "phrases": (
            "excessive gas generation due to external heat",
            "external heat induced gas generation",
            "external heat ingress",
        ),
        "groups": (
            ("external heat", "gas"),
            ("terminal", "heating", "gas"),
            ("busbar", "heating", "gas"),
        ),
    },
    "1.3.2": {
        "phrases": (
            "excessive gas generation due to internal heat",
            "internal heat induced gas generation",
            "internal overheating",
        ),
        "groups": (
            ("internal heat", "gas"),
            ("short circuit", "overheat"),
        ),
    },
    "1.3.3": {
        "phrases": (
            "insufficient cooling performance",
            "gas generation due to insufficient cooling",
            "cooling performance shortage",
            "poor cooling performance",
        ),
        "groups": (
            ("cooling", "insufficient"),
            ("cooling", "overheat"),
            ("lifetime test", "overheat"),
        ),
    },

    "2.1.1": {
        "phrases": (
            "excessive tab tensile force",
            "tab tensile load",
            "tab tension excessive",
            "tab crack due to tensile force",
        ),
        "groups": (
            ("tab", "tensile"),
            ("tab", "tension"),
        ),
    },
    "2.2.1": {
        "phrases": (
            "improper compression force",
            "inadequate compression force",
            "insufficient cell compression",
            "improper spacing",
            "lithium plating at electrode interface",
        ),
        "groups": (
            ("compression", "improper"),
            ("compression", "insufficient"),
            ("lithium plating", "interface"),
        ),
    },
    "2.2.2": {
        "phrases": (
            "improper voltage sensing",
            "improper temperature sensing",
            "voltage sensing error",
            "temperature sensing error",
            "lithium plating at low temperature",
        ),
        "groups": (
            ("voltage", "sensing", "improper"),
            ("temperature", "sensing", "improper"),
            ("low temperature", "lithium plating"),
        ),
    },

    "3.1.1": {
        "phrases": (
            "coolant leakage",
            "coolant leak",
            "coolant gasket sealing failure",
            "coolant weld sealing failure",
            "o-ring leak",
            "hose leak",
            "clamp leak",
            "fitting leak",
            "cooling plate brazing leak",
        ),
        "groups": (
            ("coolant", "gasket", "leak"),
            ("coolant", "weld", "leak"),
            ("coolant", "seal", "failure"),
        ),
    },
    "3.1.2": {
        "phrases": (
            "enclosure gasket sealing failure",
            "housing sealing failure",
            "frame weld leakage",
            "pack cover gasket leak",
            "connector gasket leak",
        ),
        "groups": (
            ("enclosure", "gasket", "leak"),
            ("housing", "seal", "failure"),
            ("frame", "weld", "leak"),
        ),
    },
    "3.2.1": {
        "phrases": (
            "insufficient component clearance",
            "component interference",
            "insufficient structural stiffness",
            "component stiffness failure",
            "harness damage by interference",
        ),
        "groups": (
            ("component", "clearance", "insufficient"),
            ("component", "interference", "damage"),
            ("harness", "interference", "damage"),
            ("bracket", "position", "interference"),
        ),
    },
    "3.2.2": {
        "phrases": (
            "component overheating",
            "insufficient heat resistance",
            "over current rating",
            "current rating exceeded",
            "insulation covering deterioration",
            "plastic thermal degradation",
        ),
        "groups": (
            ("component", "overheat"),
            ("heat resistance", "insufficient"),
            ("current rating", "exceed"),
        ),
    },
    "3.3.1": {
        "phrases": (
            "insufficient fastening force",
            "insufficient retention force",
            "bolt loosening",
            "loose bolt",
            "terminal rotation",
            "crimp failure",
            "weld joint loosening",
            "solder joint loosening",
            "connector disengagement",
            "clip disengagement",
        ),
        "groups": (
            ("bolt", "loose"),
            ("fastening", "insufficient"),
            ("retention", "insufficient"),
            ("connector", "disengage"),
            ("clip", "disengage"),
        ),
    },
    "3.3.2": {
        "phrases": (
            "poor component contact",
            "insufficient contact area",
            "contact surface contamination",
            "contact surface foreign material",
            "contact air gap",
            "high contact thermal resistance",
            "high electrical contact resistance",
            "solder contact failure",
        ),
        "groups": (
            ("contact", "air gap"),
            ("contact", "resistance", "high"),
            ("contact", "area", "insufficient"),
            ("contact", "foreign material"),
        ),
    },
    "3.4.1": {
        "phrases": (
            "plating defect",
            "coating defect",
            "protective coating failure",
            "insufficient corrosion protection",
        ),
        "groups": (
            ("plating", "corrosion"),
            ("coating", "corrosion"),
        ),
    },
    "3.4.2": {
        "phrases": (
            "moisture ingress",
            "electrolyte ingress",
            "water ingress",
            "liquid accumulation",
            "moisture accumulation",
            "electrolyte accumulation",
        ),
        "groups": (
            ("moisture", "corrosion"),
            ("water", "corrosion"),
            ("electrolyte", "corrosion"),
        ),
    },
    "3.4.3": {
        "phrases": (
            "material reaction prevention failure",
            "brazing flux corrosion",
            "silicone volatilization",
            "silicone outgassing",
            "coolant corrosion by brazing flux",
        ),
        "groups": (
            ("brazing flux", "corrosion"),
            ("silicone", "circuit", "damage"),
            ("silicone", "electronic", "damage"),
        ),
    },
    "3.5.1": {
        "phrases": (
            "insufficient operating performance",
            "relay no contact",
            "relay contact failure",
            "relay stuck",
            "relay sticking",
            "fuse malfunction",
            "fuse no operation",
            "sensing error exceeded",
            "relay noise excessive",
            "fan noise excessive",
        ),
        "groups": (
            ("relay", "stuck"),
            ("relay", "contact", "failure"),
            ("fuse", "malfunction"),
            ("sensing", "error", "exceed"),
        ),
    },
    "3.5.2": {
        "phrases": (
            "insufficient operating life",
            "insufficient durability",
            "relay life failure",
            "relay lifetime not satisfied",
            "sensing error increase over life",
        ),
        "groups": (
            ("lifetime", "not satisfied"),
            ("durability", "insufficient"),
            ("relay", "life", "failure"),
        ),
    },
    "3.6.1": {
        "phrases": (
            "component defect not detected",
            "misassembly not detected",
            "missing part not detected",
            "part omission not detected",
            "pad missing",
            "coolant passage blockage",
            "cooling passage blockage",
        ),
        "groups": (
            ("misassembly", "not detected"),
            ("missing", "not detected"),
            ("field", "misassembly"),
        ),
    },
    "3.6.2": {
        "phrases": (
            "component specification not reflected",
            "component specification missing",
            "specification list omission",
            "warning label missing",
            "grounding design omission",
            "ground design missing",
        ),
        "groups": (
            ("specification", "missing"),
            ("warning label", "missing"),
            ("grounding design", "missing"),
        ),
    },

    "4.1.1": {
        "phrases": (
            "insufficient thermal propagation prevention",
            "insufficient thermal propagation delay",
            "thermal runaway propagation",
            "insufficient external heat exhaust",
            "insufficient adjacent cell insulation",
            "insufficient fire suppression cooling",
        ),
        "groups": (
            ("thermal runaway", "propagation"),
            ("thermal propagation", "insufficient"),
            ("adjacent cell", "insulation", "insufficient"),
        ),
    },
    "4.2.1": {
        "phrases": (
            "insufficient pack enclosure fire resistance",
            "insufficient pack enclosure pressure resistance",
            "insufficient fire resistance",
            "insufficient pressure resistance",
            "flame leakage through gasket",
        ),
        "groups": (
            ("pack enclosure", "fire resistance", "insufficient"),
            ("pack enclosure", "pressure", "insufficient"),
            ("gasket", "flame", "leak"),
        ),
    },

    "5.1.1": {
        "phrases": (
            "insufficient pouch high voltage insulation",
            "insufficient can high voltage insulation",
            "pouch insulation distance insufficient",
            "can insulation distance insufficient",
            "pouch dielectric strength insufficient",
            "can dielectric strength insufficient",
        ),
        "groups": (
            ("pouch", "insulation", "insufficient"),
            ("can", "insulation", "insufficient"),
            ("pouch", "leakage current"),
            ("can", "leakage current"),
        ),
    },
    "5.2.1": {
        "phrases": (
            "insufficient high voltage insulation of live parts",
            "live part insulation distance insufficient",
            "live part dielectric strength insufficient",
            "busbar insulation distance insufficient",
            "busbar dielectric strength insufficient",
        ),
        "groups": (
            ("live part", "insulation", "insufficient"),
            ("busbar", "insulation", "insufficient"),
            ("high voltage", "insulation", "insufficient"),
        ),
    },
    "5.3.1": {
        "phrases": (
            "excessive ground resistance",
            "ground resistance too high",
            "grounding missing",
            "ground missing",
            "ground continuity failure",
            "ground continuity not satisfied",
        ),
        "groups": (
            ("ground", "resistance", "high"),
            ("ground", "continuity", "failure"),
            ("grounding", "missing"),
        ),
    },

    "6.1.1": {
        "phrases": (
            "frequent component management defect",
            "frequent iqc defect",
            "incoming quality defect",
            "defective part incoming",
            "poor incoming quality control",
            "specification alignment insufficient",
        ),
        "groups": (
            ("iqc", "defect"),
            ("incoming", "defective part"),
            ("incoming quality", "defect"),
        ),
    },
    "6.1.2": {
        "phrases": (
            "frequent lqc defect",
            "lqc defect",
            "lqc misassembly",
            "lqc missing part",
            "lqc damage",
        ),
        "groups": (
            ("lqc", "defect"),
            ("lqc", "misassembly"),
            ("lqc", "missing"),
        ),
    },
    "6.1.3": {
        "phrases": (
            "frequent eol defect",
            "eol defect",
            "eol voltage measurement failure",
            "eol current measurement failure",
            "eol resistance measurement failure",
            "eol temperature measurement failure",
            "eol weight measurement failure",
        ),
        "groups": (
            ("eol", "defect"),
            ("eol", "voltage", "failure"),
            ("eol", "measurement", "failure"),
        ),
    },
    "6.1.4": {
        "phrases": (
            "frequent cmm defect",
            "frequent oqc defect",
            "cmm measurement defect",
            "oqc defect",
            "terminal position defect",
            "mounting position defect",
            "cooling surface position defect",
            "overall dimension defect",
        ),
        "groups": (
            ("cmm", "defect"),
            ("oqc", "defect"),
            ("position", "defect"),
        ),
    },
    "6.2.1": {
        "phrases": (
            "defect after product shipment",
            "defect at upper system incoming inspection",
            "damage during transport or storage",
            "contamination after shipment",
            "foreign material after shipment",
            "missing part after shipment",
        ),
        "groups": (
            ("after shipment", "damage"),
            ("after shipment", "contamination"),
            ("upper system", "incoming", "defect"),
            ("transport", "damage"),
        ),
    },
    "6.2.2": {
        "phrases": (
            "defect during upper system assembly",
            "upper system assembly defect",
            "upper system fastening failure",
            "upper system insulation failure",
            "upper system leak test failure",
            "upper system operation failure",
            "upper system recognition failure",
        ),
        "groups": (
            ("upper system", "assembly", "defect"),
            ("upper system", "fastening", "failure"),
            ("upper system", "insulation", "failure"),
        ),
    },
    "6.3.1": {
        "phrases": (
            "surrounding normal component damage",
            "unable to reassemble",
            "reassembly impossible",
            "hook damage during service",
            "clip damage during service",
            "connector damage during service",
        ),
        "groups": (
            ("reassembly", "impossible"),
            ("service", "component", "damage"),
            ("hook", "damage"),
            ("connector", "damage"),
        ),
    },
}


LITTLE_Y_RULES = (
    {
        "code": "1.1", "name": "파우치/캔 크랙",
        "x_codes": ("1.1.1","1.1.2","1.1.3"),
        "phrases": ("파우치 크랙","캔 크랙","파우치/캔 크랙","pouch crack","can crack","cell can crack"),
        "groups": (("파우치","크랙"),("캔","크랙"),("pouch","crack"),("can","crack")),
    },
    {
        "code": "1.2", "name": "파우치/캔 부식",
        "x_codes": ("1.2.1","1.2.2"),
        "phrases": ("파우치 부식","캔 부식","파우치/캔 부식","pouch corrosion","can corrosion"),
        "groups": (("파우치","부식"),("캔","부식"),("pouch","corrosion"),("can","corrosion")),
    },
    {
        "code": "1.3", "name": "파우치/캔 벤트",
        "x_codes": ("1.3.1","1.3.2","1.3.3"),
        "phrases": ("파우치 벤트","캔 벤트","벤트 발생","venting","cell vent","can vent","pouch vent"),
        "groups": (("가스","벤트"),("cell","vent")),
    },
    {
        "code": "2.1", "name": "탭 손상(단선/크랙)",
        "x_codes": ("2.1.1",),
        "phrases": ("탭 손상","탭 단선","탭 크랙","tab damage","tab open","tab break","tab crack"),
        "groups": (("탭","손상"),("탭","단선"),("탭","크랙"),("tab","damage"),("tab","crack")),
    },
    {
        "code": "2.2", "name": "전극 손상(석출/크랙)",
        "x_codes": ("2.2.1","2.2.2"),
        "phrases": ("전극 손상","전극 석출","리튬 석출","전극 크랙","electrode damage","lithium plating","electrode crack"),
        "groups": (("전극","손상"),("전극","석출"),("electrode","damage"),("lithium","plating")),
    },
    {
        "code": "3.1", "name": "팩 부품 누수",
        "x_codes": ("3.1.1","3.1.2"),
        "phrases": ("팩 부품 누수","냉각수 누수","외함 누수","누수 발생","coolant leak","coolant leakage","housing leak","enclosure leak"),
        "groups": (("냉각수","누수"),("외함","누수"),("coolant","leak")),
    },
    {
        "code": "3.2", "name": "팩 부품 손상/과변형",
        "x_codes": ("3.2.1","3.2.2"),
        "phrases": ("팩 부품 손상","팩 부품 파손","팩 부품 과변형","부품 파손","부품 변형","component damage","component deformation","part damage"),
        "groups": (("부품","파손"),("부품","변형"),("component","damage"),("component","deformation")),
    },
    {
        "code": "3.3", "name": "팩 부품 이탈/연결불량",
        "x_codes": ("3.3.1","3.3.2"),
        "phrases": ("팩 부품 이탈","연결불량","연결 불량","부품 이탈","접촉 불량","connection failure","poor connection","component disengagement","contact failure"),
        "groups": (("부품","이탈"),("접촉","불량"),("connection","failure"),("contact","failure")),
    },
    {
        "code": "3.4", "name": "팩 부품 부식",
        "x_codes": ("3.4.1","3.4.2","3.4.3"),
        "phrases": ("팩 부품 부식","부품 부식","전장부품 부식","component corrosion","part corrosion","electrical component corrosion"),
        "groups": (("부품","부식"),("component","corrosion")),
    },
    {
        "code": "3.5", "name": "팩 부품 오작동/미작동",
        "x_codes": ("3.5.1","3.5.2"),
        "phrases": ("팩 부품 오작동","팩 부품 미작동","오작동","미작동","malfunction","no operation","not operating","operation failure"),
        "groups": (("부품","오작동"),("부품","미작동"),("component","malfunction")),
    },
    {
        "code": "3.6", "name": "팩 부품 미검출/미반영",
        "x_codes": ("3.6.1","3.6.2"),
        "phrases": ("팩 부품 미검출","팩 부품 미반영","미검출","사양 미반영","not detected","undetected","specification not reflected"),
        "groups": (("부품","미검출"),("사양","미반영"),("not","detected")),
    },
    {
        "code": "4.1", "name": "열 폭주 전이 불만족",
        "x_codes": ("4.1.1",),
        "phrases": ("열 폭주 전이","열폭주 전이","thermal runaway propagation","thermal propagation"),
        "groups": (("열폭주","전이"),("thermal","propagation")),
    },
    {
        "code": "4.2", "name": "팩 화염 유출",
        "x_codes": ("4.2.1",),
        "phrases": ("팩 화염 유출","화염 유출","flame leakage","flame escape","fire leakage"),
        "groups": (("화염","유출"),("flame","leak")),
    },
    {
        "code": "5.1", "name": "파우치/캔 누전",
        "x_codes": ("5.1.1",),
        "phrases": ("파우치 누전","캔 누전","절연 불량","pouch leakage current","can leakage current","pouch insulation failure","can insulation failure"),
        "groups": (("파우치","누전"),("캔","누전"),("pouch","insulation"),("can","insulation")),
    },
    {
        "code": "5.2", "name": "팩 부품 활선부 누전",
        "x_codes": ("5.2.1",),
        "phrases": ("활선부 누전","고전압 절연 불량","버스바 절연 불량","live part leakage current","high voltage insulation failure","busbar insulation failure"),
        "groups": (("활선부","누전"),("고전압","절연"),("live part","insulation")),
    },
    {
        "code": "5.3", "name": "팩 부품 접지 불량",
        "x_codes": ("5.3.1",),
        "phrases": ("접지 불량","접지 저항 과다","접지 누락","ground failure","grounding failure","high ground resistance","grounding missing"),
        "groups": (("접지","불량"),("접지","저항"),("ground","resistance")),
    },
    {
        "code": "6.1", "name": "팩 수율/직행율 불만족",
        "x_codes": ("6.1.1","6.1.2","6.1.3","6.1.4"),
        "phrases": ("수율 불만족","직행율 불만족","양품 출하 전 불량","생산 불량","yield issue","low yield","production defect","pre-shipment defect"),
        "groups": (("수율","불만족"),("yield","issue"),("production","defect")),
    },
    {
        "code": "6.2", "name": "상위 시스템 공정 불량 발생",
        "x_codes": ("6.2.1","6.2.2"),
        "phrases": ("상위 시스템 공정 불량","차량 공정 불량","입고단 불량","상위 시스템 조립 불량","upper system defect","vehicle assembly defect","incoming defect","upper system assembly defect"),
        "groups": (("상위 시스템","불량"),("차량","공정","불량"),("upper system","defect")),
    },
    {
        "code": "6.3", "name": "필드 서비스 불만족",
        "x_codes": ("6.3.1",),
        "phrases": ("필드 서비스 불만족","서비스 재조립 불가","field service issue","service reassembly failure","reassembly impossible"),
        "groups": (("필드","서비스"),("field","service")),
    },
)


def _little_y_score(rule, symptom_q):
    score = 0
    for phrase in rule.get("phrases", ()):
        p = _key(phrase)
        if p and p in symptom_q:
            score = max(score, 120 + min(len(p), 40))
    for group in rule.get("groups", ()):
        parts = [_key(x) for x in group if _key(x)]
        if parts and all(x in symptom_q for x in parts):
            score = max(score, 100 + 10 * len(parts))
    return score


def classify_little_y(problem_text):
    """Return likely Little y categories from 2D symptom text.

    Output is ordered by confidence. Multiple entries are preserved when the
    symptom itself is ambiguous, so the next-stage Y.y.x recommendation can
    evaluate causes within the union of those categories.
    """
    q = _key(problem_text)
    if not q:
        return []
    scored = []
    for rule in LITTLE_Y_RULES:
        score = _little_y_score(rule, q)
        if score:
            scored.append((score, rule))
    scored.sort(key=lambda x: (-x[0], x[1]["code"]))
    if not scored:
        return []
    best = scored[0][0]
    # Keep close alternatives only; weak unrelated categories are discarded.
    return [r for s, r in scored if s >= best - 15]


def little_y_label(code):
    for rule in LITTLE_Y_RULES:
        if rule["code"] == code:
            return f'{rule["code"]}  {rule["name"]}'
    return str(code or "")


RULES = (
    _rule("1.1.1", "간섭 및 외력 작용 과다",
          ("간섭 및 외력 작용 과다", "가스켓 간섭에 의한 크랙", "가스켓 간섭에 의한 벤트"),
          (("공정", "크랙"), ("기계시험", "크랙"), ("외력", "크랙"), ("간섭", "크랙"))),
    _rule("1.1.2", "내부 전극 유동 과다",
          ("내부 전극 유동 과다", "전극 돌출")),
    _rule("1.1.3", "레이저 용접 중 손상",
          ("레이저 용접 중 손상", "레이저 오조사", "과용입"),
          (("레이저", "관통"), ("레이저", "두께부족"), ("레이저", "실링부 손상"), ("레이저", "과열"))),
    _rule("1.2.1", "Anode Connection",
          ("Anode Connection", "리드 오삽입", "리드 밀림", "셀 간 접촉")),
    _rule("1.2.2", "버스바/리드/와이어 용접부/접합부 부식",
          ("용접부 부식", "접합부 부식"),
          (("버스바", "부식"), ("리드", "부식"), ("와이어", "부식"),
           ("액침 냉매", "부식"), ("절연유", "부식"), ("포팅 레진", "부식"))),
    _rule("1.3.1", "외부 열 유입에 의한 가스 생성 과다",
          ("외부 열 유입에 의한 가스 생성 과다",),
          (("외부 열", "가스 생성"), ("터미널", "발열"), ("버스바", "발열"), ("회로", "발열"))),
    _rule("1.3.2", "내부 발열에 의한 가스 생성 과다",
          ("내부 발열에 의한 가스 생성 과다",),
          (("내부 발열", "가스 생성"), ("단락시험", "과열"))),
    _rule("1.3.3", "냉각 성능 부족에 의한 가스 생성 과다",
          ("냉각 성능 부족에 의한 가스 생성 과다", "냉각 성능 부족"),
          (("수명", "과열"), ("성능시험", "과열"))),

    _rule("2.1.1", "탭 인장력 유발 과다",
          ("탭 인장력 유발 과다", "탭 인장력", "인장력 과다")),
    _rule("2.2.1", "가압력/공간 부적절",
          ("가압력 공간 부적절", "가압력 부적절", "공간 부적절",
           "전극 계면 Li 석출", "리튬 석출")),
    _rule("2.2.2", "전압/온도 센싱 부적절",
          ("전압 온도 센싱 부적절", "전압 센싱 부적절", "온도 센싱 부적절",
           "전극 저온부 Li 석출", "저온부 리튬 석출")),

    _rule("3.1.1", "냉각수 가스켓/용접부 실링 불량",
          ("냉각수 가스켓", "용접부 실링 불량", "오링 누수", "O-ring 누수",
           "호스 누수", "클램프 누수", "피팅 누수", "냉각 플레이트 브레이징", "냉각수 누수"),
          (("냉각수", "가스켓"), ("냉각수", "용접부"), ("냉각", "누수"))),
    _rule("3.1.2", "외함 가스켓/용접부 실링 불량",
          ("외함 가스켓", "외함 실링 불량", "프레임 용접부 누수"),
          (("팩 커버", "가스켓"), ("커넥터", "가스켓"), ("볼트", "가스켓"), ("프레임", "누수"))),
    _rule("3.2.1", "부품 간격/강성 불량",
          ("부품 간격 강성 불량", "강성 불량"),
          (("부품", "간격"), ("간섭", "파손"), ("하네스", "파손"),
           ("브라켓", "위치"), ("스웰링", "발화"))),
    _rule("3.2.2", "부품 발열/내열 불량",
          ("부품 발열 내열 불량", "내열 불량", "허용전류 초과", "피복 열화", "사출물 열화"),
          (("부품", "발열"), ("부품", "과열"))),
    _rule("3.3.1", "부품간 체결력/지지력 불량",
          ("부품간 체결력 지지력 불량", "체결력 불량", "지지력 불량", "볼트 풀림",
           "링 터미널 회전", "크림프 불량", "융착부 불량", "솔더부 불량"),
          (("볼트", "빠짐"), ("클립", "빠짐"), ("커넥터", "빠짐"),
           ("후크", "빠짐"), ("용접부", "풀림"), ("체결", "풀림"))),
    _rule("3.3.2", "부품간 접촉 불량",
          ("부품간 접촉 불량", "접촉면적 부족", "접촉면 이물", "접촉면 Air Gap",
           "접촉면 열저항", "전기저항 불량", "솔더 불량"),
          (("부품", "접촉불량"), ("접촉", "이물"))),
    _rule("3.4.1", "도금/도장/보호 불량",
          ("도금 도장 보호 불량", "도금 불량", "도장 불량", "보호 불량"),
          (("도금", "부식"), ("도장", "부식"))),
    _rule("3.4.2", "수분/전해액 유입/고임 방지 불량",
          ("수분 전해액 유입 고임 방지 불량", "수분 유입", "전해액 유입",
           "고임 방지 불량", "수분 고임", "전해액 고임"),
          (("수분", "부식"), ("전해액", "부식"))),
    _rule("3.4.3", "재질 발반응 방지 불량",
          ("재질 발반응 방지 불량", "브레이징 플럭스 냉각수 부식", "실리콘 휘발"),
          (("브레이징 플럭스", "부식"), ("절연유", "부품 부식"),
           ("실리콘", "전장품 손상"), ("실리콘", "회로 손상"))),
    _rule("3.5.1", "작동 성능 부족",
          ("작동 성능 부족", "릴레이 미접촉", "릴레이 탈락", "릴레이 고착",
           "퓨즈 오작동", "퓨즈 미작동", "센싱 오차 초과", "릴레이 소음 초과", "팬 소음 초과")),
    _rule("3.5.2", "작동 수명/내구 부족",
          ("작동 수명 내구 부족", "릴레이 수명 불만족", "센싱 오차 증가"),
          (("수명", "불만족"), ("내구", "부족"))),
    _rule("3.6.1", "부품 불량/오조립/누락 미검출",
          ("부품 불량 오조립 누락 미검출", "오조립 미검출", "누락 미검출",
           "부품 불량 미검출", "패드 누락", "냉각 유로 막힘"),
          (("출하 후", "오조립"), ("필드", "오조립"))),
    _rule("3.6.2", "부품사양 미반영",
          ("부품사양 미반영", "부품사양 누락", "사양 목록 누락", "경고 라벨 누락", "접지 설계 누락")),

    _rule("4.1.1", "열 폭주 전이 방지/지연 부족",
          ("열 폭주 전이 방지 지연 부족", "열 폭주 전이 방지 부족", "열 폭주 전이 지연 부족",
           "외부 열 배출 부족", "인접 셀 단열 부족", "소화 냉각 부족"),
          (("TP시험", "셀 외부 단락"), ("열폭주", "전이"))),
    _rule("4.2.1", "팩 외함 부품 내화/내압 부족",
          ("팩 외함 부품 내화 내압 부족", "내화 부족", "내압 부족", "가스켓 화염 유출"),
          (("TP시험", "외함 부품", "소손"), ("TP시험", "외함 부품", "이탈"),
           ("TP시험", "외함 부품", "과변형"))),

    _rule("5.1.1", "파우치/캔 고전압 절연 부족",
          ("파우치 캔 고전압 절연 부족",),
          (("파우치", "절연거리 부족"), ("캔", "절연거리 부족"),
           ("파우치", "절연강도 부족"), ("캔", "절연강도 부족"),
           ("파우치", "누전"), ("캔", "누전"))),
    _rule("5.2.1", "활선부 고전압 절연 부족",
          ("활선부 고전압 절연 부족",),
          (("활선부", "절연거리 부족"), ("활선부", "절연강도 부족"),
           ("버스바", "절연거리 부족"), ("버스바", "절연강도 부족"),
           ("팩 부품", "절연거리 부족"), ("팩 부품", "절연강도 부족"),
           ("활선부", "누전"))),
    _rule("5.3.1", "접지 저항 과다",
          ("접지 저항 과다", "접지 저항 초과", "접지 떨어짐", "접지 누락", "접지 연속성 불만족")),

    _rule("6.1.1", "부품 관리/IQC 빈발 불량",
          ("부품 관리 IQC 빈발 불량", "IQC 빈발 불량", "부품 관리 불량",
           "수율 불만족", "불량품 유입", "사양 사전 정합화 미흡"),
          (("부품", "불량품 유입"), ("셀", "불량품 유입"), ("IQC", "불량"))),
    _rule("6.1.2", "LQC 빈발 불량",
          ("LQC 빈발 불량",),
          (("LQC", "불량"), ("LQC", "오조립"), ("LQC", "누락"), ("LQC", "파손"))),
    _rule("6.1.3", "EOL 빈발 불량",
          ("EOL 빈발 불량",),
          (("EOL", "불량"), ("EOL", "전압"), ("EOL", "전류"),
           ("EOL", "저항"), ("EOL", "온도"), ("EOL", "무게"))),
    _rule("6.1.4", "CMM/OQC 빈발 불량",
          ("CMM OQC 빈발 불량", "CMM 빈발 불량", "OQC 빈발 불량"),
          (("CMM", "불량"), ("OQC", "불량"), ("터미널", "위치 불량"),
           ("마운팅", "위치 불량"), ("냉각면", "위치 불량"), ("외곽 사이즈", "불량"))),
    _rule("6.2.1", "제품 출하 후 및 상위 시스템 입고 단 불량",
          ("제품 출하 후 및 상위 시스템 입고 단 불량", "상위 시스템 입고 단 불량"),
          (("출하 후", "오염"), ("출하 후", "이물"), ("출하 후", "손상"),
           ("출하 후", "탈락"), ("출하 후", "변형"), ("출하 후", "누락"),
           ("운송 보관", "손상"), ("입고", "불량"))),
    _rule("6.2.2", "상위 시스템 조립 공정 중 불량",
          ("상위 시스템 조립 공정 중 불량",),
          (("상위 시스템", "체결 불량"), ("상위 시스템", "절연 불량"),
           ("상위 시스템", "기밀 불량"), ("상위 시스템", "측정 불량"),
           ("상위 시스템", "작동 불량"), ("상위 시스템", "인식 불량"),
           ("상위 시스템", "부품 탈락"), ("상위 시스템", "부품 파손"))),
    _rule("6.3.1", "주변 정상 부품 파손/재조립 불가",
          ("주변 정상 부품 파손 재조립 불가", "재조립 불가"),
          (("고정 후크", "파손"), ("클립", "파손"), ("커넥터", "파손"),
           ("스웰링", "조립 불가"), ("주변 정상 부품", "파손"))),
)


def _matches(rule, q):
    for phrase in (rule["name"],) + rule["phrases"]:
        p = _key(phrase)
        if p and p in q:
            return True
    for group in rule["groups"]:
        parts = [_key(x) for x in group if _key(x)]
        if parts and all(x in q for x in parts):
            return True

    # English 8D reports use the same Y.y.x design-rule taxonomy.
    # Keep the English vocabulary separate so Korean rule maintenance remains
    # readable while bilingual/mixed-language cause text is matched identically.
    english = ENGLISH_RULE_ALIASES.get(rule["code"], {})
    for phrase in english.get("phrases", ()):
        p = _key(phrase)
        if p and p in q:
            return True
    for group in english.get("groups", ()):
        parts = [_key(x) for x in group if _key(x)]
        if parts and all(x in q for x in parts):
            return True
    return False


def classify_little_x(cause_text, allowed_codes=None):
    """Return matching Y.y.x codes in design-rule order.

    When allowed_codes is supplied, cause matching is restricted to the
    Little y categories selected from the 2D symptom.
    """
    q = _key(cause_text)
    if not q:
        return []
    allowed = set(allowed_codes or ())
    return [
        rule["code"] for rule in RULES
        if (not allowed or rule["code"] in allowed) and _matches(rule, q)
    ]


def _column_header_text(ws, col, header_rows=7):
    return " ".join(_text(ws.cell(r, col).value) for r in range(1, min(header_rows, ws.max_row) + 1))


def find_rule_columns(ws, header_rows=7):
    """Locate rule-violation Y/N and Y.y.x columns by header text."""
    violation = None
    little_x = None
    best_v = -1
    best_x = -1

    for col in range(1, ws.max_column + 1):
        raw = _column_header_text(ws, col, header_rows)
        q = _key(raw)

        vscore = 0
        if "규칙위반" in q or "설계규칙위반" in q:
            vscore = 100
        elif "위반여부" in q and ("규칙" in q or "설계" in q):
            vscore = 90
        elif "규칙" in q and "위반" in q:
            vscore = 70

        xscore = 0
        if q == "yyx" or "yyx" in q:
            xscore = 110
        elif "littlex" in q or "리틀x" in q:
            xscore = 100
        elif "little" in q and q.endswith("x"):
            xscore = 90
        elif "규칙번호" in q or "규칙no" in q:
            xscore = 75
        elif "팩기인원인유형" in q:
            xscore = 65

        if vscore > best_v:
            best_v, violation = vscore, col
        if xscore > best_x:
            best_x, little_x = xscore, col

    if best_v <= 0:
        violation = None
    if best_x <= 0:
        little_x = None
    if violation == little_x:
        little_x = None
    return violation, little_x


def write_rule_classification(ws, row, cause_text):
    """Write Y/N and matching Y.y.x code(s). Returns diagnostic result."""
    codes = classify_little_x(cause_text)
    violation_col, little_x_col = find_rule_columns(ws)

    if violation_col is not None:
        ws.cell(row, violation_col).value = "Y" if codes else "N"
    if little_x_col is not None:
        ws.cell(row, little_x_col).value = ", ".join(codes) if codes else ""

    return {
        "violation": "Y" if codes else "N",
        "codes": codes,
        "violation_col": violation_col,
        "little_x_col": little_x_col,
    }


NONE_CODE = "__NONE__"


def rule_by_code(code):
    for rule in RULES:
        if rule["code"] == code:
            return rule
    return None


def rule_label(code):
    if code == NONE_CODE:
        return "해당 없음 (N)"
    rule = rule_by_code(code)
    return f'{rule["code"]}  {rule["name"]}' if rule else str(code or "")


def recommendation(problem_text, cause_text=None):
    """Two-stage recommendation: 2D symptom -> Little y -> cause -> Y.y.x.

    Backward compatibility: recommendation(cause_text) still works when only
    one positional argument is supplied.
    """
    if cause_text is None:
        cause_text = problem_text
        problem_text = ""

    little_y = classify_little_y(problem_text)
    allowed_codes = []
    for y in little_y:
        for code in y["x_codes"]:
            if code not in allowed_codes:
                allowed_codes.append(code)

    # Stage 2: cause matching is constrained by 2D-derived Little y whenever
    # the symptom gives a meaningful category. If 2D is too vague, fall back
    # to cause-only matching across all 36 rules.
    codes = classify_little_x(cause_text, allowed_codes=allowed_codes or None)

    # If 2D clearly narrows to a Little y but the cause wording does not match
    # an x phrase, preserve the narrowed x choices as candidates rather than
    # inventing an automatic final answer.
    inferred_from_y_only = False
    if not codes and allowed_codes:
        codes = list(allowed_codes)
        inferred_from_y_only = True

    y_labels = [little_y_label(y["code"]) for y in little_y]
    y_reason = (
        "2D 현상 기준 Little y 후보: " + " / ".join(y_labels)
        if y_labels else
        "2D 현상에서 Little y를 명확히 좁히지 못해 전체 Y.y.x에서 원인을 비교했습니다."
    )

    if not codes:
        return {
            "recommended": NONE_CODE,
            "candidates": [],
            "ambiguous": False,
            "little_y": [y["code"] for y in little_y],
            "little_y_labels": y_labels,
            "reason": y_reason + "\n원인 내용에서도 명확히 일치하는 Y.y.x를 찾지 못했습니다.",
        }

    if inferred_from_y_only:
        return {
            "recommended": codes[0] if len(codes) == 1 else codes[0],
            "candidates": codes,
            "ambiguous": len(codes) > 1,
            "little_y": [y["code"] for y in little_y],
            "little_y_labels": y_labels,
            "reason": y_reason + "\n원인 문구만으로 세부 Y.y.x가 확정되지 않아 해당 Little y 내부 후보를 표시합니다.",
        }

    if len(codes) == 1:
        code = codes[0]
        return {
            "recommended": code,
            "candidates": codes,
            "ambiguous": False,
            "little_y": [y["code"] for y in little_y],
            "little_y_labels": y_labels,
            "reason": y_reason + f'\n원인 내용까지 종합하면 {rule_label(code)}가 가장 일치합니다.',
        }

    return {
        "recommended": codes[0],
        "candidates": codes,
        "ambiguous": True,
        "little_y": [y["code"] for y in little_y],
        "little_y_labels": y_labels,
        "reason": y_reason + "\n원인 내용상 여러 Y.y.x가 가능하므로 실제 원인에 해당하는 하나를 선택해 주세요.",
    }


def write_confirmed_rule_classification(ws, row, selected_code, confirmed=True):
    """Write only a user-confirmed design-rule decision.

    selected_code == NONE_CODE writes N and clears Y.y.x.
    A concrete Y.y.x writes Y and that ONE code only.
    If confirmed is False, nothing is changed.
    """
    violation_col, little_x_col = find_rule_columns(ws)
    if not confirmed:
        return {
            "violation": None,
            "codes": [],
            "violation_col": violation_col,
            "little_x_col": little_x_col,
        }

    code = _text(selected_code)
    if code == NONE_CODE or not code:
        violation = "N"
        code_text = ""
        codes = []
    else:
        rule = rule_by_code(code)
        if rule is None:
            violation = "N"
            code_text = ""
            codes = []
        else:
            violation = "Y"
            code_text = code
            codes = [code]

    if violation_col is not None:
        ws.cell(row, violation_col).value = violation
    if little_x_col is not None:
        ws.cell(row, little_x_col).value = code_text

    return {
        "violation": violation,
        "codes": codes,
        "violation_col": violation_col,
        "little_x_col": little_x_col,
    }
