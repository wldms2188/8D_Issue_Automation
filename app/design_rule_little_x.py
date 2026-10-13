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
    return False


def classify_little_x(cause_text):
    """Return matching Y.y.x codes in design-rule order."""
    q = _key(cause_text)
    if not q:
        return []
    return [rule["code"] for rule in RULES if _matches(rule, q)]


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


def recommendation(cause_text):
    """Return recommendation payload for the confirmation dialog."""
    codes = classify_little_x(cause_text)
    if not codes:
        return {
            "recommended": NONE_CODE,
            "candidates": [],
            "ambiguous": False,
            "reason": "원인 내용에서 설계규칙 Y.y.x와 명확히 일치하는 표현을 찾지 못했습니다.",
        }
    if len(codes) == 1:
        code = codes[0]
        return {
            "recommended": code,
            "candidates": codes,
            "ambiguous": False,
            "reason": f'원인 내용이 {rule_label(code)} 규칙과 일치하여 추천합니다.',
        }
    return {
        "recommended": codes[0],
        "candidates": codes,
        "ambiguous": True,
        "reason": "원인 내용이 여러 Y.y.x 규칙과 동시에 유사합니다. 후보 중 실제 원인에 해당하는 하나를 선택해 주세요.",
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
