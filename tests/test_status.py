from app.main import status


def test_6d_verification_complete():
    assert status({'verification_6d':'검증 완료','cause_4d':'','action_5d':''}) == '개선 완료'


def test_6d_complete():
    assert status({'verification_6d':'완료','cause_4d':'','action_5d':''}) == '개선 완료'


def test_completion_planned_is_not_complete():
    assert status({'verification_6d':'완료 예정','cause_4d':'원인','action_5d':'대책'}) == '개선 검증중'


def test_unknown_when_4d_5d_6d_empty():
    assert status({'verification_6d':'','cause_4d':'','action_5d':''}) == '원인/개선 미확인'
