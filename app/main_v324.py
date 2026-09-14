# 8D Issue Automation v3.2.4
# Status reason explicitly includes the actual 6D text.
import re

import main_v323 as v323
import main_v322 as v322

base = v323.base
N = v322.N
C = v322.C


def _status_reason(d):
    raw=N(d.get('verification_6d'))
    if not raw:
        return 'open', '6D 내용: (미기재)\n→ 6D 내용이 기재되어 있지 않아 open으로 판단했습니다.'

    progress_patterns=(
        r'진행\s*중', r'검증\s*중', r'확인\s*중', r'모니터링\s*중',
        r'시험\s*중', r'적용\s*중', r'조치\s*중', r'분석\s*중',
        r'예정', r'계획', r'추후', r'진행중', r'검증중', r'확인중', r'모니터링중'
    )
    for pat in progress_patterns:
        m=re.search(pat,raw,re.I)
        if m:
            return 'open', (
                f'6D 내용: {raw}\n'
                f'→ 진행/예정 표현("{m.group(0)}")이 있어 open으로 판단했습니다.'
            )

    if re.search(r'완료',raw,re.I):
        return 'close', (
            f'6D 내용: {raw}\n'
            '→ 6D에 "완료" 표현이 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'
        )
    if re.search(r'확인',raw,re.I):
        return 'close', (
            f'6D 내용: {raw}\n'
            '→ 6D에 "확인" 표현이 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'
        )

    return 'close', (
        f'6D 내용: {raw}\n'
        '→ 6D에 검증/확인 결과 내용이 기재되어 있고 진행 중/예정 표현이 없어 close로 판단했습니다.'
    )


# Patch v3.2.2 status logic used by preview, Excel status, and close confirmation dialog.
v322._status_reason=_status_reason


class App(v323.App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.4')


if __name__=='__main__':
    App().mainloop()
