"""Synthetic regression validation for the issue-origin classifier.
Run: py -m unittest tests.test_origin_classifier -v
This is regression coverage, not a claim of real-world accuracy.
"""
import sys
from pathlib import Path
import unittest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import origin_classifier_final as clf

TEMPLATES={
 '부품':[
  ('협력사 입고 부품 LOT 편차로 부품불량 발생','협력사 선별 및 부품교체 적용'),
  ('원자재 소재불량으로 셀불량 발생','업체개선 및 수입검사 강화'),
  ('부품 편차가 원인으로 확인됨','협력사 공정 개선 요청'),
 ],
 '설계':[
  ('설계마진 부족 및 구조 간섭이 발생원인으로 확인됨','구조변경 및 설계변경 적용'),
  ('도면 공차 부적정으로 치수 간섭 발생','공차변경 및 도면변경'),
  ('사양 설계 미흡으로 강성 부족','사양변경 및 설계반영'),
 ],
 '공정':[
  ('체결토크 산포로 조립불량 발생','작업조건 및 작업표준 변경'),
  ('용접조건 편차로 용접불량 발생','용접조건 변경 및 설비조건 관리'),
  ('도포조건 관리 미흡으로 공정산포 발생','공정조건 변경 및 검사 강화'),
 ],
 '기타':[
  ('운송중 외부충격으로 파손 발생','운송조건 변경'),
  ('보관조건 부적정 환경조건 노출','보관조건 개선'),
  ('고객사용 중 취급부주의로 손상','사용조건 안내 및 취급 교육'),
 ],
}

class OriginClassifierSyntheticTests(unittest.TestCase):
    def test_120_clear_cases(self):
        total=correct=0; errors=[]
        for expected,pairs in TEMPLATES.items():
            for i in range(30):
                cause,action=pairs[i%len(pairs)]
                d={'cause_4d':cause+f' LOT{i+1}','leak_cause':'','system_cause':'',
                   'action_5d':action,'verification_6d':'효과 검증 완료',
                   'problem':f'DUT #{i%5+1} 현상 발생','occurrence_site':'','task_name':'','issue_name':''}
                r=clf.classify(d); total+=1
                if r.recommendation==expected:correct+=1
                else:errors.append((expected,r.recommendation,r.reason))
        self.assertEqual(total,120)
        self.assertEqual(correct,120,msg=str(errors[:10]))

    def test_ambiguous_design_process_is_not_forced(self):
        d={'cause_4d':'설계 공차와 체결토크 산포가 동시에 영향','action_5d':'설계변경 및 공정조건 변경'}
        r=clf.classify(d)
        self.assertEqual(r.recommendation,'논의 중')

    def test_unknown_cause_is_tbd(self):
        self.assertEqual(clf.classify({'cause_4d':'TBD'}).recommendation,'TBD')

    def test_negated_process_phrase_does_not_force_process(self):
        d={'cause_4d':'공정조건 문제없음. 설계마진 부족 확인','action_5d':'설계변경 적용'}
        self.assertEqual(clf.classify(d).recommendation,'설계')

if __name__=='__main__': unittest.main()
