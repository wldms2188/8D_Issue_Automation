"""Synthetic regression validation for the stable refined V1 issue-origin recommender.
Run: py -m unittest tests.test_origin_classifier -v
Synthetic regression coverage only; human confirmation remains the final decision.
"""
import sys
from pathlib import Path
import unittest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import origin_v1_refined_patch as clf

TEMPLATES={
 '부품':['협력사 입고 부품 LOT 편차로 부품불량 발생','원자재 소재불량으로 셀불량 발생','부품 편차가 원인으로 확인됨'],
 '설계':['설계마진 부족 및 구조 간섭이 발생원인으로 확인됨','도면 공차 부적정으로 치수 간섭 발생','사양 설계 미흡으로 강성 부족'],
 '공정':['체결토크 산포로 조립불량 발생','용접조건 편차로 용접불량 발생','도포조건 관리 미흡으로 공정산포 발생'],
 '기타':['운송중 외부충격으로 파손 발생','보관조건 부적정 환경조건 노출','고객사용 중 취급부주의로 손상'],
}

class OriginClassifierSyntheticTests(unittest.TestCase):
    def test_120_clear_cases(self):
        total=correct=0; errors=[]
        for expected,causes in TEMPLATES.items():
            for i in range(30):
                d={'cause_4d':causes[i%len(causes)]+f' LOT{i+1}','leak_cause':'','system_cause':''}
                rec,reason=clf.recommend_origin(d); total+=1
                if rec==expected: correct+=1
                else: errors.append((expected,rec,reason,d['cause_4d']))
        self.assertEqual(total,120)
        self.assertEqual(correct,120,msg=str(errors[:10]))

    def test_ambiguous_design_process_is_not_forced(self):
        rec,_=clf.recommend_origin({'cause_4d':'설계 공차와 체결토크 산포가 동시에 영향'})
        self.assertEqual(rec,'논의 중')

    def test_empty_cause_is_tbd(self):
        rec,_=clf.recommend_origin({'cause_4d':'','leak_cause':'','system_cause':''})
        self.assertEqual(rec,'TBD')

    def test_negated_process_phrase_does_not_force_process(self):
        rec,_=clf.recommend_origin({'cause_4d':'공정조건 문제없음. 설계마진 부족 확인'})
        self.assertEqual(rec,'설계')

if __name__=='__main__': unittest.main()
