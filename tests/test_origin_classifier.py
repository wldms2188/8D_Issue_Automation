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

ENGLISH_TEMPLATES={
 '부품':[
  'Supplier component defect due to material variation',
  'Incoming part had a dimensional variation from the supplier',
  'Raw material property variation caused the cell defect',
  'Part variation was confirmed as the root cause',
 ],
 '설계':[
  'Insufficient design margin caused structural interference',
  'Tolerance stack-up created insufficient clearance',
  'Drawing dimension error caused the interference',
  'Specification and geometry were inadequate for the load',
 ],
 '공정':[
  'Assembly torque variation caused the failure',
  'Welding condition variation generated the defect',
  'Fixture alignment and process parameter were incorrect',
  'Coating condition and equipment setup caused the issue',
 ],
 '기타':[
  'External impact occurred during transportation',
  'Improper storage environmental condition caused damage',
  'Customer mishandling during use caused the damage',
  'Packaging damage occurred during shipment handling',
 ],
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

    def test_large_combinatorial_clear_cases(self):
        prefixes=('확인 결과 ','분석 결과 ','4D 원인: ','재현시험 결과 ')
        suffixes=(' 확인됨',' 영향으로 발생',' 원인으로 판단',' 재현됨')
        qualifiers=('',' 반복',' 특정 LOT',' DV 단계')
        total=0
        for expected,causes in TEMPLATES.items():
            for cause in causes:
                for pre in prefixes:
                    for suf in suffixes:
                        for qual in qualifiers:
                            rec,_=clf.recommend_origin({'cause_4d':pre+cause+qual+suf})
                            self.assertEqual(rec,expected,msg=(expected,rec,pre+cause+qual+suf))
                            total+=1
        self.assertEqual(total,768)

    def test_english_160_clear_cases(self):
        total=correct=0
        for expected,causes in ENGLISH_TEMPLATES.items():
            for i in range(40):
                text=causes[i%len(causes)]+f' / case {i+1}'
                rec,_=clf.recommend_origin({'cause_4d':text})
                total+=1
                if rec==expected:correct+=1
        self.assertEqual(total,160)
        self.assertEqual(correct,160)

    def test_large_english_combinatorial_clear_cases(self):
        prefixes=(
            'Root cause analysis confirmed that ',
            'Reproduction test showed that ',
            'Investigation concluded that ',
            'Failure analysis identified that ',
            '4D analysis determined that ',
        )
        suffixes=(
            ' as the root cause.',
            ' caused the observed failure.',
            ' resulted in the defect.',
            ' was reproduced under the same condition.',
            ' directly contributed to the issue.',
        )
        qualifiers=(
            '',
            ' during DV evaluation',
            ' under repeated cycling',
            ' on a specific lot',
        )
        total=0
        for expected,causes in ENGLISH_TEMPLATES.items():
            for cause in causes:
                for pre in prefixes:
                    for suf in suffixes:
                        for qual in qualifiers:
                            text=pre+cause+qual+suf
                            rec,_=clf.recommend_origin({'cause_4d':text})
                            self.assertEqual(rec,expected,msg=(expected,rec,text))
                            total+=1
        self.assertEqual(total,1600)

    def test_english_negation_before_and_after_term(self):
        cases=[
            ('설계','Process variation was ruled out. Insufficient design margin caused the issue.'),
            ('설계','Not caused by assembly process; drawing tolerance was incorrect.'),
            ('공정','No issue with component material. Assembly torque variation was confirmed.'),
            ('부품','Design was within specification. Supplier component defect was confirmed.'),
            ('기타','Process condition was normal. External impact during transport caused damage.'),
        ]
        for expected,text in cases:
            rec,_=clf.recommend_origin({'cause_4d':text})
            self.assertEqual(rec,expected,msg=(expected,rec,text))

    def test_english_mixed_evidence_stays_discussion_needed(self):
        cases=[
            'Design tolerance and assembly torque both contributed to the failure.',
            'Supplier part variation and welding condition together caused the issue.',
            'Transport impact and component defect were both observed as causes.',
        ]
        for text in cases:
            rec,_=clf.recommend_origin({'cause_4d':text})
            self.assertEqual(rec,'논의 중',msg=(rec,text))

    def test_occurrence_cause_wins_over_escape_and_system_text(self):
        d={
            'cause_4d':'Supplier component defect due to material variation',
            'leak_cause':'Inspection process failed to detect the defect',
            'system_cause':'Control plan and work instruction were not linked',
        }
        rec,_=clf.recommend_origin(d)
        self.assertEqual(rec,'부품')

        d={
            'cause_4d':'Insufficient design margin caused interference',
            'leak_cause':'Operator inspection missed the defect',
            'system_cause':'Manufacturing control process was inadequate',
        }
        rec,_=clf.recommend_origin(d)
        self.assertEqual(rec,'설계')

    def test_supporting_causes_are_used_only_when_occurrence_has_no_category_clue(self):
        d={
            'cause_4d':'Root cause under additional review',
            'leak_cause':'External impact during transport was identified',
            'system_cause':'',
        }
        rec,_=clf.recommend_origin(d)
        self.assertEqual(rec,'기타')

    def test_ambiguous_design_process_is_not_forced(self):
        rec,_=clf.recommend_origin({'cause_4d':'설계 공차와 체결토크 산포가 동시에 영향'})
        self.assertEqual(rec,'논의 중')

    def test_empty_cause_is_tbd(self):
        rec,_=clf.recommend_origin({'cause_4d':'','leak_cause':'','system_cause':''})
        self.assertEqual(rec,'TBD')

    def test_negated_process_phrase_does_not_force_process(self):
        rec,_=clf.recommend_origin({'cause_4d':'공정조건 문제없음. 설계마진 부족 확인'})
        self.assertEqual(rec,'설계')

    def test_english_causes_use_same_categories(self):
        cases={
            '부품':'Supplier component defect due to material variation',
            '설계':'Insufficient design margin and tolerance interference',
            '공정':'Assembly process condition caused torque variation',
            '기타':'External impact during transport and storage',
        }
        for expected,text in cases.items():
            rec,_=clf.recommend_origin({'cause_4d':text})
            self.assertEqual(rec,expected,msg=(expected,rec,text))

    def test_english_negation(self):
        rec,_=clf.recommend_origin({'cause_4d':'Process condition no issue. Design margin insufficient.'})
        self.assertEqual(rec,'설계')

if __name__=='__main__': unittest.main()
