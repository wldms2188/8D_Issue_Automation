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

    def test_bilingual_process_clues_include_bolt_loosening_and_fastening(self):
        korean=[
            '볼트 풀림이 발생했으며 체결토크 부족이 원인으로 확인됨',
            '체결 미흡으로 조립 중 풀림 발생',
            '너트 풀림 및 체결불량이 확인됨',
        ]
        english=[
            'Bolt loosening occurred due to insufficient fastening torque.',
            'Scratch caused by loosening of the unloading hoist bracket fixing bolt.',
            'Fastener loosening was reproduced after under-torque assembly.',
            'Loose bolt was confirmed after the fastening process.',
        ]
        for text in korean+english:
            rec,_=clf.recommend_origin({'cause_4d':text})
            self.assertEqual(rec,'공정',msg=(rec,text))

    def test_bilingual_design_clues_include_spec_and_tolerance_language(self):
        korean=[
            '설계 스펙 미흡으로 공차 간섭 발생',
            '규격 미흡 및 치수 부적합이 원인',
            '설계 요구사항과 실제 공차 조건이 불일치',
        ]
        english=[
            'Design specification was insufficient for the required clearance.',
            'Tolerance stack-up exceeded the design spec limit.',
            'Design requirement and geometry caused interference.',
        ]
        for text in korean+english:
            rec,_=clf.recommend_origin({'cause_4d':text})
            self.assertEqual(rec,'설계',msg=(rec,text))

    def test_occurrence_site_reason_does_not_expose_screen_or_ppt_source(self):
        for site,expected in [('제품 생산','공정'),('부품 생산','부품')]:
            rec,reason=clf.recommend_origin({
                'cause_4d':'원인 추가 검토 중',
                'occurrence_site':'다른 값',
                '_origin_occurrence_site':site,
            })
            self.assertEqual(rec,expected)
            self.assertIn(f'"{site}" 중 발생한 이슈이며',reason)
            self.assertNotIn('화면 선택값',reason)
            self.assertNotIn('8D 원문값',reason)

    def test_gui_occurrence_site_overrides_conflicting_ppt_site(self):
        d={
            'cause_4d':'원인 추가 검토 중',
            'occurrence_site':'부품 생산',
            '_origin_occurrence_site':'제품 생산',
        }
        rec,reason=clf.recommend_origin(d)
        self.assertEqual(rec,'공정')
        self.assertIn('중 발생한 이슈이며',reason)

        d={
            'cause_4d':'원인 추가 검토 중',
            'occurrence_site':'제품 생산',
            '_origin_occurrence_site':'부품 생산',
        }
        rec,reason=clf.recommend_origin(d)
        self.assertEqual(rec,'부품')
        self.assertIn('화면 선택값',reason)

    def test_occurrence_site_product_production_supports_process(self):
        cases=[
            {'cause_4d':'원인 추가 검토 중','occurrence_site':'제품 생산'},
            {'cause_4d':'Root cause under additional review','occurrence_site':'Product production'},
            {'cause_4d':'설계 공차와 체결토크 산포가 동시에 영향','occurrence_site':'제품 생산'},
            {'cause_4d':'Design tolerance and assembly torque both contributed','occurrence_site':'Product manufacturing'},
        ]
        for d in cases:
            rec,reason=clf.recommend_origin(d)
            self.assertEqual(rec,'공정',msg=(rec,reason,d))
            self.assertIn('발생처',reason)

    def test_occurrence_site_part_production_supports_part(self):
        cases=[
            {'cause_4d':'원인 추가 검토 중','occurrence_site':'부품 생산'},
            {'cause_4d':'Root cause under additional review','occurrence_site':'Part production'},
            {'cause_4d':'부품 편차와 조립 공정이 함께 의심됨','occurrence_site':'부품 생산'},
            {'cause_4d':'Supplier part variation and assembly process both suspected','occurrence_site':'Component production'},
        ]
        for d in cases:
            rec,reason=clf.recommend_origin(d)
            self.assertEqual(rec,'부품',msg=(rec,reason,d))
            self.assertIn('발생처',reason)

    def test_occurrence_site_never_overrides_clear_explicit_cause(self):
        cases=[
            ('설계',{'cause_4d':'설계마진 부족으로 구조 간섭 발생','occurrence_site':'제품 생산'}),
            ('부품',{'cause_4d':'협력사 부품 LOT 편차가 발생원인','occurrence_site':'제품 생산'}),
            ('공정',{'cause_4d':'체결토크 산포가 발생원인','occurrence_site':'부품 생산'}),
            ('기타',{'cause_4d':'운송 중 외부충격이 발생원인','occurrence_site':'부품 생산'}),
        ]
        for expected,d in cases:
            rec,_=clf.recommend_origin(d)
            self.assertEqual(rec,expected,msg=(expected,rec,d))

    def test_occurrence_site_can_support_when_4d_is_empty(self):
        rec,reason=clf.recommend_origin({
            'cause_4d':'','leak_cause':'','system_cause':'','occurrence_site':'제품 생산'
        })
        self.assertEqual(rec,'공정')
        self.assertIn('보조',reason)

        rec,reason=clf.recommend_origin({
            'cause_4d':'','leak_cause':'','system_cause':'','occurrence_site':'부품 생산'
        })
        self.assertEqual(rec,'부품')
        self.assertIn('보조',reason)

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
