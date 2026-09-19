import sys, unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))

import main_enterprise as ent
import main_recovery_step12 as step12
import main_enterprise_v2 as v2
import main_enterprise_v3 as v3

class UIStateTests(unittest.TestCase):
    def test_selected_8d_replaces_initial_prompt(self):
        s=ent.target_ready_status(r'C:\\work\\sample.pptx','READY  ·  8D 원본을 선택해 주세요.')
        self.assertIn('8D 원본 선택 완료',s)
        self.assertNotIn('선택해 주세요',s)

    def test_selected_8d_ready_message_is_explicit(self):
        s=ent.target_ready_status('sample.pptx','READY  ·  8D 원본을 선택해 주세요.')
        self.assertEqual(s,'READY  ·  8D 원본 선택 완료 · 미리보기 가능')

    def test_preview_complete_is_not_final_update_complete(self):
        pct,label=v2.progress_state('READY · 8D 추출 완료',15)
        self.assertEqual(pct,25)
        self.assertEqual(label,'8D 추출 완료')

    def test_final_update_complete_is_100(self):
        pct,label=v2.progress_state('COMPLETE · 선택한 자료의 업데이트가 완료되었습니다.',85)
        self.assertEqual((pct,label),(100,'업데이트 완료'))

    def test_initial_window_keeps_status_message_visible(self):
        w,h=v3.EnterpriseAppV3._initial_window_size(1920,1080)
        self.assertEqual((w,h),(1360,950))
        self.assertLess(h,1000)

    def test_weekly_status_follows_agreed_6d_priority(self):
        self.assertEqual(ent.weekly_recommended_status({'verification_6d':'검증 완료'}),'개선 완료')
        self.assertEqual(ent.weekly_recommended_status({'verification_6d':'완료 예정'}),'개선 검증중')
        self.assertEqual(ent.weekly_recommended_status({'verification_6d':'Validation in progress.'}),'개선 검증중')
        self.assertEqual(ent.weekly_recommended_status({'cause_4d':'원인 확인','action_5d':''}),'개선 검증중')
        self.assertEqual(ent.weekly_recommended_status({'cause_4d':'','action_5d':'','verification_6d':''}),'원인/개선 미확인')
        self.assertEqual(ent.weekly_status_from_choice({'verification_6d':'DV abnormal'},'close'),'개선 완료')

    def test_explicit_result_after_progress_label_has_priority(self):
        completed=[
            '진행 : 이상없음',
            '진행 중 : 이상 없음',
            '검증 진행 : 정상',
            '평가 진행 : 문제 없음',
            'Progress: No abnormalities',
            'Validation in progress: No additional abnormalities',
            'Verification ongoing: Result normal',
            'Test progress: All tests passed',
        ]
        for text in completed:
            state,_=ent.weekly_verification_state(text)
            self.assertEqual(state,'complete',msg=text)
            status,_=ent.issue_db_recommended_status({
                'action_5d':'Corrective action completed.',
                'verification_6d':text,
            })
            self.assertEqual(status,'close',msg=text)

    def test_explicit_result_abnormal_or_pending_still_wins(self):
        cases=[
            ('진행 중 : 이상 발생','abnormal'),
            ('검증 진행 : NG','abnormal'),
            ('Progress: Final result failed','abnormal'),
            ('진행 : 추가 검증 예정','pending'),
            ('Progress: Additional validation required','pending'),
        ]
        for text,expected in cases:
            state,_=ent.weekly_verification_state(text)
            self.assertEqual(state,expected,msg=text)

    def test_so_far_no_abnormality_remains_provisional(self):
        provisional=[
            '진행 중 : 현재까지 이상 없음',
            'Validation in progress: No abnormalities so far',
            'Verification ongoing: No abnormality to date',
        ]
        for text in provisional:
            state,_=ent.weekly_verification_state(text)
            self.assertEqual(state,'pending',msg=text)

    def test_no_additional_abnomalities_is_complete(self):
        text='No additional abnomalities were observed after the verification.'
        state,_=ent.weekly_verification_state(text)
        self.assertEqual(state,'complete')
        self.assertEqual(ent.weekly_recommended_status({'verification_6d':text}),'개선 완료')

    def test_customer_project_mismatch_uses_canonical_8d_value(self):
        d={'customer':'MBAG','task_name':'EB565M'}
        mismatch,extracted,selected=ent.customer_project_mismatch(d,'MBAG_EB565M')
        self.assertFalse(mismatch)
        self.assertEqual(extracted,'MBAG_EB565M')
        mismatch,_,_=ent.customer_project_mismatch(d,'GM_A1')
        self.assertTrue(mismatch)

    def test_weekly_english_completion_variants(self):
        completed=[
            'Verification completed.',
            'Validation is complete.',
            'All tests passed.',
            'No abnormality observed after verification.',
            'No abnomality detected.',
            'No abnormalities found.',
            'No recurrence observed.',
            'Result is within specification.',
            'Effectiveness confirmed.',
        ]
        for text in completed:
            state,_=ent.weekly_verification_state(text)
            self.assertEqual(state,'complete',msg=text)
            self.assertEqual(ent.weekly_recommended_status({'verification_6d':text}),'개선 완료',msg=text)

    def test_issue_db_comprehensive_completed_5d_6d_matrix(self):
        action_done=[
            'Corrective action completed.',
            'Countermeasure implemented.',
            'Change successfully implemented.',
            'Design revision completed.',
            'Process parameter updated.',
            'New control applied.',
            'Improvement action released.',
            '개선 완료 및 적용 완료',
        ]
        verification_done=[
            'No additional abnormalities were observed.',
            'No additional abnomalities were observed.',
            'No further abnormalities detected.',
            'No abnormality occurred after implementation.',
            'No abnormalities found during validation.',
            'No recurrence observed.',
            'No repeat issue was found.',
            'Verification completed and passed.',
            'Validation completed successfully.',
            'Effectiveness confirmed.',
            'All acceptance criteria met.',
            'Result is within specification.',
            'Evaluation completed with normal result.',
            'All tests passed.',
        ]
        total=0
        for a in action_done:
            for v in verification_done:
                status,_=ent.issue_db_recommended_status({'action_5d':a,'verification_6d':v})
                self.assertEqual(status,'close',msg=(a,v))
                total+=1
        self.assertEqual(total,112)

    def test_issue_db_comprehensive_pending_matrix_stays_open(self):
        action_done=[
            'Corrective action completed.',
            'Countermeasure implemented.',
            'Design revision completed.',
            'Process parameter updated.',
        ]
        verification_pending=[
            'Verification in progress.',
            'Validation pending.',
            'Verification is scheduled next week.',
            'To be verified after DV.',
            'Not yet validated.',
            'Additional verification required.',
            'Remaining validation is in progress.',
            'Awaiting verification result.',
            'Verification to follow.',
            'Under validation.',
        ]
        total=0
        for a in action_done:
            for v in verification_pending:
                status,_=ent.issue_db_recommended_status({'action_5d':a,'verification_6d':v})
                self.assertEqual(status,'open',msg=(a,v))
                total+=1
        self.assertEqual(total,40)

    def test_issue_db_comprehensive_abnormal_matrix_stays_open(self):
        abnormal=[
            'Final validation failed.',
            'Abnormality observed during DV.',
            'Defect detected after implementation.',
            'Issue recurred after the action.',
            'Result was NG.',
            'Result is out of spec.',
            'Acceptance criteria not met.',
            'Requirement not met.',
        ]
        for v in abnormal:
            status,_=ent.issue_db_recommended_status({
                'action_5d':'Corrective action completed.',
                'verification_6d':v
            })
            self.assertEqual(status,'open',msg=v)

    def test_no_additional_abnormalities_is_not_misread_as_abnormal(self):
        variants=[
            'No additional abnormalities.',
            'No additional abnormalities were observed.',
            'No additional abnomalities were observed.',
            'No further abnormalities were detected.',
            'No abnormality occurred.',
            'No abnormalities found.',
        ]
        for v in variants:
            state,_=ent.weekly_verification_state(v)
            self.assertEqual(state,'complete',msg=v)
            status,_=ent.issue_db_recommended_status({
                'action_5d':'Corrective action completed.',
                'verification_6d':v
            })
            self.assertEqual(status,'close',msg=v)

    def test_weekly_english_pending_variants(self):
        pending=[
            'Verification in progress.',
            'Validation pending.',
            'Verification is scheduled.',
            'To be completed after DV.',
            'Not completed yet.',
            'Under validation.',
            'Monitoring ongoing.',
            'Awaiting verification result.',
            'Expected to be completed next week.',
        ]
        for text in pending:
            state,_=ent.weekly_verification_state(text)
            self.assertEqual(state,'pending',msg=text)
            self.assertEqual(ent.weekly_recommended_status({'verification_6d':text}),'개선 검증중',msg=text)

    def test_weekly_failed_result_beats_completed_word(self):
        text='Verification completed, but final result failed / abnormal.'
        state,_=ent.weekly_verification_state(text)
        self.assertEqual(state,'abnormal')
        self.assertEqual(ent.weekly_recommended_status({'verification_6d':text}),'개선 검증중')

    def test_issue_db_status_popup_when_5d_or_6d_exists(self):
        self.assertTrue(ent.issue_db_status_confirmation_required({'action_5d':'대책 있음','verification_6d':''}))
        self.assertTrue(ent.issue_db_status_confirmation_required({'action_5d':'','verification_6d':'검증 내용'}))
        self.assertTrue(ent.issue_db_status_confirmation_required({'action_5d':'대책','verification_6d':'검증'}))
        self.assertFalse(ent.issue_db_status_confirmation_required({'action_5d':'','verification_6d':''}))

    def test_weekly_popup_only_for_final_open_issue_db(self):
        self.assertTrue(ent.weekly_status_confirmation_required('open'))
        self.assertFalse(ent.weekly_status_confirmation_required('close'))
        self.assertFalse(ent.weekly_status_confirmation_required(None))

    def test_weekly_selected_status_is_independent_from_issue_db(self):
        d={'action_5d':'fix','verification_6d':'DV abnormal'}
        g={'_issue_status_selected':'open','_weekly_status_selected':'개선 완료'}
        self.assertEqual(step12._signal_status_from_g(d,g),'개선 완료')

    def test_status_reasons_are_separate(self):
        d={'action_5d':'Guide revised','verification_6d':'Validation is in progress.'}
        db_reason=ent.issue_db_status_reason(d,'open')
        weekly_reason=ent.weekly_status_reason(d,'개선 검증중')
        self.assertIn('진행 중',db_reason)
        self.assertIn('진행 중',weekly_reason)

if __name__=='__main__':
    unittest.main()
