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
