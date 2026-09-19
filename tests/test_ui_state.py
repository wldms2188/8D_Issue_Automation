import sys, unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))

import main_enterprise as ent
import main_enterprise_v2 as v2
import main_enterprise_v3 as v3

class UIStateTests(unittest.TestCase):
    def test_selected_8d_replaces_initial_prompt(self):
        s=ent.target_ready_status(r'C:\\work\\sample.pptx','READY  ·  8D 원본을 선택해 주세요.')
        self.assertIn('8D 원본 선택 완료',s)
        self.assertNotIn('선택해 주세요',s)

    def test_preview_complete_is_not_final_update_complete(self):
        pct,label=v2.progress_state('READY · 8D 추출 완료',15)
        self.assertEqual(pct,25)
        self.assertEqual(label,'8D 추출 완료')

    def test_final_update_complete_is_100(self):
        pct,label=v2.progress_state('COMPLETE · 선택한 자료의 업데이트가 완료되었습니다.',85)
        self.assertEqual((pct,label),(100,'업데이트 완료'))

    def test_initial_window_uses_more_vertical_space(self):
        w,h=v3.EnterpriseAppV3._initial_window_size(1920,1080)
        self.assertGreaterEqual(w,1180)
        self.assertGreaterEqual(h,900)
        self.assertLessEqual(h,1080)

if __name__=='__main__':
    unittest.main()
