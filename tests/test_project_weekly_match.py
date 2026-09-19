import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import project_weekly_match_final as w

class ProjectWeeklyMatchTests(unittest.TestCase):
 def test_customer_prefix_is_ignored(self):
  self.assertEqual(w._project_key('MBAG_EB-L(EU)'),w._project_key('Other_EB-L(EU)'))
 def test_eu_us_remain_distinct(self):
  self.assertNotEqual(w._project_key('MBAG_EB-L(EU)'),w._project_key('MBAG_EB-L(US)'))
 def test_no_underscore_project_is_whole(self):
  self.assertEqual(w._project_key('Model Care 25'),'modelcare25')
 def test_underscore_inside_parentheses_is_not_customer_split(self):
  self.assertEqual(w._project_key('JF2S Delta (Set_Biz)'),'jf2sdeltasetbiz')
 def test_customer_project_is_not_prefixed_twice(self):
  d={'customer':'MBAG','task_name':'MBAG_EB-L(EU)'}
  self.assertEqual(w.s13._customer_task(d),'MBAG_EB-L(EU)')
  self.assertEqual(w.s13.v319._weekly_task(d),'MBAG_EB-L(EU)')

 def test_repeated_customer_prefix_is_collapsed(self):
  d={'customer':'MBAG','task_name':'MBAG_MBAG_EB-L(EU)'}
  self.assertEqual(w.s13._customer_task(d),'MBAG_EB-L(EU)')
  self.assertEqual(w.s13.v319._weekly_task(d),'MBAG_EB-L(EU)')

if __name__=='__main__':unittest.main()
