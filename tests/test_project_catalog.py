import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import project_autocomplete_final as p

class ProjectCatalogTests(unittest.TestCase):
 def test_customer_project_split_uses_first_underscore(self):
  self.assertEqual(p.split_customer_task('MBAG_EB565M'),('MBAG','EB565M'))
  self.assertEqual(p.split_customer_task('Renault_EV2020 CTP 800V'),('Renault','EV2020 CTP 800V'))
  self.assertEqual(p.split_customer_task('JF2S Delta (Set_Biz)'),('', 'JF2S Delta (Set_Biz)'))
 def test_project_matching_ignores_customer_prefix(self):
  self.assertEqual(p.project_key('MBAG_EB-L(EU)',True),p.project_key('Other_EB-L(US)',True))
 def test_weekly_identity_preserves_parenthesis_variant(self):
  self.assertNotEqual(p.project_key('MBAG_EB-L(EU)',False),p.project_key('MBAG_EB-L(US)',False))
  self.assertEqual(p.project_key('MBAG_EB-L(EU)',False),p.project_key('Other_EB-L(EU)',False))
 def test_no_underscore_uses_whole_project(self):
  self.assertEqual(p.split_customer_task('Model Care 25'),('', 'Model Care 25'))
  self.assertEqual(p.project_key('Model Care 25',False),p.project_key('Model Care 25',False))
 def test_parenthesis_underscore_stays_in_project(self):
  self.assertEqual(p.split_customer_task('JF2S Delta (Set_Biz)'),('', 'JF2S Delta (Set_Biz)'))
 def test_parenthesis_variants_are_suggested(self):
  xs=p.canonical_candidates('EB-L','원통형Pack개발품질팀')
  self.assertIn('MBAG_EB-L(EU)',xs); self.assertIn('MBAG_EB-L(US)',xs)
 def test_filename_exact_project_match_uses_existing_catalog(self):
  xs=p.filename_project_candidates(r'C:\\tmp\\8D_Report_MBAG_EB565M_Scratch.pptx')
  self.assertEqual(xs,['MBAG_EB565M'])

 def test_filename_shared_model_word_surfaces_multiple_projects(self):
  xs=p.filename_project_candidates(r'C:\\tmp\\8D_Report_EB-L_issue.pptx','원통형Pack개발품질팀')
  self.assertIn('MBAG_EB-L(EU)',xs)
  self.assertIn('MBAG_EB-L(US)',xs)

 def test_filename_shared_vda_word_surfaces_relevant_variants(self):
  xs=p.filename_project_candidates(r'C:\\tmp\\8D_STLA_VDA590_issue.pptx','파우치형Pack개발품질2팀')
  self.assertIn('STLA_VDA590_2P8S',xs)
  self.assertIn('STLA_VDA590_1P16S',xs)
  self.assertNotIn('STLA_VDA355',xs)

 def test_team_filter(self):
  xs=p.canonical_candidates('V710','파우치형Pack개발품질1팀')
  self.assertEqual(xs,['Ford_V710'])
if __name__=='__main__':unittest.main()
