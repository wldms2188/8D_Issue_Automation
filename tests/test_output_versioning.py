import sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import final_output_polish as f
import output_variant_final as o

class OutputVersionTests(unittest.TestCase):
 def test_version_path_increments_existing_versions(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)
   (p/'weekly_v0.1.pptx').touch(); (p/'weekly_v0.2.pptx').touch()
   self.assertEqual(f._version_path(p/'weekly_업데이트.pptx').name,'weekly_v0.3.pptx')
 def test_latest_version_finds_highest_number(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td); out=p/'자동화_결과'; out.mkdir()
   src=p/'weekly.pptx'; src.touch()
   (out/'weekly_v0.2.pptx').touch(); (out/'weekly_v0.10.pptx').touch()
   self.assertEqual(o._latest_version(src).name,'weekly_v0.10.pptx')
if __name__=='__main__':unittest.main()
