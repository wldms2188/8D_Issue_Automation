import sys,tempfile,unittest
from pathlib import Path
from pptx import Presentation
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
 def test_update_only_ignores_unchanged_slides_that_shift_index(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td); src=p/'weekly.pptx'; dst=p/'weekly_v0.1.pptx'
   a=Presentation(); a.slides.add_slide(a.slide_layouts[6]).shapes.add_textbox(0,0,100,100).text='Summary'
   a.slides.add_slide(a.slide_layouts[6]).shapes.add_textbox(0,0,100,100).text='Detail A'
   a.slides.add_slide(a.slide_layouts[6]).shapes.add_textbox(0,0,100,100).text='Detail B'
   a.save(src)
   b=Presentation(src)
   b.slides[0].shapes[0].text='Summary updated'
   # Insert a new changed detail before old details, causing their indices to shift.
   new=b.slides.add_slide(b.slide_layouts[6]); new.shapes.add_textbox(0,0,100,100).text='New detail'
   sid=b.slides._sldIdLst[-1]; b.slides._sldIdLst.remove(sid); b.slides._sldIdLst.insert(1,sid)
   b.save(dst)
   reduced,n=o._ppt_update_only(src,dst)
   self.assertEqual(n,2)
   r=Presentation(reduced)
   texts=[' '.join(sh.text for sh in sl.shapes if hasattr(sh,'text')) for sl in r.slides]
   self.assertEqual(len(texts),2)
   self.assertTrue(any('Summary updated' in x for x in texts))
   self.assertTrue(any('New detail' in x for x in texts))

if __name__=='__main__':unittest.main()
