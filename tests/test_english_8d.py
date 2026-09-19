import sys,unittest,tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'app'))
import english_8d_final as e

class English8DTests(unittest.TestCase):
 def test_common_fields_translate_without_changing_identifiers(self):
  d={'problem':'Crack occurred at MBAG_EB-L(EU) Sample DUT3, 2.5 mm.',
     'cause_4d':'Insufficient design margin and tolerance interference confirmed.',
     'action_5d':'Corrective action: drawing change for MBAG_EB-L(EU).'}
  x=e.enhance_dict(d)
  self.assertIn('MBAG_EB-L(EU)',x['problem']); self.assertIn('DUT3',x['problem']); self.assertIn('2.5 mm',x['problem'])
  self.assertIn('크랙',x['problem']); self.assertIn('설계마진',x['cause_4d']); self.assertIn('공차',x['cause_4d'])
 def test_korean_is_not_rewritten(self):
  s='체결토크 산포로 조립불량 발생'; self.assertEqual(e.enhance_dict({'cause_4d':s})['cause_4d'],s)
 def test_mixed_preserves_model(self):
  s='Root cause: torque 산포 at BDI_BOLT'; x=e.enhance_dict({'cause_4d':s}); self.assertIn('BDI_BOLT',x['cause_4d'])
 def test_incomplete_translation_is_flagged(self):
  x=e.enhance_dict({'problem':'Unexpected electrical behavior remained after repeated vehicle evaluation.'})
  self.assertTrue(x['_english_translation_incomplete'])
  self.assertEqual(e.apply_english_choice(x,True)['problem'],'Unexpected electrical behavior remained after repeated vehicle evaluation.')
 def test_technical_identifiers_do_not_trigger_failure(self):
  x=e.enhance_dict({'problem':'Crack occurred at MBAG_EB-L(EU) DUT3.'})
  self.assertFalse(x['_english_translation_incomplete'])
 def test_numbered_sections_keep_all_content_and_stop_at_next_d(self):
  blocks=[
   '2D Problem Description\nCrack found at weld\nLeak observed',
   '3D',
   'Containment\nStop shipment\n100% inspection',
   '4D Root Cause\nDesign margin insufficient\nTolerance stack-up',
   '5D Corrective Action\nDrawing changed\nTolerance revised',
   '6D Verification\nDV test passed\nNo recurrence',
   '7D Prevent Recurrence\nUpdate lesson learned',
   '8D Closure\nClosed']
  x=e.extract_sections_from_blocks(blocks)
  self.assertIn('Crack found at weld',x['problem'])
  self.assertIn('Leak observed',x['problem'])
  self.assertIn('Stop shipment',x['temporary_action'])
  self.assertIn('100% inspection',x['temporary_action'])
  self.assertIn('Design margin insufficient',x['cause_4d'])
  self.assertIn('Tolerance stack-up',x['cause_4d'])
  self.assertIn('Drawing changed',x['action_5d'])
  self.assertIn('DV test passed',x['verification_6d'])
  self.assertNotIn('7D',x['verification_6d'])
  self.assertNotIn('Closure',x['verification_6d'])

 def test_unsplit_4d_defaults_to_occurrence_cause(self):
  x=e.extract_sections_from_blocks(['4D','Root cause item A','Root cause item B','5D','Action'])
  self.assertEqual(x['cause_4d'],'Root cause item A\nRoot cause item B')
  self.assertEqual(x['leak_cause'],'')

 def test_explicit_escape_cause_splits_4d(self):
  x=e.extract_sections_from_blocks(['4D Root Cause','Occurrence A','Escape Cause','Detection control missing','5D Corrective Action','Fix'])
  self.assertIn('Occurrence A',x['cause_4d'])
  self.assertIn('Detection control missing',x['leak_cause'])

 def test_english_pending_status_is_open(self):
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'Validation is in progress.'})[0],'open')
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'Verification completed and passed.'})[0],'close')

 def test_supplied_english_layout_boundaries_and_photo_captions(self):
  blocks=[
   '1D','Team build','Model EB565M (E122A)','2D','Problem description','Defect Phenomenon','Scratch on metal strap',
   '[Close-up Photo]','3D','Containment','Stop shipment','4D','Root cause','Guide interference confirmed','Non-defect : N/A',
   '5D','Corrective action','Guide position improved','[Unloading Process]','[After Improvement]',
   '6D','Validation','DV validation passed','7D','Preventive action','Standard update','8D','Follow up','Improvement : N/A']
  x=e.extract_sections_from_blocks(blocks)
  self.assertIn('Scratch on metal strap',x['problem'])
  self.assertNotIn('Team build',x['problem'])
  self.assertNotIn('[Close-up Photo]',x['problem'])
  self.assertIn('Stop shipment',x['temporary_action'])
  self.assertIn('Guide interference confirmed',x['cause_4d'])
  self.assertIn('Guide position improved',x['action_5d'])
  self.assertNotIn('[Unloading Process]',x['action_5d'])
  self.assertNotIn('[After Improvement]',x['action_5d'])
  self.assertIn('DV validation passed',x['verification_6d'])
  self.assertNotIn('Preventive action',x['verification_6d'])
  self.assertNotIn('Follow up',x['verification_6d'])

 def test_varied_english_section_names(self):
  variants=[
   (['Problem Definition','Scratch found','Immediate Action','Stop shipment','Cause Investigation','Guide interference','Countermeasure','Guide revised','Effectiveness Check','Passed DV'],
    {'problem':'Scratch found','temporary_action':'Stop shipment','cause_4d':'Guide interference','action_5d':'Guide revised','verification_6d':'Passed DV'}),
   (['Failure Phenomenon','Dent found','Customer Protection','100% sorting','5 Why','Fixture gap','Permanent Action','Fixture changed','Validation Results','No recurrence'],
    {'problem':'Dent found','temporary_action':'100% sorting','cause_4d':'Fixture gap','action_5d':'Fixture changed','verification_6d':'No recurrence'}),
   (['Issue Description','Leak found','Short-Term Action','Quarantine lot','Failure Cause','Seal damage','Corrective Measure','Seal changed','Effect Confirmation','Test passed'],
    {'problem':'Leak found','temporary_action':'Quarantine lot','cause_4d':'Seal damage','action_5d':'Seal changed','verification_6d':'Test passed'}),
  ]
  for blocks,expected in variants:
   x=e.extract_sections_from_blocks(blocks)
   for key,val in expected.items():self.assertIn(val,x[key],msg=(blocks,key,x))

 def test_semantic_7d_8d_boundaries_stop_6d_without_numbers(self):
  x=e.extract_sections_from_blocks([
   '6D','Validation','DV passed',
   'Customer Response','Horizontal deployment done',
   'Request Items','Customer request A'])
  self.assertIn('DV passed',x['verification_6d'])
  self.assertNotIn('Horizontal deployment done',x['verification_6d'])
  self.assertNotIn('Customer request A',x['verification_6d'])

 def test_authoritative_spatial_fields_clear_legacy_bleed(self):
  d={'problem':'Member Alice\nQuality Bob','verification_6d':'Customer Response\nRequest Items'}
  x=e.enhance_dict(d,section_blocks=['2D','Scratch in strap','6D','DV passed'],
                   authoritative_keys={'problem','verification_6d'})
  self.assertEqual(x['problem'],'Scratch in strap')
  self.assertEqual(x['verification_6d'],'DV passed')

 def test_spatial_cell_mapping_prevents_1d_and_7d_8d_bleed(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'layout.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   marker_y=[0.5,2.0,3.5,5.0]
   for idx,y in enumerate(marker_y,1):
    sh=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.4),Inches(0.3)); sh.text=f'{idx}D'
   for idx,y in enumerate(marker_y,5):
    sh=sl.shapes.add_textbox(Inches(5.1),Inches(y),Inches(0.4),Inches(0.3)); sh.text=f'{idx}D'
   lt=sl.shapes.add_table(4,2,Inches(0.7),Inches(0.5),Inches(4.0),Inches(6.0)).table
   rt=sl.shapes.add_table(4,2,Inches(5.7),Inches(0.5),Inches(4.0),Inches(6.0)).table
   left=[('Member','Alice'),('Defect Phenomenon','Scratch in strap'),('Containment','Stop shipment'),('Root cause','Guide interference')]
   right=[('Corrective action','Guide revised'),('Validation','DV passed'),('Customer Response','Horizontal deployment done'),('Request Items','Customer request A')]
   for r,(a,b) in enumerate(left):
    lt.rows[r].height=Inches(1.5); lt.cell(r,0).text=a; lt.cell(r,1).text=b
   for r,(a,b) in enumerate(right):
    rt.rows[r].height=Inches(1.5); rt.cell(r,0).text=a; rt.cell(r,1).text=b
   prs.save(p)
   blocks=e._spatial_section_blocks(p)
   x=e.extract_sections_from_blocks(blocks)
   self.assertIn('Scratch in strap',x['problem'])
   self.assertNotIn('Alice',x['problem'])
   self.assertNotIn('Member',x['problem'])
   self.assertIn('DV passed',x['verification_6d'])
   self.assertNotIn('Horizontal deployment done',x['verification_6d'])
   self.assertNotIn('Customer request A',x['verification_6d'])

 def test_extract_recovers_2d_when_marker_is_below_real_heading(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'marker_below_heading.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   # Real 2D heading/content are above the visual 2D marker.
   sh=sl.shapes.add_textbox(Inches(0.9),Inches(0.15),Inches(4.5),Inches(0.7))
   sh.text='Problem Description\nScratch on metal strap'
   for label,y in [('2D',0.8),('3D',2.2),('4D',3.6),('5D',5.0)]:
    m=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.5),Inches(0.3)); m.text=label
   sh=sl.shapes.add_textbox(Inches(0.9),Inches(1.0),Inches(4.5),Inches(0.7)); sh.text='Containment\nStop shipment'
   sh=sl.shapes.add_textbox(Inches(0.9),Inches(2.4),Inches(4.5),Inches(0.7)); sh.text='Root Cause\nGuide interference'
   prs.save(p)
   x=e.extract(p)
   self.assertIn('Scratch on metal strap',x['problem'])
   self.assertNotIn('Stop shipment',x['problem'])
   self.assertIn('Stop shipment',x['temporary_action'])
   self.assertIn('Guide interference',x['cause_4d'])

 def test_spatial_semantic_heading_corrects_one_row_shift(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'shifted_layout.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   # D markers are intentionally one visual row ahead of the semantic content.
   for label,y in [('2D',0.5),('3D',2.0),('4D',3.5),('5D',5.0)]:
    sh=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.5),Inches(0.3)); sh.text=label
   contents=[
    ('Containment\nStop shipment',0.7),
    ('Root Cause\nGuide interference confirmed',2.2),
    ('Corrective Action\nGuide revised',3.7),
   ]
   for text,y in contents:
    sh=sl.shapes.add_textbox(Inches(0.9),Inches(y),Inches(4.5),Inches(0.8)); sh.text=text
   prs.save(p)
   blocks=e._spatial_section_blocks(p)
   x=e.extract_sections_from_blocks(blocks)
   self.assertNotIn('Stop shipment',x['problem'])
   self.assertIn('Stop shipment',x['temporary_action'])
   self.assertIn('Guide interference confirmed',x['cause_4d'])
   self.assertIn('Guide revised',x['action_5d'])

 def test_translation_coverage_reports_partial_dictionary_conversion(self):
  x=e.enhance_dict({'problem':'Crack occurred after repeated vehicle evaluation.'})
  pct=e.translation_coverage(x)
  self.assertGreater(pct,0)
  self.assertLess(pct,100)

 def test_5d_heading_and_first_row_are_not_duplicated(self):
  x=e.extract_sections_from_blocks([
   '5D','Corrective Action','Guide revised','Corrective Action','Guide revised',
   '6D','Validation'])
  self.assertEqual(x['action_5d'],'Guide revised')

 def test_5d_inline_heading_keeps_payload_once(self):
  x=e.extract_sections_from_blocks([
   '5D','Corrective Action: Guide revised','Guide revised','6D','Validation'])
  self.assertEqual(x['action_5d'],'Guide revised')

 def test_occurrence_date_accepts_korean_short_label_and_range(self):
  self.assertEqual(e._occurrence_date_from_blocks(['발생',"'26.8/27~28"]),"'26.8/27~28")
  self.assertEqual(e.parse_year_month_extended("'26.8/27~28"),('2026','8'))

 def test_occurrence_date_accepts_english_and_common_misspelling(self):
  self.assertEqual(e._occurrence_date_from_blocks(['Occurrence',"'26.8/27~28"]),"'26.8/27~28")
  self.assertEqual(e._occurrence_date_from_blocks(['Occurence: 2026.09.03']),'2026.09.03')

 def test_enhance_dict_updates_occurrence_date_from_full_metadata(self):
  x=e.enhance_dict({'occurrence_date':''},raw_text="Customer\nABC\n발생\n'26.8/27~28")
  self.assertEqual(x['occurrence_date'],"'26.8/27~28")

 def test_abnormal_and_abnomal_6d_are_open(self):
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'DV result abnormal.'})[0],'open')
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'Validation result abnomal.'})[0],'open')
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'DV result NG / failed.'})[0],'open')

 def test_normal_passed_6d_is_close(self):
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'DV result normal and passed.'})[0],'close')
  self.assertEqual(e.judge_status_bilingual({'verification_6d':'No abnormality after DV.'})[0],'close')

 def test_issue_db_year_month_uses_extended_date(self):
  from openpyxl import Workbook
  ws=Workbook().active
  g={'plm_no':'P1','form_factor':'파우치형','product_type':'EV Pack','team':'T','owner':'O','sample':'S','stage':'DV','occurrence_site':'etc.'}
  d={'occurrence_date':"'26.8/27~28",'problem':'p','cause_4d':'c','action_5d':'a','verification_6d':'Validation in progress.'}
  e.v310.base.write_row(ws,1,d,g)
  self.assertEqual(str(ws.cell(1,6).value),'2026')
  self.assertIn(str(ws.cell(1,7).value),('8','8월'))

if __name__=='__main__':unittest.main()
