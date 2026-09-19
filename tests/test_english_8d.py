import sys,unittest,tempfile,io
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
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

 def test_english_images_follow_strict_d_regions(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'english_images_by_d.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])

   # Two-column D layout similar to the real 8D structure.
   for label,x,y in [
    ('2D',0.2,1.0),('3D',0.2,3.0),('4D',0.2,5.0),
    ('4D',6.0,1.0),('5D',6.0,3.0),('6D',6.0,5.0),
   ]:
    sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(.55),Inches(.3)); sh.text=label

   # Clarify the right-side 4D ownership.
   h=sl.shapes.add_textbox(Inches(6.8),Inches(1.0),Inches(1.5),Inches(.3)); h.text='Escape Cause'

   def add_pic(x,y,seed):
    im=Image.new('RGB',(180,100),(20*seed%255,40*seed%255,60*seed%255))
    b=io.BytesIO(); im.save(b,'PNG'); b.seek(0)
    sl.shapes.add_picture(b,Inches(x),Inches(y),Inches(1.2),Inches(.7))

   add_pic(1.2,1.25,1)  # 2D
   add_pic(1.2,3.25,2)  # 3D
   add_pic(1.2,5.15,3)  # 4D cause
   add_pic(7.2,1.25,4)  # 4D escape
   add_pic(7.2,3.25,5)  # 5D
   add_pic(7.2,5.15,6)  # 6D
   prs.save(p)

   imgs=e._english_section_images(p)
   for key in ('2D','3D','4D_CAUSE','4D_LEAK','5D','6D'):
    self.assertEqual(len(imgs[key]),1,msg=(key,{k:len(v) for k,v in imgs.items()}))

   # The 2D picture must never leak into 3D and vice versa.
   self.assertLess(imgs['2D'][0][1][1],imgs['3D'][0][1][1])
   self.assertLess(imgs['4D_LEAK'][0][1][1],imgs['5D'][0][1][1])
   self.assertLess(imgs['5D'][0][1][1],imgs['6D'][0][1][1])

 def test_unsplit_4d_defaults_to_occurrence_cause(self):
  x=e.extract_sections_from_blocks(['4D','Root cause item A','Root cause item B','5D','Action'])
  self.assertEqual(x['cause_4d'],'Root cause item A\nRoot cause item B')
  self.assertEqual(x['leak_cause'],'')

 def test_explicit_escape_cause_splits_4d(self):
  x=e.extract_sections_from_blocks(['4D Root Cause','Occurrence A','Escape Cause','Detection control missing','5D Corrective Action','Fix'])
  self.assertIn('Occurrence A',x['cause_4d'])
  self.assertIn('Detection control missing',x['leak_cause'])

 def test_english_nonpending_6d_still_proposes_close_confirmation(self):
  decision,text=e.judge_status_bilingual({'action_5d':'Guide revised','verification_6d':'1000 cycle evaluation result reviewed.'})
  self.assertEqual(decision,'close')
  self.assertTrue(text)

 def test_english_completed_action_plus_no_additional_abnormalities_is_close(self):
  cases=[
   ('Corrective action completed.','No additional abnormalities were observed.'),
   ('Countermeasure implemented.','No additional abnomalities were observed.'),
   ('Design revision completed.','No further abnormalities detected.'),
   ('Process update completed.','No recurrence observed.'),
  ]
  for action,verify in cases:
   decision,_=e.judge_status_bilingual({'action_5d':action,'verification_6d':verify})
   self.assertEqual(decision,'close',msg=(action,verify))

 def test_english_no_abnomality_and_completed_are_close(self):
  for text in [
   'No abnomality observed.',
   'No abnormalities detected.',
   'Verification completed.',
   'Validation complete.',
   'All tests passed.',
   'No recurrence observed.',
  ]:
   self.assertEqual(e.judge_status_bilingual({'verification_6d':text})[0],'close',msg=text)

 def test_english_planned_ongoing_and_failed_are_open(self):
  for text in [
   'Verification scheduled next week.',
   'Validation pending.',
   'To be completed.',
   'Under verification.',
   'Monitoring ongoing.',
   'Verification completed but result failed.',
  ]:
   self.assertEqual(e.judge_status_bilingual({'verification_6d':text})[0],'open',msg=text)

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
   x,_=e._english_fields_from_spatial(blocks)
   self.assertIn('Scratch in strap',x['problem'])
   self.assertNotIn('Alice',x['problem'])
   self.assertNotIn('Member',x['problem'])
   self.assertIn('DV passed',x['verification_6d'])
   self.assertNotIn('Horizontal deployment done',x['verification_6d'])
   self.assertNotIn('Customer request A',x['verification_6d'])

 def test_semantic_table_headers_are_not_rejected_by_geometry(self):
  # The table/header meaning is authoritative. A geometry disagreement must not
  # cause content to be dropped or moved to another D section.
  semantic,detected=e._semantic_fields_and_detected([
   'Problem Description','Scratch on metal strap',
   'Containment','Stop shipment',
   'Root Cause','Guide interference',
   'Corrective Action','Guide revised',
   'Verification','DV passed'])
  geo,_=e._english_fields_from_spatial([
   '2D','Stop shipment','3D','Guide interference','4D','Guide revised','5D','DV passed'])
  self.assertIn('Scratch on metal strap',semantic['problem'])
  self.assertIn('Stop shipment',semantic['temporary_action'])
  self.assertIn('Guide interference',semantic['cause_4d'])
  self.assertIn('Guide revised',semantic['action_5d'])
  self.assertIn('DV passed',semantic['verification_6d'])
  self.assertNotEqual(semantic['problem'],geo.get('problem',''))

 def test_extract_prefers_table_headers_over_d_marker_geometry(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'semantic_first.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   # Deliberately misleading D marker positions.
   for label,y in [('2D',0.5),('3D',1.8),('4D',3.1),('5D',4.4),('6D',5.7)]:
    m=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.5),Inches(0.3)); m.text=label
   tb=sl.shapes.add_table(5,2,Inches(0.9),Inches(0.5),Inches(7.5),Inches(6.2)).table
   rows=[
    ('Problem Description','Scratch on metal strap'),
    ('Containment','Stop shipment'),
    ('Root Cause','Guide interference'),
    ('Corrective Action','Guide revised'),
    ('Verification','DV passed'),
   ]
   for r,(a,b) in enumerate(rows):
    tb.cell(r,0).text=a; tb.cell(r,1).text=b
   prs.save(p)
   x=e.extract(p)
   self.assertTrue(x['_english_mode'])
   self.assertIn('Scratch on metal strap',x.get('problem_en_original') or x['problem'])
   self.assertIn('Stop shipment',x.get('temporary_action_en_original') or x['temporary_action'])
   self.assertIn('Guide interference',x.get('cause_4d_en_original') or x['cause_4d'])
   self.assertIn('Guide revised',x.get('action_5d_en_original') or x['action_5d'])
   self.assertIn('DV passed',x.get('verification_6d_en_original') or x['verification_6d'])

 def test_4d_subheadings_split_only_within_4d(self):
  x,detected=e._english_fields_from_spatial([
   '4D','Root Cause','Guide interference',
   'Escape Point','Inspection gap',
   'Systemic Root Cause','Control plan missing',
   '5D','Root Cause','Guide revised'])
  self.assertIn('Guide interference',x['cause_4d'])
  self.assertIn('Inspection gap',x['leak_cause'])
  self.assertIn('Control plan missing',x['system_cause'])
  # "Root Cause" inside the geometrical 5D area stays in 5D; it cannot move back to 4D.
  self.assertIn('Root Cause',x['action_5d'])
  self.assertIn('Guide revised',x['action_5d'])

 def test_language_gate_uses_80_percent_english_threshold(self):
  self.assertTrue(e._document_looks_english('A'*80+'가'*20))
  self.assertFalse(e._document_looks_english('A'*79+'가'*21))
  self.assertAlmostEqual(e.english_content_ratio('A'*80+'가'*20),80.0)

 def test_midpoint_regions_prevent_3d_and_4d_from_shifting_up_one_section(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'midpoint_regions.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])

   # Markers are reference centres. The actual content is intentionally above
   # each marker; old marker-to-next-marker bands shifted 3D into 2D and 4D into 3D.
   for label,y in [('2D',1.0),('3D',2.8),('4D',4.6),('5D',6.2)]:
    m=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.5),Inches(0.3)); m.text=label

   p2=sl.shapes.add_textbox(Inches(1.0),Inches(1.15),Inches(5.0),Inches(0.5)); p2.text='Problem Description\nScratch found'
   p3=sl.shapes.add_textbox(Inches(1.0),Inches(2.45),Inches(5.0),Inches(0.5)); p3.text='Containment\nStop shipment'
   p4=sl.shapes.add_textbox(Inches(1.0),Inches(4.25),Inches(5.0),Inches(0.5)); p4.text='Root Cause\nGuide interference'
   prs.save(p)

   blocks=e._spatial_section_blocks(p)
   fields,_=e._english_fields_from_spatial(blocks)
   self.assertIn('Scratch found',fields['problem'])
   self.assertIn('Stop shipment',fields['temporary_action'])
   self.assertNotIn('Stop shipment',fields['problem'])
   self.assertIn('Guide interference',fields['cause_4d'])
   self.assertNotIn('Guide interference',fields['temporary_action'])

 def test_scoped_table_semantics_match_korean_style_row_extraction(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'english_table_rows.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   for label,y in [('2D',0.8),('3D',2.2),('4D',3.6),('5D',5.0),('6D',6.4)]:
    m=sl.shapes.add_textbox(Inches(0.1),Inches(y),Inches(0.5),Inches(0.3)); m.text=label

   # Each table row owns its content, like the Korean exact-label extractor.
   rows=[
    ('Problem Description','Scratch found',0.9),
    ('Containment','Stop shipment',2.3),
    ('Root Cause','Guide interference',3.7),
    ('Corrective Action','Guide revised',5.1),
    ('Verification','DV passed',6.5),
   ]
   for a,b,y in rows:
    tb=sl.shapes.add_table(1,2,Inches(0.9),Inches(y),Inches(7.0),Inches(0.55)).table
    tb.cell(0,0).text=a; tb.cell(0,1).text=b

   # Same wording, but outside the real 4D band: must be ignored.
   bad=sl.shapes.add_table(1,2,Inches(0.9),Inches(0.1),Inches(7.0),Inches(0.45)).table
   bad.cell(0,0).text='Root Cause'; bad.cell(0,1).text='WRONG TABLE'
   prs.save(p)

   fields,detected=e._scoped_table_semantic_fields(p)
   self.assertEqual(fields['problem'],'Scratch found')
   self.assertEqual(fields['temporary_action'],'Stop shipment')
   self.assertEqual(fields['cause_4d'],'Guide interference')
   self.assertEqual(fields['action_5d'],'Guide revised')
   self.assertEqual(fields['verification_6d'],'DV passed')
   self.assertNotIn('WRONG TABLE',fields['cause_4d'])

 def test_unrelated_table_outside_4d_region_is_not_extracted(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'scoped_4d.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])

   # Unrelated table above the real 4D area. It deliberately contains the same
   # semantic label and must never win just because the wording matches.
   bad=sl.shapes.add_table(1,2,Inches(1.0),Inches(0.4),Inches(6.5),Inches(0.7)).table
   bad.cell(0,0).text='Root Cause'
   bad.cell(0,1).text='WRONG OTHER TABLE CAUSE'

   # 4D/5D markers define the real allowed 4D band.
   m4=sl.shapes.add_textbox(Inches(0.1),Inches(2.0),Inches(0.5),Inches(0.3)); m4.text='4D'
   m5=sl.shapes.add_textbox(Inches(0.1),Inches(5.0),Inches(0.5),Inches(0.3)); m5.text='5D'

   good=sl.shapes.add_table(1,2,Inches(1.0),Inches(2.4),Inches(6.5),Inches(1.0)).table
   good.cell(0,0).text='Root Cause'
   good.cell(0,1).text='CORRECT 4D GUIDE INTERFERENCE'
   prs.save(p)

   x=e.extract(p)
   original=x.get('cause_4d_en_original') or x.get('cause_4d','')
   self.assertIn('CORRECT 4D GUIDE INTERFERENCE',original)
   self.assertNotIn('WRONG OTHER TABLE CAUSE',original)

 def test_single_4d_band_splits_side_by_side_root_and_escape_columns(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'four_columns.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   m=sl.shapes.add_textbox(Inches(0.1),Inches(0.5),Inches(0.4),Inches(0.3)); m.text='4D'
   n=sl.shapes.add_textbox(Inches(0.1),Inches(3.2),Inches(0.4),Inches(0.3)); n.text='5D'
   a=sl.shapes.add_textbox(Inches(0.9),Inches(0.7),Inches(2.8),Inches(0.4)); a.text='Root Cause'
   b=sl.shapes.add_textbox(Inches(0.9),Inches(1.3),Inches(2.8),Inches(0.6)); b.text='Guide interference'
   cc=sl.shapes.add_textbox(Inches(5.0),Inches(0.7),Inches(2.8),Inches(0.4)); cc.text='Escape Point'
   dd=sl.shapes.add_textbox(Inches(5.0),Inches(1.3),Inches(2.8),Inches(0.6)); dd.text='Inspection gap'
   prs.save(p)
   blocks=e._spatial_section_blocks(p)
   x,_=e._english_fields_from_spatial(blocks)
   self.assertIn('Guide interference',x['cause_4d'])
   self.assertNotIn('Inspection gap',x['cause_4d'])
   self.assertIn('Inspection gap',x['leak_cause'])

 def test_other_4d_table_items_append_under_occurrence_cause(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'extra_4d.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   tb=sl.shapes.add_table(5,2,Inches(0.8),Inches(0.6),Inches(7.5),Inches(5.0)).table
   rows=[
    ('Root Cause','Guide interference'),
    ('Why Made','Design review gap'),
    ('Escape Cause','Inspection control missing'),
    ('System Cause','Control plan not linked'),
    ('Corrective Action','Guide revised'),
   ]
   for r,(a,b) in enumerate(rows):
    tb.cell(r,0).text=a; tb.cell(r,1).text=b
   prs.save(p)
   fields,detected=e._table_semantic_fields(p)
   self.assertIn('Guide interference',fields['cause_4d'])
   self.assertIn('- Why Made\nDesign review gap',fields['cause_4d'])
   self.assertNotIn('System Cause',fields['cause_4d'])
   self.assertIn('Control plan not linked',fields['system_cause'])
   self.assertIn('Inspection control missing',fields['leak_cause'])
   self.assertIn('Guide revised',fields['action_5d'])

 def test_translation_coverage_reports_partial_dictionary_conversion(self):
  x=e.enhance_dict({'problem':'Crack occurred after repeated vehicle evaluation.'})
  pct=e.translation_coverage(x)
  self.assertGreater(pct,0)
  self.assertLess(pct,100)

 def test_korean_extra_4d_items_append_but_system_cause_stays_separate(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td)/'korean_extra_4d.pptx'
   prs=Presentation(); sl=prs.slides.add_slide(prs.slide_layouts[6])
   tb=sl.shapes.add_table(6,2,Inches(0.8),Inches(0.6),Inches(8.0),Inches(5.5)).table
   rows=[
    ('발생 원인','가이드 간섭'),
    ('유출 원인','검사 누락'),
    ('시스템 원인','관리항목 미연계'),
    ('재현시험','동일 조건에서 불량 재현'),
    ('추가 검토','공차 영향성 추가 검토'),
    ('5D 개선 대책','가이드 수정'),
   ]
   for r,(a,b) in enumerate(rows):
    tb.cell(r,0).text=a; tb.cell(r,1).text=b
   prs.save(p)
   extras=e._korean_4d_extra_items(p)
   joined='\n\n'.join(extras)
   self.assertIn('- 재현시험\n동일 조건에서 불량 재현',joined)
   self.assertIn('- 추가 검토\n공차 영향성 추가 검토',joined)
   self.assertNotIn('시스템 원인',joined)
   base={'cause_4d':'가이드 간섭','leak_cause':'검사 누락','system_cause':'관리항목 미연계'}
   out=e._augment_korean_4d_extras(p,base)
   self.assertTrue(out['cause_4d'].startswith('가이드 간섭'))
   self.assertIn('- 재현시험\n동일 조건에서 불량 재현',out['cause_4d'])
   self.assertIn('- 추가 검토\n공차 영향성 추가 검토',out['cause_4d'])
   self.assertEqual(out['leak_cause'],'검사 누락')
   self.assertEqual(out['system_cause'],'관리항목 미연계')

 def test_korean_document_keeps_existing_extractor_unchanged(self):
  old_original=e._original; old_blocks=e._shape_blocks
  try:
   e._original=lambda _p:{'cause_4d':'기존 발생원인','leak_cause':'기존 유출원인','system_cause':'기존 시스템원인'}
   e._shape_blocks=lambda _p:['4D 발생 원인','한글 원인 내용','유출 원인','한글 유출 내용']
   x=e.extract('dummy.pptx')
   self.assertEqual(x['cause_4d'],'기존 발생원인')
   self.assertEqual(x['leak_cause'],'기존 유출원인')
   self.assertEqual(x['system_cause'],'기존 시스템원인')
  finally:
   e._original=old_original; e._shape_blocks=old_blocks

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
