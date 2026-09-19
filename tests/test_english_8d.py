import sys,unittest
from pathlib import Path
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

if __name__=='__main__':unittest.main()
