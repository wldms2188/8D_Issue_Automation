from pathlib import Path
import io, sys, traceback
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app'))

# Import the exact production module under test.
import main_v308 as v308

OUT = ROOT / 'test_output'
OUT.mkdir(exist_ok=True)


def make_png(label):
    im = Image.new('RGB', (500, 260), 'white')
    d = ImageDraw.Draw(im)
    d.rectangle((5, 5, 495, 255), outline='black', width=4)
    d.text((30, 110), label, fill='black')
    b = io.BytesIO(); im.save(b, 'PNG'); return b.getvalue()


def add_text(slide, text, x, y, w=1.5, h=.4, size=12):
    sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    sh.text = text
    for p in sh.text_frame.paragraphs:
        for r in p.runs: r.font.size = Pt(size)
    return sh


def make_weekly(path):
    prs = Presentation()
    prs.slide_width = Inches(11.0); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    s1 = prs.slides[0]
    s1.shapes.add_table(2, 6, Inches(.4), Inches(.5), Inches(10), Inches(1.2))
    tb = s1.shapes[-1].table
    headers = ['이슈', '과제명', '문제/현상', '진행 현황', 'Signal', '기타']
    for i,h in enumerate(headers): tb.cell(0,i).text = h
    for i in range(6): tb.cell(1,i).text = ''
    add_text(s1, '000팀 주요 논의 사항', .5, 2.0, 5, .5, 16)
    s2 = prs.slides.add_slide(blank)
    # Template markers intentionally start at arbitrary positions for NEW issue test.
    for label,x,y in [('2D',.2,2.2),('3D',3.0,3.0),('4D',.3,5.0),('5D',7.0,2.0),('6D',8.0,5.0)]:
        sh = add_text(s2, label, x, y, .55, .45, 12); sh.name = 'TEMPLATE_'+label
    add_text(s2, '발생단계 (발생일자)', .5, .5, 3, .4, 10)
    prs.save(path)


def make_source(path):
    prs = Presentation()
    prs.slide_width = Inches(11.0); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    # Metadata slide
    s = prs.slides[0]
    add_text(s, '고객사: Ford', .5, .3, 3, .4)
    add_text(s, 'Ford_TestTask', .5, .8, 4, .4)
    # Section slides with text and images.
    sections = [
        ('2D', 'Problem: battery leakage observed'),
        ('3D', 'Temporary action: quarantine lot and inspect'),
        ('4D', 'Cause: sealing process variation'),
        ('5D', 'Action: update process control and fixture'),
        ('6D', 'Verification: reliability test passed'),
    ]
    for label,text in sections:
        sl = prs.slides.add_slide(blank)
        add_text(sl, label, .5, .4, 1, .4, 16)
        add_text(sl, text, .5, 1.0, 6, .7, 12)
        sl.shapes.add_picture(io.BytesIO(make_png(label+' IMAGE 1')), Inches(7), Inches(1), width=Inches(3), height=Inches(1.8))
        sl.shapes.add_picture(io.BytesIO(make_png(label+' IMAGE 2')), Inches(7), Inches(3.1), width=Inches(3), height=Inches(1.8))
    prs.save(path)


def run_case(mode):
    weekly = OUT / f'weekly_{mode}.pptx'
    src = OUT / 'source_8d.pptx'
    make_weekly(weekly)
    make_source(src)
    d = {
        'customer':'Ford',
        'issue_name':'Ford_TestIssue',
        'task_name':'TestTask',
        'problem':'Problem: battery leakage observed',
        'temporary_action':'Temporary action: quarantine lot and inspect',
        'customer_response':'Customer response: containment completed',
        'cause_4d':'Cause: sealing process variation',
        'leak_cause':'Leak cause: sealing defect',
        'system_cause':'System cause: process control gap',
        'action_5d':'Action: update process control and fixture',
        'verification_6d':'Verification: reliability test passed',
        'team':'Pack개발품질1',
        'occurrence_stage':'양산',
        'occurrence_date':'2026-09-10',
        'status':'개선 완료',
        '_section_images':{},
    }
    # Controlled section images: two per D, exactly the type v3.0.8 expects.
    blobs=[]
    for label in ['2D','3D','4D','5D','6D']:
        items=[]
        for n in (1,2):
            b=make_png(label+f' TEST {n}')
            items.append((b,(0,0,1,1),1))
        d['_section_images'][label]=items
    g={}
    # Call the actual patched weekly entrypoint.
    result = v308.base.weekly(str(weekly), str(OUT / f'output_{mode}.pptx'), d, g, mode)
    out = Path(result[1])
    assert out.exists(), f'Output PPT not created: {out}'
    check = Presentation(out)
    assert len(check.slides) >= 2
    page1 = check.slides[0]
    texts = '\n'.join(sh.text for sh in page1.shapes if hasattr(sh,'text'))
    assert 'TestTask' in texts, 'Task not written to summary'
    assert 'Ford_TestIssue' in texts, 'Issue not written to summary'
    assert 'Pack개발품질1팀 주요 논의 사항' in texts, 'Team title not updated'
    page2 = check.slides[1]
    page2_text = '\n'.join(sh.text for sh in page2.shapes if hasattr(sh,'text'))
    for expected in ['Problem: battery leakage observed','Temporary action: quarantine lot and inspect','Cause: sealing process variation','Action: update process control and fixture','Verification: reliability test passed']:
        assert expected in page2_text, f'Missing detail text: {expected}'
    return str(out)


if __name__ == '__main__':
    try:
        print('TEST v3.0.8 synthetic NEW issue')
        print(run_case('신규'))
        print('TEST v3.0.8 synthetic EXISTING issue')
        print(run_case('기존'))
        print('RESULT: PASS')
    except Exception:
        traceback.print_exc()
        print('RESULT: FAIL')
        raise
