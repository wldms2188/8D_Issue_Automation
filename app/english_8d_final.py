"""English/mixed 8D extraction and post-processing.

The legacy extractor is kept for metadata, while 2D~6D body text is rebuilt from
source order so section contents are not lost when a template uses English labels
or only a D-number marker. 7D/8D are boundaries, not 6D content.
"""
import re
import main_v310 as v310
import main_enterprise_v3 as v3
import ui_enterprise as ui
import main_recovery_step9 as step9
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

FIELD_LABELS={
 'problem':('problem description','problem statement','problem symptom','failure description','symptom','problem','2d'),
 'temporary_action':('interim containment action','containment action','containment','temporary action','ica','3d'),
 'cause_4d':('root cause','occurrence cause','cause of occurrence','cause analysis','root cause analysis','4d root cause','4d'),
 'leak_cause':('escape cause','escape root cause','non detection cause','non-detection cause','detection cause'),
 'system_cause':('system cause','systemic cause'),
 'action_5d':('permanent corrective action','corrective action','corrective actions','pca','5d'),
 'verification_6d':('verification','effectiveness verification','validation result','validation','6d'),
}
PHRASES=(
 ('root cause','발생원인'),('escape cause','유출원인'),('problem description','현상'),
 ('corrective action','개선대책'),('containment action','임시조치'),('containment','임시조치'),('temporary action','임시조치'),
 ('verification result','효과검증 결과'),('effectiveness verification','효과검증'),('verification','효과검증'),
 ('confirmed','확인됨'),('identified','확인됨'),('occurred','발생'),('occurrence','발생'),
 ('failure','불량'),('defect','불량'),('crack','크랙'),('leakage','누설'),('leak','누설'),
 ('interference','간섭'),('insufficient','부족'),('excessive','과다'),('deformation','변형'),
 ('assembly','조립'),('fastening','체결'),('torque','토크'),('welding','용접'),
 ('process condition','공정조건'),('design margin','설계마진'),('tolerance','공차'),
 ('drawing','도면'),('supplier','협력사'),('incoming','입고'),('material','소재'),
 ('customer','고객'),('sample','샘플'),('test condition','시험조건'),('test','시험'),
)

SECTION_FIELDS={
 '1D':None,'2D':'problem','3D':'temporary_action','4D':'cause_4d',
 '4D_LEAK':'leak_cause','4D_SYSTEM':'system_cause',
 '5D':'action_5d','6D':'verification_6d',
}
# Ordered from specific to broad. 7D/8D intentionally have no destination field.
SEMANTIC_MARKERS=(
 ('4D_LEAK',('escape root cause','escape cause','non detection cause','non-detection cause','detection cause','유출 원인','유출원인')),
 ('4D_SYSTEM',('systemic cause','system cause','시스템 원인','시스템원인')),
 ('3D',('interim containment action','interim containment','containment action','containment actions','containment','temporary action','temporary actions','immediate action','immediate actions','short term action','short-term action','short term corrective action','interim action','interim actions','protective action','sorting action','customer protection','임시 조치','임시조치','임시 대책','임시대책')),
 ('5D',('permanent corrective action','permanent corrective actions','permanent action','corrective actions','corrective action','countermeasure','countermeasures','corrective measure','corrective measures','solution','solutions','improvement action','improvement actions','개선 대책','개선대책','개선 사항','개선사항')),
 ('6D',('effectiveness verification','effectiveness validation','verification result','verification results','validation result','validation results','verification','validation','effectiveness check','effectiveness confirmation','effect confirmation','result confirmation','confirm effectiveness','효과 검증','효과검증','유효성 검증','유효성검증')),
 ('4D',('root cause analysis','root cause investigation','cause analysis','cause investigation','occurrence cause','cause of occurrence','cause of defect','failure cause','defect cause','root cause','why analysis','5 why','5why','발생 원인','발생원인','원인 분석','원인분석')),
 ('2D',('problem description','problem definition','problem statement','problem symptom','failure description','failure phenomenon','defect description','defect phenomenon','issue description','issue phenomenon','phenomenon','problem','symptom','문제 현황','문제현황','불량 현상','불량현상')),
)

def looks_english(s):
    s=str(s or ''); a=len(re.findall(r'[A-Za-z]',s)); k=len(re.findall(r'[가-힣]',s))
    return a>=8 and a>k*2

def translate_offline(s):
    out=str(s or '')
    for en,ko in sorted(PHRASES,key=lambda x:len(x[0]),reverse=True):
        out=re.sub(r'(?i)(?<![A-Za-z])'+re.escape(en)+r'(?![A-Za-z])',ko,out)
    return out

def _label_extract(text,labels):
    lines=[x.strip() for x in str(text or '').splitlines()]
    for i,line in enumerate(lines):
        low=line.casefold()
        for lab in labels:
            pos=low.find(lab)
            if pos<0: continue
            tail=re.sub(r'^\s*[:：\-–—]\s*','',line[pos+len(lab):]).strip()
            if tail:return tail
            vals=[]
            for nxt in lines[i+1:]:
                if not nxt:continue
                nl=nxt.casefold()
                if any(any(x in nl for x in labs) for labs in FIELD_LABELS.values()):break
                vals.append(nxt)
                if len(vals)>=8:break
            if vals:return '\n'.join(vals)
    return ''

def untranslated_english(original,translated):
    if not looks_english(original): return False
    cleaned=re.sub(r'\b[A-Z0-9][A-Z0-9_.+\-/()]*\b',' ',str(translated or ''))
    words=re.findall(r'[A-Za-z]{3,}',cleaned)
    return len(words)>=2

def _norm_line(s):
    return re.sub(r'[ \t]+',' ',str(s or '')).strip()

def _marker(line):
    """Return (section, remainder). D-number has priority over wording."""
    s=_norm_line(line)
    if not s:return None,s
    m=re.match(r'^\s*([1-8])\s*[dD]\b[\s.:：\-–—)]*(.*)$',s)
    if m:
        n=m.group(1); rest=m.group(2).strip()
        if n in ('1','7','8'): return n+'D',rest
        sec=n+'D'
        if n=='4':
            low=rest.casefold()
            if any(x in low for x in ('escape cause','escape root cause','non detection','non-detection','유출원인','유출 원인')): sec='4D_LEAK'
            elif any(x in low for x in ('system cause','systemic cause','시스템원인','시스템 원인')): sec='4D_SYSTEM'
        return sec,rest
    low=s.casefold()
    for sec,labels in SEMANTIC_MARKERS:
        for lab in labels:
            ll=lab.casefold()
            if low==ll:
                return sec,''
            # Treat "Heading: content" as a heading, but do not strip ordinary
            # content sentences such as "Root cause item A".
            if low.startswith(ll):
                tail=s[len(lab):]
                if re.match(r'^\\s*[:：\\-–—]',tail):
                    rest=re.sub(r'^\\s*[:：\\-–—]\\s*','',tail).strip()
                    return sec,rest
    return None,s

def _strip_semantic_label(text,sec):
    s=_norm_line(text); low=s.casefold()
    for ssec,labels in SEMANTIC_MARKERS:
        if ssec!=sec:continue
        for lab in labels:
            if low.startswith(lab.casefold()):
                return s[len(lab):].lstrip(' :：-–—')
    return s


# Short photo captions are useful for image collage context but must not become 2D~6D text.
CAPTION_HINTS=('process','improvement','before','after','sample','photo','image','view','detail','unloading','loading','공정','개선전','개선후','사진','이미지')

def _is_photo_caption(line):
    s=_norm_line(line)
    if not s or len(s)>80:return False
    q=s.strip('[]() ').casefold()
    if re.fullmatch(r'(before|after)( improvement)?',q):return True
    if re.fullmatch(r'.{0,30}(process|photo|image|view)',q):return True
    if any(x in q for x in ('photo','image','picture','close-up','close up')):return True
    if (s.startswith('[') and s.endswith(']')) and any(x in q for x in CAPTION_HINTS):return True
    return False

def _box(sh):
    try:return (float(sh.left),float(sh.top),float(sh.width),float(sh.height))
    except Exception:return (0,0,0,0)

def _center(sh):
    x,y,w,h=_box(sh);return x+w/2,y+h/2

def _shape_text(sh):
    if getattr(sh,'has_table',False):
        return '\n'.join(_norm_line(cell.text) for row in sh.table.rows for cell in row.cells if _norm_line(cell.text))
    return str(getattr(sh,'text','') or '').strip()

def _d_marker_number(text):
    s=_norm_line(text)
    m=re.fullmatch(r'[①②③④⑤⑥⑦⑧]|[1-8]\s*[dD]',s)
    if not m:return None
    circ={'①':1,'②':2,'③':3,'④':4,'⑤':5,'⑥':6,'⑦':7,'⑧':8}
    return circ.get(s,int(re.search(r'[1-8]',s).group()) if re.search(r'[1-8]',s) else None)

def _semantic_section_for_text(text):
    hits=[]
    low=_norm_line(text).casefold()
    for sec,labels in SEMANTIC_MARKERS:
        for lab in labels:
            if lab.casefold() in low:
                hits.append((len(lab),sec))
    return max(hits)[1] if hits else None

def _spatial_section_blocks(path):
    """Map content by the physical D-marker position, independent of English heading wording."""
    prs=Presentation(path); blocks=[]
    for sl in prs.slides:
        shapes=list(sl.shapes); markers=[]
        for sh in shapes:
            n=_d_marker_number(_shape_text(sh))
            if n:markers.append((n,sh))
        if not markers:continue
        # Most 8D forms use two columns. A D marker starts a region that extends to
        # the next D marker below it in the same column; the opposite column is independent.
        xs=sorted(_center(sh)[0] for _,sh in markers)
        mid=(min(xs)+max(xs))/2 if len(xs)>1 else None
        grouped={}
        for n,sh in markers:
            cx,cy=_center(sh); col=0 if mid is None or cx<=mid else 1
            grouped.setdefault(col,[]).append((cy,n,sh))
        for col,ms in grouped.items():
            ms.sort()
            for i,(cy,n,msh) in enumerate(ms):
                if n<2 or n>6:continue
                top=cy
                bottom=ms[i+1][0] if i+1<len(ms) else float('inf')
                mx,my,mw,mh=_box(msh); mright=mx+mw
                vals=[]
                for order,sh in enumerate(shapes):
                    if sh is msh:continue
                    sx,sy,sw,shh=_box(sh); cx2,cy2=_center(sh)
                    scol=0 if mid is None or cx2<=mid else 1
                    if scol!=col or cy2<top or cy2>=bottom:continue
                    # Ignore the narrow vertical D-label/section-title band itself.
                    if sx+sw<=mright*1.05:continue
                    t=_shape_text(sh)
                    if not t:continue
                    for line in t.replace('\r','\n').splitlines():
                        line=_norm_line(line)
                        if line and not _is_photo_caption(line):vals.append((sy,sx,order,line))
                vals.sort(key=lambda z:(z[0],z[1],z[2]))
                if vals:
                    blocks.append(str(n)+'D')
                    blocks.extend(v[3] for v in vals)
    return blocks

def extract_sections_from_blocks(blocks):
    """Collect every line under 2D~6D. Unsplit 4D defaults to occurrence cause."""
    out={v:'' for v in SECTION_FIELDS.values() if v}
    buckets={k:[] for k,v in SECTION_FIELDS.items() if v}
    current=None
    for block in blocks:
        for raw in str(block or '').replace('\r','\n').splitlines():
            line=_norm_line(raw)
            if not line:continue
            if _is_photo_caption(line):continue
            sec,rest=_marker(line)
            if sec:
                current=sec
                if sec in ('7D','8D'):
                    continue
                if rest and sec in buckets:
                    buckets[sec].append(rest)
                continue
            if current in buckets:
                buckets[current].append(line)
    for sec,key in SECTION_FIELDS.items():
        if not key:continue
        vals=[]; seen=set()
        for x in buckets[sec]:
            q=re.sub(r'\s+',' ',x).strip().casefold()
            if not q or q in seen:continue
            seen.add(q); vals.append(x)
        out[key]='\n'.join(vals).strip()
    return out

def _shape_blocks(path):
    """Read source in slide/top/left order; tables are emitted cell-by-cell in row order."""
    prs=Presentation(path); blocks=[]
    def walk(shapes):
        for sh in shapes:
            if getattr(sh,'shape_type',None)==MSO_SHAPE_TYPE.GROUP:
                yield from walk(sh.shapes)
            else:
                yield sh
    for sl in prs.slides:
        items=[]
        for order,sh in enumerate(walk(sl.shapes)):
            top=int(getattr(sh,'top',0) or 0); left=int(getattr(sh,'left',0) or 0)
            vals=[]
            if getattr(sh,'has_table',False):
                for row in sh.table.rows:
                    for cell in row.cells:
                        t=_norm_line(cell.text)
                        if t: vals.append(t)
            else:
                t=str(getattr(sh,'text','') or '').strip()
                if t: vals.append(t)
            if vals:items.append((top,left,order,'\n'.join(vals)))
        items.sort(key=lambda x:(x[0],x[1],x[2]))
        blocks.extend(x[3] for x in items)
    return blocks

def enhance_dict(d,raw_text='',section_blocks=None):
    d=dict(d or {})
    if section_blocks is not None:
        rebuilt=extract_sections_from_blocks(section_blocks)
        # Region-based extraction is authoritative when a section was found.
        for key,val in rebuilt.items():
            if val:d[key]=val
    source='\n'.join(str(d.get(k) or '') for k in FIELD_LABELS)
    source=(source+'\n'+str(raw_text or '')).strip()
    for key,labels in FIELD_LABELS.items():
        current=str(d.get(key) or '').strip()
        if not current:current=_label_extract(source,labels)
        if current and looks_english(current):
            d[key+'_en_original']=current
            d[key]=translate_offline(current)
            d[key+'_translation_incomplete']=untranslated_english(current,d[key])
    d['_english_detected']=any(bool(d.get(k+'_en_original')) for k in FIELD_LABELS)
    d['_english_translation_incomplete']=any(bool(d.get(k+'_translation_incomplete')) for k in FIELD_LABELS)
    return d

_original=v310.base.extract

def _ppt_text(path):
    try:return '\n'.join(_shape_blocks(path))
    except Exception:return ''

def extract(path):
    d=_original(path)
    try:blocks=_shape_blocks(path)
    except Exception:blocks=[]
    try:spatial=_spatial_section_blocks(path)
    except Exception:spatial=[]
    section_blocks=[]
    if spatial:section_blocks.extend(spatial)
    if blocks:section_blocks.extend(blocks)
    return enhance_dict(d,'\n'.join(blocks),section_blocks)
v310.base.extract=extract

def apply_english_choice(d,use_original):
    d=dict(d or {})
    if use_original:
        for key in FIELD_LABELS:
            original=d.get(key+'_en_original')
            if original:d[key]=original
    return d

# English status terms must behave like their Korean equivalents.
_original_status=v310.base.status
def status_bilingual(d):
    v=str(d.get('verification_6d') or '').strip()
    q=re.sub(r'\s+',' ',v).casefold()
    if any(x in q for x in ('in progress','ongoing','planned','plan to','pending','scheduled','tbd','to be verified','under verification')):
        return v310.impl.base.STATUS['verify'] if hasattr(v310.impl.base,'STATUS') else '개선 검증중'
    if any(x in q for x in ('verification complete','verification completed','verified','validated','effectiveness confirmed','completed','passed')):
        return v310.impl.base.STATUS['complete'] if hasattr(v310.impl.base,'STATUS') else '개선 완료'
    return _original_status(d)
v310.base.status=status_bilingual

_original_judge=step9._judge_issue_status
def judge_status_bilingual(d):
    text=str(d.get('verification_6d') or '').strip()
    if not text:return 'open',''
    q=re.sub(r'\s+',' ',text).casefold()
    if any(x in q for x in ('진행 중','진행중','예정','in progress','ongoing','planned','plan to','pending','scheduled','tbd','to be verified','under verification')):
        return 'open',text
    return 'close',text
step9._judge_issue_status=judge_status_bilingual

_original_run=v3.EnterpriseAppV3.run
def _run_with_translation_confirmation(self):
    g=self.gui(); path=g.get('ppt8d','').strip()
    if path:
        probe=extract(path)
        if probe.get('_english_translation_incomplete'):
            use_original=ui.ask_yes_no(
                self,'영문 번역 확인',
                '일부 영문 표현으로 인해 영문 전체가 번역되지는 못했습니다.\n\n'
                '전체 영문 내용으로 입력하시겠습니까?\n\n'
                '예: 전체 영문 원문으로 입력\n'
                '아니오: 자동 번역된 한글 내용으로 입력')
            original_extract=v310.base.extract
            def chosen_extract(p):
                return apply_english_choice(original_extract(p),use_original)
            v310.base.extract=chosen_extract
            try:return _original_run(self)
            finally:v310.base.extract=original_extract
    return _original_run(self)
v3.EnterpriseAppV3.run=_run_with_translation_confirmation
