"""Offline English/mixed 8D post-processing.
Maps common English 8D labels to the existing Korean fields and translates frequent quality phrases.
Identifiers, numbers, units, model/project names are preserved. No network/API required.
"""
import re
import main_v310 as v310

FIELD_LABELS={
 'problem':('problem description','problem statement','problem symptom','failure description','symptom','2d'),
 'temporary_action':('interim containment action','containment action','temporary action','ica','3d'),
 'cause_4d':('root cause','occurrence cause','cause of occurrence','4d root cause'),
 'leak_cause':('escape cause','escape root cause','non detection cause','detection cause'),
 'system_cause':('system cause','systemic cause'),
 'action_5d':('permanent corrective action','corrective action','pca','5d'),
 'verification_6d':('verification','effectiveness verification','validation result','6d'),
}
PHRASES=(
 ('root cause','발생원인'),('escape cause','유출원인'),('problem description','현상'),
 ('corrective action','개선대책'),('containment action','임시조치'),('temporary action','임시조치'),
 ('verification result','효과검증 결과'),('verification','효과검증'),
 ('confirmed','확인됨'),('identified','확인됨'),('occurred','발생'),('occurrence','발생'),
 ('failure','불량'),('defect','불량'),('crack','크랙'),('leakage','누설'),('leak','누설'),
 ('interference','간섭'),('insufficient','부족'),('excessive','과다'),('deformation','변형'),
 ('assembly','조립'),('fastening','체결'),('torque','토크'),('welding','용접'),
 ('process condition','공정조건'),('design margin','설계마진'),('tolerance','공차'),
 ('drawing','도면'),('supplier','협력사'),('incoming','입고'),('material','소재'),
 ('customer','고객'),('sample','샘플'),('test condition','시험조건'),('test','시험'),
)

def looks_english(s):
    s=str(s or ''); a=len(re.findall(r'[A-Za-z]',s)); k=len(re.findall(r'[가-힣]',s))
    return a>=8 and a>k*2

def translate_offline(s):
    """Conservative terminology translation; never fabricates a free-form translation."""
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
            if pos<0:continue
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

def enhance_dict(d,raw_text=''):
    d=dict(d or {})
    source='\n'.join(str(d.get(k) or '') for k in FIELD_LABELS)
    source=(source+'\n'+str(raw_text or '')).strip()
    for key,labels in FIELD_LABELS.items():
        current=str(d.get(key) or '').strip()
        if not current:
            current=_label_extract(source,labels)
        if current and looks_english(current):
            d[key+'_en_original']=current
            d[key]=translate_offline(current)
    d['_english_detected']=any(looks_english(str(d.get(k+'_en_original') or '')) for k in FIELD_LABELS)
    return d

_original=v310.base.extract
def extract(path):
    d=_original(path)
    # Existing extractor already collects PPT text into fields; post-process those fields safely.
    return enhance_dict(d)
v310.base.extract=extract
