"""English/mixed 8D extraction and post-processing.

The legacy extractor is kept for metadata, while 2D~6D body text is rebuilt from
source order so section contents are not lost when a template uses English labels
or only a D-number marker. 7D/8D are boundaries, not 6D content.
"""
import re
from pathlib import Path
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
 ('1D',('team build','team building','team formation','team composition','team members','team member','basic information','general information','팀 구성')),
 ('7D',('customer response','customer action','horizontal deployment','read across','lessons learned','prevent recurrence','preventive action','prevention action','standardization','고객 대응','고객대응','수평 전개','수평전개','재발 방지','재발방지')),
 ('8D',('request items','request item','customer request','requests','request','follow up','follow-up','closure','close out','요청 사항','요청사항','후속 조치','후속조치')),
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
    """Return (section, remainder). D-number and semantic headings are boundaries."""
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
            if low.startswith(ll):
                tail=s[len(lab):]
                # Heading with explicit delimiter, e.g. "Corrective Action: Replace guide"
                if re.match(r'^\s*[:：\-–—]',tail):
                    rest=re.sub(r'^\s*[:：\-–—]\s*','',tail).strip()
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

def _center_box(box):
    x,y,w,h=box
    return x+w/2,y+h/2

def _shape_text(sh):
    if getattr(sh,'has_table',False):
        return '\n'.join(_norm_line(cell.text) for row in sh.table.rows for cell in row.cells if _norm_line(cell.text))
    return str(getattr(sh,'text','') or '').strip()

def _positioned_units(sl):
    """Yield text at cell/shape granularity so one large table cannot bleed across D regions."""
    out=[]
    for order,sh in enumerate(sl.shapes):
        sx,sy,sw,shh=_box(sh)
        if getattr(sh,'has_table',False):
            tb=sh.table
            col_lefts=[]; x=sx
            for col in tb.columns:
                col_lefts.append(x); x+=float(col.width)
            row_tops=[]; y=sy
            for row in tb.rows:
                row_tops.append(y); y+=float(row.height)
            for r,row in enumerate(tb.rows):
                for cc,cell in enumerate(row.cells):
                    t=str(cell.text or '').strip()
                    if not t:continue
                    box=(col_lefts[cc],row_tops[r],float(tb.columns[cc].width),float(row.height))
                    out.append({'text':t,'box':box,'order':order*10000+r*100+cc,'kind':'cell'})
        else:
            t=str(getattr(sh,'text','') or '').strip()
            if t:out.append({'text':t,'box':(sx,sy,sw,shh),'order':order*10000,'kind':'shape'})
    return out

def _d_marker_number(text):
    s=_norm_line(text)
    m=re.fullmatch(r'[①②③④⑤⑥⑦⑧]|[1-8]\s*[dD]',s)
    if not m:return None
    circ={'①':1,'②':2,'③':3,'④':4,'⑤':5,'⑥':6,'⑦':7,'⑧':8}
    return circ.get(s,int(re.search(r'[1-8]',s).group()) if re.search(r'[1-8]',s) else None)

def _semantic_section_for_text(text):
    sec,_=_marker(text)
    return sec

def _spatial_section_blocks(path,return_detected=False):
    """Map 2D~6D using geometry, but let explicit section headings correct row offsets.

    Some company 8D templates place the D-number marker one visual row above the
    actual content. In that layout, pure geometry shifts 3D into 2D, 4D into 3D,
    etc. An explicit heading such as "Containment" or "Root Cause" is therefore
    authoritative for that content and all following lines in the same region.
    """
    prs=Presentation(path); blocks=[]; detected=set()
    valid_sections=set(SECTION_FIELDS) | {'7D','8D'}
    for sl in prs.slides:
        units=_positioned_units(sl)
        markers=[]
        for u in units:
            n=_d_marker_number(u['text'])
            if n:
                cx,cy=_center_box(u['box'])
                markers.append({'n':n,'u':u,'cx':cx,'cy':cy})
        if not markers:continue

        tol=float(prs.slide_width)*0.06
        cols=[]
        for m in sorted(markers,key=lambda z:z['cx']):
            if not cols or abs(m['cx']-cols[-1]['cx'])>tol:
                cols.append({'cx':m['cx'],'markers':[m]})
            else:
                cols[-1]['markers'].append(m)
                cols[-1]['cx']=sum(x['cx'] for x in cols[-1]['markers'])/len(cols[-1]['markers'])

        for ci,col in enumerate(cols):
            ms=sorted(col['markers'],key=lambda z:z['cy'])
            next_col_x=min((x['u']['box'][0] for x in cols[ci+1]['markers']),default=float('inf')) if ci+1<len(cols) else float('inf')
            for i,m in enumerate(ms):
                n=m['n']; mu=m['u']
                if n<2 or n>6:continue
                default_sec=str(n)+'D'
                default_key=SECTION_FIELDS.get(default_sec)
                if default_key:detected.add(default_key)

                top=m['cy']
                bottom=ms[i+1]['cy'] if i+1<len(ms) else float('inf')
                mx,my,mw,mh=mu['box']; mright=mx+mw

                region=[]
                for u in units:
                    if u is mu or _d_marker_number(u['text']):continue
                    ux,uy,uw,uh=u['box']; cx2,cy2=_center_box(u['box'])
                    if cy2<top or cy2>=bottom:continue
                    if ux+uw<=mright*1.02 or ux>=next_col_x:continue
                    region.append((uy,ux,u['order'],u))
                region.sort(key=lambda z:(z[0],z[1],z[2]))

                local_sec=default_sec
                emitted_sec=None

                def ensure_section(sec):
                    nonlocal emitted_sec
                    if sec!=emitted_sec:
                        blocks.append(sec)
                        emitted_sec=sec

                for _uy,_ux,_order,u in region:
                    for raw in u['text'].replace('\r','\n').splitlines():
                        line=_norm_line(raw)
                        if not line or _is_photo_caption(line):continue
                        sem,rest=_marker(line)
                        if sem in valid_sections:
                            # Explicit semantic/D heading wins over a geometrically shifted marker.
                            local_sec=sem
                            key=SECTION_FIELDS.get(sem)
                            if key:detected.add(key)
                            ensure_section(sem)
                            if rest:
                                nested,nested_rest=_marker(rest)
                                if nested==sem:
                                    rest=nested_rest
                                if rest and not _is_photo_caption(rest):
                                    blocks.append(rest)
                            continue

                        ensure_section(local_sec)
                        blocks.append(line)
    return (blocks,detected) if return_detected else blocks

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
            # Do not emit the section name itself (e.g. "Corrective Action" -> "개선대책").
            msec,mrest=_marker(x)
            if msec==sec and not mrest:
                continue
            # Normalize bullets/punctuation so repeated first-row text is not duplicated.
            dq=re.sub(r'^[\s•·▪◦\-–—]+','',q)
            if not dq or dq in seen:continue
            seen.add(dq); vals.append(x)
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

OCCURRENCE_LABELS=(
    '발생','발생일','발생일자','발생 일자',
    'occurrence','occurence','occurrence date','occurence date','date of occurrence'
)
_DATE_RE=re.compile(
    r"(?<!\d)(?:'?\d{2}|\d{4})\s*[./-]\s*\d{1,2}"
    r"(?:\s*[./-]\s*\d{1,2})?"
    r"(?:\s*[~～]\s*(?:\d{1,2}(?:\s*[./-]\s*\d{1,2})?))?"
)

def _date_token(text):
    m=_DATE_RE.search(str(text or ''))
    return re.sub(r'\s+','',m.group(0)) if m else ''

def _occurrence_date_from_blocks(blocks):
    lines=[]
    for block in blocks or []:
        lines.extend(_norm_line(x) for x in str(block or '').replace('\r','\n').splitlines() if _norm_line(x))
    labels=tuple(x.casefold() for x in OCCURRENCE_LABELS)
    for i,line in enumerate(lines):
        low=line.casefold().strip()
        # Exact heading or heading followed by punctuation/date. Generic Korean
        # "발생" is accepted only as a heading, never inside "발생원인".
        matched=None
        for lab in sorted(labels,key=len,reverse=True):
            if low==lab or re.match(r'^'+re.escape(lab)+r'\s*[:：\-–—]?\s*',low):
                tail=re.sub(r'^'+re.escape(lab)+r'\s*[:：\-–—]?\s*','',line,flags=re.I)
                if low==lab or _date_token(tail):
                    matched=lab; break
        if not matched:
            continue
        same=_date_token(line)
        if same:return same
        for nxt in lines[i+1:i+4]:
            tok=_date_token(nxt)
            if tok:return tok
    return ''

_original_parse_year_month=getattr(v310.base,'parse_year_month',None)
def parse_year_month_extended(text):
    s=str(text or '')
    m=re.search(r"(?<!\d)'?(\d{2}|\d{4})\s*[./-]\s*(\d{1,2})",s)
    if m:
        y=m.group(1)
        return (y if len(y)==4 else '20'+y, str(int(m.group(2))))
    if _original_parse_year_month:
        return _original_parse_year_month(text)
    return '',''
v310.base.parse_year_month=parse_year_month_extended

def enhance_dict(d,raw_text='',section_blocks=None,authoritative_keys=None):
    d=dict(d or {})
    if section_blocks is not None:
        rebuilt=extract_sections_from_blocks(section_blocks)
        # Region-based extraction is authoritative when a section was found.
        authoritative_keys=set(authoritative_keys or ())
        for key,val in rebuilt.items():
            if key in authoritative_keys or val:
                d[key]=val
    source='\n'.join(str(d.get(k) or '') for k in FIELD_LABELS)
    source=(source+'\n'+str(raw_text or '')).strip()
    occ=_occurrence_date_from_blocks(str(raw_text or '').splitlines())
    if occ:
        d['occurrence_date']=occ
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
    try:spatial,detected=_spatial_section_blocks(path,True)
    except Exception:spatial,detected=[],set()

    raw='\n'.join(blocks)
    if spatial:
        # Geometry is primary, but a D marker can sit one row below its real section.
        # After semantic remapping that can leave the earlier section empty. Fill only
        # those empty sections from source-order semantic headings; never overwrite a
        # non-empty spatial result.
        spatial_fields=extract_sections_from_blocks(spatial)
        source_fields=extract_sections_from_blocks(blocks)
        for key in spatial_fields:
            sval=str(spatial_fields.get(key) or '').strip()
            fval=str(source_fields.get(key) or '').strip()
            if key in detected:
                d[key]=sval or fval
            elif sval:
                d[key]=sval
        return enhance_dict(d,raw)

    return enhance_dict(d,raw,blocks,None)
v310.base.extract=extract

def apply_english_choice(d,use_original):
    d=dict(d or {})
    if use_original:
        for key in FIELD_LABELS:
            original=d.get(key+'_en_original')
            if original:d[key]=original
    return d

def _count_translatable_english_words(text):
    """Count ordinary English words while ignoring technical IDs/model tokens."""
    s=str(text or '')
    s=re.sub(r'\b[A-Z0-9][A-Z0-9_.+\-/()]*\b',' ',s)
    return len(re.findall(r'[A-Za-z]{2,}',s))

def translation_coverage(d):
    """Estimated dictionary-based Korean conversion coverage for detected English fields."""
    total=0; remaining=0
    for key in FIELD_LABELS:
        original=str(d.get(key+'_en_original') or '')
        if not original:continue
        translated=str(d.get(key) or '')
        total+=_count_translatable_english_words(original)
        remaining+=_count_translatable_english_words(translated)
    if total<=0:
        return 100 if d.get('_english_detected') else 0
    converted=max(0,total-remaining)
    return max(0,min(100,int(round(converted*100.0/total))))

def _ask_translation_choice(self,d,path):
    """Ask at preview-time whether to view/use the dictionary-translated Korean text."""
    if not d.get('_english_detected'):
        self._english_preview_choice=(str(path),False)
        return False
    pct=translation_coverage(d)
    msg=(
        '영문 8D를 감지했습니다.\n\n'
        f'현재 자동 한글 번역률은 약 {pct}%입니다.\n'
        '한글로 변환된 내용을 기준으로 미리보기할까요?\n\n'
        '· 한글 번역 : 자동 변환된 한글 내용 사용\n'
        '· 영문 원문 : 원문 그대로 사용'
    )
    choice=ui.dialog(
        self,'영문 8D 번역 선택',msg,'question',
        (('영문 원문',False),('한글 번역',True)),
        width=560,height=330
    )
    if choice is None:
        return None
    use_original=not bool(choice)
    self._english_preview_choice=(str(path),use_original)
    self._english_translation_percent=pct
    return use_original

# Korean/English 6D status terms use one decision engine for Issue DB and weekly Signal.
_original_status=v310.base.status

_PENDING_STATUS=(
    '진행 중','진행중','검증 중','검증중','예정','계획','미완료',
    'in progress','ongoing','planned','plan to','pending','scheduled','tbd',
    'to be verified','under verification','under validation','not completed'
)
_CLOSE_STATUS=(
    '검증 완료','검증완료','개선 완료','개선완료','정상','이상 없음','이상없음','양호',
    'no abnormality','not abnormal','no defect','no issue','normal result',
    'verification complete','verification completed','effectiveness confirmed',
    'validated','verified','completed','complete','passed','pass','acceptable'
)
_OPEN_STATUS=(
    '불량','이상 발생','이상발생','미흡','재발','부적합','ng','nok',
    'abnormal','abnomal','abnormality','fail','failed','failure',
    'not ok','out of spec','out-of-spec','oos','defect remains','issue remains',
    'recurred','recurrence','not acceptable'
)

def _status_decision(text):
    raw=str(text or '').strip()
    if not raw:return 'open','6D 내용 없음'
    q=re.sub(r'\s+',' ',raw).casefold()
    if any(x.casefold() in q for x in _PENDING_STATUS):
        return 'open','진행/예정 표현 감지'
    # Explicit normal/no-abnormal wording must win before "abnormality" checks.
    if any(x.casefold() in q for x in _CLOSE_STATUS):
        return 'close','완료/정상 표현 감지'
    if any(x.casefold() in q for x in _OPEN_STATUS):
        return 'open','이상/실패 표현 감지'
    if re.search(r'\b(?:abnormal|abnomal|ng|nok|fail(?:ed|ure)?|oos)\b',q):
        return 'open','이상/실패 표현 감지'
    if re.search(r'\b(?:pass(?:ed)?|verified|validated|complete(?:d)?|normal|ok)\b',q):
        return 'close','완료/정상 표현 감지'
    # 6D text exists but does not prove completion: propose open and require final user confirmation.
    return 'open','6D 내용은 있으나 완료/정상 판단어가 명확하지 않음'

def status_bilingual(d):
    decision,_reason=_status_decision(d.get('verification_6d'))
    if decision=='close':
        return v310.impl.base.STATUS['complete'] if hasattr(v310.impl.base,'STATUS') else '개선 완료'
    if str(d.get('action_5d') or '').strip() or str(d.get('verification_6d') or '').strip():
        return v310.impl.base.STATUS['verify'] if hasattr(v310.impl.base,'STATUS') else '개선 검증중'
    return _original_status(d)
v310.base.status=status_bilingual

def judge_status_bilingual(d):
    text=str(d.get('verification_6d') or '').strip()
    decision,_reason=_status_decision(text)
    return decision,text
step9._judge_issue_status=judge_status_bilingual

_original_preview=v3.EnterpriseAppV3.preview
_original_run=v3.EnterpriseAppV3.run

def _preview_with_translation_choice(self):
    g=self.gui(); path=g.get('ppt8d','').strip()
    if not path:
        return ui.warning(self,'입력 확인','8D 원본 PPT를 선택해 주세요.')
    try:
        self.status_var.set('ANALYZING · 8D 내용을 추출하고 있습니다...')
        self._set_progress(15,'8D 내용 추출 중')
        d=extract(path)
        use_original=_ask_translation_choice(self,d,path)
        if use_original is None:
            self.status_var.set('READY · 미리보기가 취소되었습니다.')
            return
        d=apply_english_choice(d,use_original)
        self._last_preview_path=path
        self.log.delete('1.0','end')
        self.log.insert('end','[ 8D 추출 완료 ]\n'+'─'*72+'\n')
        self.log.insert('end',f'원본 파일       : {Path(path).name}\n')
        if d.get('_english_detected'):
            pct=getattr(self,'_english_translation_percent',translation_coverage(d))
            mode='영문 원문' if use_original else '한글 번역'
            self.log.insert('end',f'영문 처리       : {mode} 선택 · 자동 한글 번역률 약 {pct}%\n')
        self.log.insert('end','상태            : 8D 내용 추출 완료 · 미리보기 확인 가능\n')
        self.log.insert('end','※ 아래 [8D 내용 미리보기] 창에서 2D~8D 추출 내용을 확인해 주세요.\n')
        self.status_var.set('READY · 8D 추출 완료')
        self._show_preview_window(d)
    except Exception as e:
        self.status_var.set('ERROR · 추출 실패')
        ui.error(self,'미리보기 오류',repr(e))

def _run_with_translation_confirmation(self):
    g=self.gui(); path=g.get('ppt8d','').strip()
    if not path:
        return _original_run(self)

    probe=extract(path)
    cached=getattr(self,'_english_preview_choice',None)
    use_original=None
    if cached and cached[0]==str(path):
        use_original=bool(cached[1])
    elif probe.get('_english_detected'):
        use_original=_ask_translation_choice(self,probe,path)
        if use_original is None:
            self.status_var.set('READY · 실행이 취소되었습니다.')
            return

    if use_original is None:
        return _original_run(self)

    original_extract=v310.base.extract
    def chosen_extract(p):
        result=original_extract(p)
        if str(p)==str(path):
            return apply_english_choice(result,use_original)
        return result
    v310.base.extract=chosen_extract
    try:
        return _original_run(self)
    finally:
        v310.base.extract=original_extract

v3.EnterpriseAppV3.preview=_preview_with_translation_choice
v3.EnterpriseAppV3.run=_run_with_translation_confirmation

