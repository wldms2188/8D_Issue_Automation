"""English/mixed 8D extraction and post-processing.

The legacy extractor is kept for metadata, while 2D~6D body text is rebuilt from
source order so section contents are not lost when a template uses English labels
or only a D-number marker. 7D/8D are boundaries, not 6D content.
"""
import copy
import re
from pathlib import Path
import main_v310 as v310
import main_enterprise_v3 as v3
import ui_enterprise as ui
import main_recovery_step9 as step9
import project_autocomplete_final as project_catalog
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
    """Map content strictly by the visible D-circle/number regions.

    D marker geometry is authoritative for 2D/3D/4D/5D/6D. Text headings are
    retained as content and must never move a line into a different D section.
    If a template has two separate 4D marker columns, the left one defaults to
    occurrence/root cause and the next one to escape/leak cause.
    """
    prs=Presentation(path); blocks=[]; detected=set()
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

        four_cols=sorted({round(m['cx'],2) for m in markers if m['n']==4})

        for ci,col in enumerate(cols):
            ms=sorted(col['markers'],key=lambda z:z['cy'])
            next_col_x=min((x['u']['box'][0] for x in cols[ci+1]['markers']),default=float('inf')) if ci+1<len(cols) else float('inf')
            for i,m in enumerate(ms):
                n=m['n']; mu=m['u']
                if n<2 or n>6:continue

                region_sec=str(n)+'D'
                if n==4 and len(four_cols)>1:
                    nearest=min(range(len(four_cols)),key=lambda j:abs(four_cols[j]-round(m['cx'],2)))
                    region_sec='4D' if nearest==0 else '4D_LEAK'

                key=SECTION_FIELDS.get(region_sec)
                if key:detected.add(key)
                if region_sec=='4D':
                    detected.add('cause_4d')
                elif region_sec=='4D_LEAK':
                    detected.add('leak_cause')

                # The D marker is the centre/reference of its section, not the
                # top edge of the next content band. Use midpoints to neighbouring
                # D markers. This prevents 3D text just above the 3D marker from
                # being misread as 2D, and 4D text as 3D.
                if i>0:
                    top=(ms[i-1]['cy']+m['cy'])/2
                elif i+1<len(ms):
                    top=max(0.0,m['cy']-(ms[i+1]['cy']-m['cy'])*0.45)
                else:
                    top=0.0
                if i+1<len(ms):
                    bottom=(m['cy']+ms[i+1]['cy'])/2
                elif i>0:
                    bottom=min(float(prs.slide_height),m['cy']+(m['cy']-ms[i-1]['cy'])*0.45)
                else:
                    bottom=float(prs.slide_height)
                mx,my,mw,mh=mu['box']; mright=mx+mw
                vals=[]
                for u in units:
                    if u is mu or _d_marker_number(u['text']):continue
                    ux,uy,uw,uh=u['box']; _cx,cy=_center_box(u['box'])
                    if cy<top or cy>=bottom:continue
                    if ux+uw<=mright*1.02 or ux>=next_col_x:continue
                    for raw in u['text'].replace('\r','\n').splitlines():
                        line=_norm_line(raw)
                        if not line or _is_photo_caption(line):continue
                        vals.append((uy,ux,u['order'],line))
                vals.sort(key=lambda z:(z[0],z[1],z[2]))

                # A single 4D band can contain Root/Escape/System cause columns
                # side-by-side. If those 4D sub-headings are horizontally separated,
                # use their x positions only to split WITHIN 4D; never across D sections.
                if region_sec in ('4D','4D_LEAK') and vals:
                    heads=[]
                    for vy,vx,vo,line in vals:
                        fk,rest=_four_d_subfield(line)
                        if fk:heads.append((vy,vx,vo,fk,rest))
                    horizontal=(len(heads)>=2 and (max(h[1] for h in heads)-min(h[1] for h in heads))>float(prs.slide_width)*0.08)
                    if horizontal:
                        sec_token={'cause_4d':'4D','leak_cause':'4D_LEAK','system_cause':'4D_SYSTEM'}
                        for vy,vx,vo,line in vals:
                            fk,rest=_four_d_subfield(line)
                            if fk:
                                detected.add(fk)
                                if rest:
                                    blocks.append(sec_token[fk]); blocks.append(rest)
                                continue
                            # Prefer a heading above/same-row, then nearest x column.
                            eligible=[h for h in heads if h[0]<=vy+float(prs.slide_height)*0.015] or heads
                            h=min(eligible,key=lambda z:(abs(vx-z[1]),abs(vy-z[0])))
                            detected.add(h[3]); blocks.append(sec_token[h[3]]); blocks.append(line)
                        continue

                blocks.append(region_sec)
                blocks.extend(v[3] for v in vals)

    return (blocks,detected) if return_detected else blocks

FOUR_D_SUBLABELS={
    'system_cause':(
        'systemic root cause','system root cause','systemic cause','system cause',
        'management system cause','system cause analysis'
    ),
    'leak_cause':(
        'escape root cause','escape cause','escape point cause','escape point','point of escape',
        'non detection cause','non-detection cause','detection failure cause',
        'why not detected','detection cause'
    ),
    'cause_4d':(
        'occurrence root cause','root cause of occurrence','occurrence cause',
        'technical root cause','direct cause','root cause analysis','root cause'
    ),
}

def _four_d_subfield(line):
    """Recognize only 4D sub-headings; never changes the surrounding D region."""
    s=_norm_line(line); low=s.casefold()
    for key,labels in FOUR_D_SUBLABELS.items():
        for lab in sorted(labels,key=len,reverse=True):
            ll=lab.casefold()
            if low==ll:
                return key,''
            if low.startswith(ll):
                tail=s[len(lab):]
                # In 4D only, tolerate a label followed directly by its value.
                rest=re.sub(r'^\s*[:：\-–—)]*\s*','',tail).strip()
                return key,rest
    return None,s

def _english_fields_from_spatial(blocks):
    """Build English 2D~6D fields from D geometry without cross-D semantic reassignment."""
    out={k:'' for k in ('problem','temporary_action','cause_4d','leak_cause','system_cause','action_5d','verification_6d')}
    buckets={k:[] for k in out}
    detected=set()
    current=None
    sub4='cause_4d'
    direct={'2D':'problem','3D':'temporary_action','5D':'action_5d','6D':'verification_6d'}

    for block in blocks or []:
        for raw in str(block or '').replace('\r','\n').splitlines():
            line=_norm_line(raw)
            if not line or _is_photo_caption(line):continue

            # These tokens are inserted by the geometry mapper itself.
            if line in ('2D','3D','4D','4D_LEAK','4D_SYSTEM','5D','6D'):
                current=line
                if current in direct:
                    detected.add(direct[current])
                elif current=='4D':
                    sub4='cause_4d'; detected.add('cause_4d')
                elif current=='4D_LEAK':
                    sub4='leak_cause'; detected.add('leak_cause')
                else:
                    sub4='system_cause'; detected.add('system_cause')
                continue

            if current in direct:
                key=direct[current]
                sec,rest=_marker(line)
                # Remove only the heading that agrees with the current geometric D.
                # A mismatched heading remains literal content; it never moves sections.
                if sec==current:
                    if rest:buckets[key].append(rest)
                else:
                    buckets[key].append(line)
                continue

            if current in ('4D','4D_LEAK','4D_SYSTEM'):
                four_key,rest=_four_d_subfield(line)
                if four_key:
                    sub4=four_key; detected.add(sub4)
                    if rest:buckets[sub4].append(rest)
                    continue
                # Any 2D/3D/5D/6D-looking text remains literal 4D content.
                # Only a 4D sub-heading may split occurrence/escape/system cause.
                buckets[sub4].append(line)

    for key,vals in buckets.items():
        seen=set(); cleaned=[]
        for x in vals:
            q=re.sub(r'\s+',' ',x).strip().casefold()
            dq=re.sub(r'^[\s•·▪◦\-–—]+','',q)
            if not dq or dq in seen:continue
            seen.add(dq); cleaned.append(x)
        out[key]='\n'.join(cleaned).strip()
    return out,detected

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

def _table_blocks(path):
    """Read only table cells in slide/shape/row/column order.

    Table item/header meaning is the highest-confidence English section signal.
    """
    prs=Presentation(path); blocks=[]
    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False):
                continue
            for row in sh.table.rows:
                vals=[]
                for cell in row.cells:
                    t=_norm_line(cell.text)
                    if t:vals.append(t)
                if vals:blocks.extend(vals)
    return blocks

def _table_semantic_fields(path):
    """Extract table sections with explicit support for extra 4D items.

    Root/occurrence cause and escape cause keep their normal destinations.
    Any other named row while inside 4D is additionally appended to occurrence
    cause in the requested form:

        - Item name
        item content
    """
    keys=('problem','temporary_action','cause_4d','leak_cause','system_cause','action_5d','verification_6d')
    buckets={k:[] for k in keys}
    detected=set()
    prs=Presentation(path)

    def add(key,text):
        text=_norm_line(text)
        if text and not _is_photo_caption(text):
            buckets[key].append(text)

    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False):
                continue
            in_4d=False
            current_key=None
            for row in sh.table.rows:
                cells=[_norm_line(cell.text) for cell in row.cells]
                nonempty=[x for x in cells if x]
                if not nonempty:
                    continue

                label=nonempty[0]
                sec,rest=_marker(label)

                if sec:
                    if sec in ('1D','7D','8D'):
                        in_4d=False; current_key=None
                        continue

                    key=SECTION_FIELDS.get(sec)
                    in_4d=sec in ('4D','4D_LEAK','4D_SYSTEM')
                    current_key=key
                    if key:detected.add(key)

                    row_values=[]
                    if rest:row_values.append(rest)
                    row_values.extend(nonempty[1:])

                    # System cause already has its own 4D destination.
                    # Do not duplicate it under occurrence cause.
                    if sec=='4D_SYSTEM':
                        content='\n'.join(x for x in row_values if x).strip()
                        if content:add('system_cause',content)
                        continue

                    if key:
                        for value in row_values:add(key,value)
                    continue

                # Any named row with a value while still inside 4D is an "other 4D item".
                # Do not discard it just because its item name is not in our dictionary.
                if in_4d and len(nonempty)>=2:
                    item=nonempty[0]
                    content='\n'.join(nonempty[1:]).strip()
                    if content:
                        buckets['cause_4d'].append(f'- {item}\n{content}')
                        detected.add('cause_4d')
                    continue

                # Single-cell continuation belongs to the last known field.
                if current_key:
                    for value in nonempty:add(current_key,value)

    out={k:'' for k in keys}
    for key,vals in buckets.items():
        seen=set(); cleaned=[]
        for x in vals:
            q=re.sub(r'\s+',' ',x).strip().casefold()
            if not q or q in seen:continue
            seen.add(q); cleaned.append(x)
        out[key]='\n'.join(cleaned).strip()
    return out,detected

def _d_regions_for_slide(sl,prs):
    """Build broad D-section regions using marker midpoints.

    Geometry gates which table belongs to which D area; semantic labels inside
    that area still decide the field, matching the Korean table-first behavior.
    """
    units=_positioned_units(sl)
    markers=[]
    for u in units:
        n=_d_marker_number(u['text'])
        if n:
            cx,cy=_center_box(u['box'])
            markers.append({'n':n,'u':u,'cx':cx,'cy':cy})
    if not markers:return []

    tol=float(prs.slide_width)*0.06
    cols=[]
    for m in sorted(markers,key=lambda z:z['cx']):
        if not cols or abs(m['cx']-cols[-1]['cx'])>tol:
            cols.append({'cx':m['cx'],'markers':[m]})
        else:
            cols[-1]['markers'].append(m)
            cols[-1]['cx']=sum(x['cx'] for x in cols[-1]['markers'])/len(cols[-1]['markers'])

    regions=[]
    for ci,col in enumerate(cols):
        ms=sorted(col['markers'],key=lambda z:z['cy'])
        next_x=min((x['u']['box'][0] for x in cols[ci+1]['markers']),default=float(prs.slide_width)) if ci+1<len(cols) else float(prs.slide_width)
        for i,m in enumerate(ms):
            if m['n']<2 or m['n']>6:continue
            if i>0:
                top=(ms[i-1]['cy']+m['cy'])/2
            elif i+1<len(ms):
                top=max(0.0,m['cy']-(ms[i+1]['cy']-m['cy'])*0.45)
            else:
                top=0.0
            if i+1<len(ms):
                bottom=(m['cy']+ms[i+1]['cy'])/2
            elif i>0:
                bottom=min(float(prs.slide_height),m['cy']+(m['cy']-ms[i-1]['cy'])*0.45)
            else:
                bottom=float(prs.slide_height)
            mx,my,mw,mh=m['u']['box']
            regions.append({
                'n':m['n'],'top':top,'bottom':bottom,
                'left':mx+mw*0.5,'right':next_x,'cx':m['cx'],'cy':m['cy']
            })
    return regions

def _image_region_overlap_fraction(ib,reg):
    """Fraction of an image rectangle physically inside one D region."""
    x,y,w,h=ib
    if w<=0 or h<=0:
        return 0.0
    left=reg['left']; right=reg['right']; top=reg['top']; bottom=reg['bottom']
    iw=max(0.0,min(x+w,right)-max(x,left))
    ih=max(0.0,min(y+h,bottom)-max(y,top))
    return (iw*ih)/max(float(w*h),1.0)


def _english_section_images(path):
    """Map English-source pictures only to the D region that physically owns them.

    The legacy collector used nearest headings, which can attach a 2D image to 3D
    (or vice versa) when English captions/headings are long.  English mode instead
    uses the same D-region geometry that gates text extraction.  Pictures outside
    a valid D region are ignored rather than assigned to the wrong section.
    """
    keys=('2D','3D','4D_CAUSE','4D_LEAK','5D','6D')
    out={k:[] for k in keys}
    prs=Presentation(path)

    # User rule: ONLY source 8D page 1 supplies Issue DB / weekly images.
    # Never borrow or mix pictures from later source pages.
    for si in range(min(1,len(prs.slides))):
        sl=prs.slides[si]
        regions=_d_regions_for_slide(sl,prs)
        if not regions:
            continue

        # 4D sub-headings are used only to split a single 4D region into
        # occurrence-cause vs escape/system-cause image ownership.
        sub4=[]
        for u in _positioned_units(sl):
            sec,_rest=_marker(u.get('text',''))
            if sec not in ('4D','4D_LEAK','4D_SYSTEM'):
                continue
            bx=u.get('box')
            if not bx:
                continue
            cx,cy=_center_box(bx)
            target='4D_LEAK' if sec in ('4D_LEAK','4D_SYSTEM') else '4D_CAUSE'
            sub4.append((target,cx,cy))

        four_regions=sorted([r for r in regions if r['n']==4],key=lambda r:r['cx'])

        for sh in sl.shapes:
            if getattr(sh,'shape_type',None) not in (MSO_SHAPE_TYPE.PICTURE,MSO_SHAPE_TYPE.GROUP):
                continue
            blob=v310.image_blob(sh)
            if not blob:
                continue
            ib=_box(sh)
            if not ib or ib[2]<=0 or ib[3]<=0:
                continue
            # Ignore tiny icons/logos; keep content photos/figures.
            if ib[2] < float(prs.slide_width)*0.025 or ib[3] < float(prs.slide_height)*0.025:
                continue

            cx,cy=_center_box(ib)

            # Ownership is based on how much of the picture is physically inside
            # each D region, not merely where its centre happens to land. A picture
            # split almost evenly across two D regions is ambiguous and is ignored
            # rather than being shown under the wrong D on the weekly page.
            scored=[]
            for r in regions:
                frac=_image_region_overlap_fraction(ib,r)
                if frac>0:
                    scored.append((frac,r))
            if not scored:
                continue
            scored.sort(key=lambda z:z[0],reverse=True)
            best_frac,reg=scored[0]
            second_frac=scored[1][0] if len(scored)>1 else 0.0
            if best_frac<0.35:
                continue
            if second_frac>0 and second_frac>=best_frac*0.85:
                # User rule: if an image is almost evenly split between 2D and 3D,
                # prefer 2D because phenomenon photos are more common than interim-
                # containment photos. Other cross-D near-ties remain excluded.
                tied=scored[:2]
                tied_ns={item[1]['n'] for item in tied}
                if tied_ns=={2,3}:
                    reg=next(item[1] for item in tied if item[1]['n']==2)
                    best_frac=_image_region_overlap_fraction(ib,reg)
                else:
                    continue
            n=reg['n']
            if n in (2,3,5,6):
                sec=f'{n}D'
            elif n==4:
                same4=[
                    h for h in sub4
                    if reg['top']<=h[2]<reg['bottom']
                    and h[1]>=reg['left']*0.90 and h[1]<reg['right']
                ]
                if same4:
                    sec=min(same4,key=lambda h:abs(cx-h[1])+.35*abs(cy-h[2]))[0]
                elif len(four_regions)>=2:
                    nearest_i=min(range(len(four_regions)),key=lambda i:abs(cx-four_regions[i]['cx']))
                    sec='4D_CAUSE' if nearest_i==0 else '4D_LEAK'
                else:
                    mid=(reg['left']+reg['right'])/2.0
                    sec='4D_CAUSE' if cx<mid else '4D_LEAK'
            else:
                continue

            out[sec].append((blob,ib,si))

    for key,vals in out.items():
        vals=sorted(vals,key=lambda z:(z[2],z[1][1],z[1][0]))
        seen=set(); unique=[]
        for item in vals:
            h=hash(item[0])
            if h in seen:
                continue
            seen.add(h); unique.append(item)
        out[key]=unique[:8]
    return out


def _semantic_section_number(sec):
    if sec in ('4D','4D_LEAK','4D_SYSTEM'):
        return 4
    m=re.match(r'^([2-6])D$',str(sec or ''))
    return int(m.group(1)) if m else None

def _cell_in_expected_d_region(box,sec,regions):
    """A semantic table label is accepted only inside its matching D area."""
    n=_semantic_section_number(sec)
    if not n or not regions:
        return True
    cx,cy=_center_box(box)
    candidates=[r for r in regions if r['n']==n]
    if not candidates:
        return True
    for r in candidates:
        if r['top']<=cy<r['bottom'] and cx>=r['left']*0.92 and cx<r['right']:
            return True
    return False

def _scoped_table_semantic_fields(path):
    """Korean-style exact table extraction, limited to the matching D region.

    The section name in the table is authoritative for WHAT the content is.
    D geometry is used only to ensure the row belongs to that D area, preventing
    unrelated tables elsewhere from leaking into 2D~6D.
    """
    keys=('problem','temporary_action','cause_4d','leak_cause','system_cause','action_5d','verification_6d')
    buckets={k:[] for k in keys}
    detected=set()
    prs=Presentation(path)

    def add(key,text):
        text=_norm_line(text)
        if text and not _is_photo_caption(text):
            buckets[key].append(text)

    for sl in prs.slides:
        regions=_d_regions_for_slide(sl,prs)
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False):
                continue
            tb=sh.table
            sx,sy,sw,shh=_box(sh)
            col_left=[]; x=sx
            for col in tb.columns:
                col_left.append(x); x+=float(col.width)
            row_top=[]; y=sy
            for row in tb.rows:
                row_top.append(y); y+=float(row.height)

            for r,row in enumerate(tb.rows):
                cells=[_norm_line(cell.text) for cell in row.cells]
                for cc,raw in enumerate(cells):
                    if not raw:
                        continue
                    sec,rest=_marker(raw)
                    if not sec:
                        continue
                    key=SECTION_FIELDS.get(sec)
                    if not key:
                        continue

                    box=(col_left[cc],row_top[r],float(tb.columns[cc].width),float(row.height))
                    if not _cell_in_expected_d_region(box,sec,regions):
                        continue

                    detected.add(key)
                    vals=[]
                    if rest:
                        vals.append(rest)
                    # Same-row cells to the right are the Korean extractor's first
                    # choice for section content.
                    vals.extend(cells[cc+1:])
                    for v in vals:
                        if v:
                            add(key,v)

                    # Preserve extra named rows that belong to the same 4D band.
                    if sec=='4D':
                        for rr in range(r+1,len(tb.rows)):
                            rcells=[_norm_line(cell.text) for cell in tb.rows[rr].cells]
                            nonempty=[x for x in rcells if x]
                            if not nonempty:
                                continue
                            nsec,_=_marker(nonempty[0])
                            if nsec:
                                break
                            rbox=(col_left[0],row_top[rr],float(tb.columns[0].width),float(tb.rows[rr].height))
                            if not _cell_in_expected_d_region(rbox,'4D',regions):
                                break
                            if len(nonempty)>=2:
                                item=nonempty[0]
                                content='\n'.join(nonempty[1:]).strip()
                                if item and content:
                                    buckets['cause_4d'].append(f'- {item}\n{content}')
                                    detected.add('cause_4d')

    out={k:'' for k in keys}
    for key,vals in buckets.items():
        seen=set(); cleaned=[]
        for x in vals:
            q=re.sub(r'\s+',' ',x).strip().casefold()
            if not q or q in seen:
                continue
            seen.add(q); cleaned.append(x)
        out[key]='\n'.join(cleaned).strip()
    return out,detected

def _semantic_fields_and_detected(blocks):
    fields=extract_sections_from_blocks(blocks or [])
    detected=set()
    for block in blocks or []:
        for raw in str(block or '').replace('\r','\n').splitlines():
            line=_norm_line(raw)
            if not line:
                continue
            sec,_rest=_marker(line)
            key=SECTION_FIELDS.get(sec)
            if key:
                detected.add(key)
    return fields,detected

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

def enhance_dict(d,raw_text='',section_blocks=None,authoritative_keys=None,allow_label_fallback=True):
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
        if not current and allow_label_fallback:current=_label_extract(source,labels)
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

def _korean_4d_extra_items(path):
    """Return Korean 4D table items other than 발생/유출/시스템 원인.

    Examples such as 재현시험, 추가 검토, 분석 결과 are preserved as:
        - 항목명
        내용
    """
    prs=Presentation(path)
    extras=[]

    def ck(text):
        return re.sub(r'[\s_:：/\\\-·•]+','',_norm_line(text)).casefold()

    occurrence=('발생원인','4d발생원인','원인분석','4d원인분석')
    escape=('유출원인','4d유출원인')
    system=('시스템원인','4d시스템원인','시스템적원인','관리시스템원인','it시스템')
    stop=(
        '5d','5d개선대책','개선대책','개선사항','영구개선','시정조치',
        '6d','6d효과검증','효과검증','효과성검증','유효성검증',
        '7d','수평전개','재발방지','8d','요청사항','후속조치'
    )
    structural=('구분','항목','내용','담당','담당자','완료일','일정','비고','결과')

    def matches(q,labels):
        return any(q==x or q.startswith(x) for x in labels)

    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh,'has_table',False):
                continue

            in_4d=False
            current_extra=None

            for row in sh.table.rows:
                cells=[_norm_line(cell.text) for cell in row.cells]
                nonempty=[x for x in cells if x]
                if not nonempty:
                    continue

                label=nonempty[0]
                q=ck(label)

                # Enter/continue 4D on the known 4D headings.
                if matches(q,occurrence) or matches(q,escape) or matches(q,system):
                    in_4d=True
                    current_extra=None
                    continue

                # Explicit 4D title/header also opens the region.
                if q=='4d' or q.startswith('4d원인'):
                    in_4d=True
                    current_extra=None
                    continue

                # 5D+ headings terminate the 4D region.
                if in_4d and matches(q,stop):
                    in_4d=False
                    current_extra=None
                    continue

                if not in_4d:
                    continue

                # Never treat the three standard cause rows as an "extra" item.
                if matches(q,occurrence) or matches(q,escape) or matches(q,system):
                    current_extra=None
                    continue

                # A named 4D row with a value is an additional 4D item.
                if len(nonempty)>=2 and q not in structural:
                    item=label.strip()
                    content='\n'.join(nonempty[1:]).strip()
                    if item and content:
                        extras.append([item,content])
                        current_extra=len(extras)-1
                    continue

                # A following one-cell row can be continuation text for the last extra item.
                if len(nonempty)==1 and current_extra is not None and q not in structural:
                    prev=extras[current_extra][1]
                    extras[current_extra][1]=(prev+'\n'+nonempty[0]).strip()

    out=[]; seen=set()
    for item,content in extras:
        text=f'- {item}\n{content}'.strip()
        key=re.sub(r'\s+',' ',text).casefold()
        if key and key not in seen:
            seen.add(key); out.append(text)
    return out

def _augment_korean_4d_extras(path,d):
    """Append only non-standard Korean 4D items below the existing 발생원인."""
    d=dict(d or {})
    try:
        extras=_korean_4d_extra_items(path)
    except Exception:
        extras=[]
    if not extras:
        return d

    base_text=str(d.get('cause_4d') or '').strip()
    parts=[base_text] if base_text else []
    existing=re.sub(r'\s+',' ',base_text).casefold()
    for extra in extras:
        compact=re.sub(r'\s+',' ',extra).casefold()
        if compact and compact not in existing:
            parts.append(extra)
    d['cause_4d']='\n\n'.join(parts).strip()
    return d

def english_content_ratio(raw):
    """Percentage of Latin letters among Latin + Hangul letters in the PPT text."""
    s=str(raw or '')
    en=len(re.findall(r'[A-Za-z]',s))
    ko=len(re.findall(r'[가-힣]',s))
    total=en+ko
    return (100.0*en/total) if total else 0.0

def _document_looks_english(raw):
    # User rule: English logic only when English content is at least 80%.
    return english_content_ratio(raw)>=80.0

_EXTRACT_CACHE={}

def _extract_cache_key(path):
    try:
        p=Path(path)
        st=p.stat()
        return (str(p.resolve()),int(st.st_mtime_ns),int(st.st_size))
    except Exception:
        return None

def _extract_cache_get(path):
    key=_extract_cache_key(path)
    if key is None:
        return None
    value=_EXTRACT_CACHE.get(key)
    return copy.deepcopy(value) if value is not None else None

def _extract_cache_put(path,value):
    key=_extract_cache_key(path)
    if key is None:
        return value
    # Keep only a few recent source files to avoid retaining large embedded images.
    if key not in _EXTRACT_CACHE and len(_EXTRACT_CACHE)>=3:
        _EXTRACT_CACHE.clear()
    _EXTRACT_CACHE[key]=copy.deepcopy(value)
    return value

def extract(path):
    cached=_extract_cache_get(path)
    if cached is not None:
        return cached

    # Always start from the validated legacy/Korean chain for metadata, images and fallbacks.
    d=_original(path)

    # If PPT metadata could not identify the project, use the user's project catalog
    # against the source 8D filename. Only an unambiguous single match is adopted.
    try:
        filename_projects=project_catalog.filename_project_candidates(path)
    except Exception:
        filename_projects=[]
    d['_filename_project_candidates']=filename_projects
    if not str(d.get('task_name') or '').strip() and len(filename_projects)==1:
        detected=filename_projects[0]
        d['task_name']=detected
        customer,_task=project_catalog.split_customer_task(detected)
        if customer and not str(d.get('customer') or '').strip():
            d['customer']=customer
    try:blocks=_shape_blocks(path)
    except Exception:blocks=[]
    raw='\n'.join(blocks)
    ratio=english_content_ratio(raw)
    d['_english_ratio']=ratio
    d['_english_mode']=ratio>=80.0

    # Image policy is shared by Korean and English:
    #   - only page 1 of the source 8D is eligible,
    #   - each picture belongs only to the D region that physically contains it,
    #   - Issue DB representative image is built only from page-1 2D pictures.
    try:
        page1_images=_english_section_images(path)
        d['_section_images']=page1_images
        rep=v310.composite(page1_images.get('2D',[]))
        d['_images']=[(1,1,1,rep)] if rep else []
    except Exception:
        d['_section_images']={k:[] for k in ('2D','3D','4D_CAUSE','4D_LEAK','5D','6D')}
        d['_images']=[]

    # Under 80% English: keep the existing Korean text extraction and add only
    # extra 4D table items (e.g. 재현시험/추가 검토) below 발생원인.
    if not d['_english_mode']:
        return _extract_cache_put(path,_augment_korean_4d_extras(path,d))

    # 80%+ English extraction now mirrors the Korean table logic:
    #   1) Matching table/header label inside the correct D region.
    #   2) Midpoint-based D-region geometry when the table has no usable label.
    #   3) Whole-document semantic fallback only when no D geometry is available.
    try:table_fields,table_detected=_scoped_table_semantic_fields(path)
    except Exception:table_fields,table_detected={},set()
    try:spatial,geo_marked=_spatial_section_blocks(path,True)
    except Exception:spatial,geo_marked=[],set()
    geo_fields,geo_detected=_english_fields_from_spatial(spatial) if spatial else ({},set())

    all_keys=('problem','temporary_action','cause_4d','leak_cause','system_cause','action_5d','verification_6d')
    for key in all_keys:
        if key in table_detected and table_fields.get(key,''):
            d[key]=table_fields[key]
        elif key in geo_detected and geo_fields.get(key,''):
            d[key]=geo_fields[key]

    if not spatial:
        source_fields,source_detected=_semantic_fields_and_detected(blocks)
        for key in all_keys:
            if key not in table_detected and key in source_detected:
                d[key]=source_fields.get(key,'')

    result=enhance_dict(d,raw,allow_label_fallback=False)
    result['_english_ratio']=ratio
    result['_english_mode']=True
    return _extract_cache_put(path,result)

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

def _ask_translation_choice(self,d,path,action='preview'):
    """Ask separately for preview or update; keep button labels short for the dialog width."""
    if not d.get('_english_mode'):
        return False
    pct=translation_coverage(d)
    is_update=(action=='update')
    verb='업데이트' if is_update else '미리보기'
    msg=(
        '영문 8D를 감지했습니다.\n\n'
        f'현재 자동 한글 번역률은 약 {pct}%입니다.\n'
        f'이번 {verb}에 사용할 언어를 선택해 주세요.\n\n'
        '선택한 버튼의 내용으로만 처리됩니다.'
    )
    buttons=(
        ('영문 원문',False),
        ('한글 번역',True),
    )
    choice=ui.dialog(
        self,f'영문 8D {verb} 선택',msg,'question',
        buttons,width=600,height=320
    )
    if choice is None:
        return None
    use_original=not bool(choice)
    self._english_translation_percent=pct
    return use_original

# Korean/English 6D status terms use one decision engine for Issue DB and weekly Signal.
_original_status=v310.base.status

def _status_decision(text):
    """Use the exact same Korean/English completion engine as weekly Signal."""
    raw=str(text or '').strip()
    if not raw:
        return 'open','6D 내용 없음'
    # main_enterprise is already loaded through main_enterprise_v3.
    import main_enterprise as ent
    state,reason=ent.weekly_verification_state(raw)
    return ('close',reason) if state=='complete' else ('open',reason)

def status_bilingual(d):
    decision,_reason=_status_decision(d.get('verification_6d'))
    if decision=='close':
        return v310.impl.base.STATUS['complete'] if hasattr(v310.impl.base,'STATUS') else '개선 완료'
    if str(d.get('action_5d') or '').strip() or str(d.get('verification_6d') or '').strip():
        return v310.impl.base.STATUS['verify'] if hasattr(v310.impl.base,'STATUS') else '개선 검증중'
    return _original_status(d)
v310.base.status=status_bilingual

def judge_status_bilingual(d):
    """Use the same 5D/6D bilingual recommendation as the enterprise status dialog."""
    import main_enterprise as ent
    decision,_reason=ent.issue_db_recommended_status(d)
    return decision,str(d.get('verification_6d') or '').strip()

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
        use_original=_ask_translation_choice(self,d,path,'preview')
        if use_original is None:
            self.status_var.set('READY · 미리보기가 취소되었습니다.')
            return
        d=apply_english_choice(d,use_original)
        self._last_preview_path=path
        self.log.delete('1.0','end')
        self.log.insert('end','[ 8D 추출 완료 ]\n'+'─'*72+'\n')
        self.log.insert('end',f'원본 파일       : {Path(path).name}\n')
        if d.get('_english_mode'):
            pct=getattr(self,'_english_translation_percent',translation_coverage(d))
            mode='영문 원문' if use_original else '한글 번역'
            self.log.insert('end',f'미리보기 언어   : {mode} · 자동 한글 번역률 약 {pct}%\n')
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
    use_original=None
    if probe.get('_english_mode'):
        # Always ask again for update, even if the user already previewed the same file.
        use_original=_ask_translation_choice(self,probe,path,'update')
        if use_original is None:
            self.status_var.set('READY · 업데이트가 취소되었습니다.')
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
