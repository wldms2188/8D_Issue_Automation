# 8D Issue Automation v3.1.0
# Clean weekly renderer: stable v2.9.5 extraction + direct section rendering.
# No PowerPoint XML grouping/ungrouping, no global font rewrite, no nested v3.x renderer patching.
import sys, re, io, copy, datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# Keep the known-good data extraction / Excel / GUI foundation.
import main_v295 as impl
base = impl.base

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from PIL import Image as PILImage, ImageOps, ImageDraw

EMU = 914400
SECTIONS = ('2D', '3D', '4D', '5D', '6D')

# Reference layout from the supplied weekly-meeting example.
ZONES = {
    '2D': {'x': .53, 'y': 2.53, 'w': 4.76, 'h': .98},
    '3D': {'x': .53, 'y': 3.91, 'w': 4.76, 'h': 1.14},
    '4D_CAUSE': {'x': .53, 'y': 5.38, 'w': 4.76, 'h': 1.79},
    '4D_LEAK': {'x': 5.77, 'y': 2.47, 'w': 4.76, 'h': .87},
    '5D': {'x': 5.77, 'y': 3.74, 'w': 4.76, 'h': 1.77},
    '6D': {'x': 5.77, 'y': 5.91, 'w': 4.78, 'h': .94},
}

NEW_MARKERS = {
    '2D': (.35, 2.18),
    '3D': (.35, 3.53),
    '4D': (5.59, 2.18),
    '5D': (5.59, 3.45),
    '6D': (5.59, 5.60),
}


def N(x):
    return str(x or '').replace('\r\n', '\n').replace('\r', '\n').strip()


def C(x):
    try:
        return base.compact(x)
    except Exception:
        return re.sub(r'[^0-9A-Za-z가-힣]', '', N(x)).lower()


def set_text(shape, text, size=8):
    if not hasattr(shape, 'text_frame'):
        return
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(.05)
    tf.margin_right = Inches(.05)
    tf.margin_top = Inches(.03)
    tf.margin_bottom = Inches(.03)
    p = tf.paragraphs[0]
    p.text = N(text)
    for run in p.runs:
        run.font.size = Pt(size)
        run.font.name = '맑은 고딕'
        try:
            r = run._r.get_or_add_rPr()
            r.set('a:latin', '맑은 고딕')
            r.set('a:ea', '맑은 고딕')
            r.set('a:cs', '맑은 고딕')
        except Exception:
            pass


def box(sh):
    return (float(sh.left), float(sh.top), float(sh.width), float(sh.height))


def overlap(a, b):
    ax, ay, aw, ah = a; bx, by, bw, bh = b
    return max(0, min(ax + aw, bx + bw) - max(ax, bx)) * max(0, min(ay + ah, by + bh) - max(ay, by))


def walk(container):
    for sh in getattr(container, 'shapes', []):
        yield sh
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.GROUP:
            yield from walk(sh)


def is_new(mode):
    q = C(mode)
    return any(k in q for k in ('신규', 'new', '신규이슈'))


def extract_task_exact(path, d):
    """Only use the requested customer-prefix rule; never use an arbitrary next cell."""
    customer = N(d.get('customer'))
    issue = N(d.get('issue_name'))
    if not customer:
        return d
    prs = Presentation(path)
    texts = []
    for sl in prs.slides:
        for sh in walk(sl):
            t = N(getattr(sh, 'text', ''))
            if t:
                texts.append(t)
            if getattr(sh, 'has_table', False):
                tb = sh.table
                for r in range(len(tb.rows)):
                    for c in range(len(tb.columns)):
                        t = N(tb.cell(r, c).text)
                        if t:
                            texts.append(t)
    pat = re.compile(r'(?<![\w가-힣])' + re.escape(customer) + r'\s*[_\-/／|:]\s*([^\n\r,;|]+)', re.I)
    for t in texts + [issue]:
        m = pat.search(t)
        if not m:
            continue
        value = N(m.group(1)).strip(' _-/／|:：')
        if value and C(value) not in ('상세시험조건', 'model', 'packer'):
            d['task_name'] = value
            return d
    return d


_original_extract = base.extract


def extract(path):
    d = _original_extract(path)
    try:
        d = extract_task_exact(path, d)
    except Exception:
        pass
    try:
        d['_section_images'] = collect_section_images(path)
    except Exception:
        d['_section_images'] = {k: [] for k in SECTIONS}
    return d


base.extract = extract


def section_from_text(text):
    q = C(text)
    if re.search(r'2d', q) or any(k in q for k in ('문제현상', '문제현황', '불량현상')):
        return '2D'
    if re.search(r'3d', q) or any(k in q for k in ('임시조치', '임시대책', '고객대응')):
        return '3D'
    if re.search(r'4d', q) or any(k in q for k in ('원인분석', '발생원인', '유출원인', '시스템원인')):
        return '4D'
    if re.search(r'5d', q) or any(k in q for k in ('개선대책', '개선사항', '영구개선')):
        return '5D'
    if re.search(r'6d', q) or any(k in q for k in ('효과검증', '유효성검증', '효과성검증')):
        return '6D'
    return None


def image_blob(sh):
    try:
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.PICTURE:
            return sh.image.blob
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.GROUP:
            pics = [p for p in walk(sh) if getattr(p, 'shape_type', None) == MSO_SHAPE_TYPE.PICTURE]
            if not pics:
                return None
            minx = min(box(p)[0] for p in pics); miny = min(box(p)[1] for p in pics)
            maxx = max(box(p)[0] + box(p)[2] for p in pics); maxy = max(box(p)[1] + box(p)[3] for p in pics)
            W = max(200, int((maxx - minx) / EMU * 140)); H = max(120, int((maxy - miny) / EMU * 140))
            canvas = PILImage.new('RGB', (W, H), 'white')
            for p in pics:
                with PILImage.open(io.BytesIO(p.image.blob)) as im:
                    im = im.convert('RGB')
                    px, py, pw, ph = box(p)
                    w = max(1, int(pw / EMU * 140)); h = max(1, int(ph / EMU * 140))
                    im.thumbnail((w, h), PILImage.Resampling.LANCZOS)
                    canvas.paste(im, (max(0, int((px - minx) / EMU * 140)), max(0, int((py - miny) / EMU * 140))))
            b = io.BytesIO(); canvas.save(b, 'PNG'); return b.getvalue()
    except Exception:
        return None
    return None


def collect_section_images(path):
    """Associate pictures with the nearest explicit 2D~6D heading on the same source slide."""
    prs = Presentation(path)
    out = {k: [] for k in SECTIONS}
    for si, sl in enumerate(prs.slides):
        anchors = []
        for sh in walk(sl):
            sec = section_from_text(N(getattr(sh, 'text', '')))
            if sec:
                anchors.append((sec, box(sh)))
        if not anchors:
            continue
        for sh in sl.shapes:
            if getattr(sh, 'shape_type', None) not in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP):
                continue
            blob = image_blob(sh)
            if not blob:
                continue
            ib = box(sh); ranked = []
            for sec, ab in anchors:
                ov = overlap(ib, ab)
                cx = ib[0] + ib[2] / 2; cy = ib[1] + ib[3] / 2
                ax = ab[0] + ab[2] / 2; ay = ab[1] + ab[3] / 2
                dist = ((cx - ax) ** 2 + (cy - ay) ** 2) ** .5
                ranked.append((1 if ov > 0 else 0, ov, -dist, sec))
            ranked.sort(reverse=True)
            out[ranked[0][3]].append((blob, ib, si))
    for sec, vals in out.items():
        seen = set(); unique = []
        for item in sorted(vals, key=lambda x: x[1][2] * x[1][3], reverse=True):
            h = hash(item[0])
            if h in seen:
                continue
            seen.add(h); unique.append(item)
        out[sec] = unique[:4]
    return out


def composite(items, cols=2):
    if not items:
        return None
    ims = []
    for blob, _, _ in items:
        try:
            with PILImage.open(io.BytesIO(blob)) as im:
                im = im.convert('RGB'); im.thumbnail((620, 360), PILImage.Resampling.LANCZOS); ims.append(im.copy())
        except Exception:
            pass
    if not ims:
        return None
    cellw, cellh = 620, 390
    rows = (len(ims) + cols - 1) // cols
    canvas = PILImage.new('RGB', (cols * cellw, rows * cellh), 'white')
    draw = ImageDraw.Draw(canvas)
    for i, im in enumerate(ims):
        x = (i % cols) * cellw; y = (i // cols) * cellh
        fitted = ImageOps.contain(im, (cellw - 20, cellh - 30))
        canvas.paste(fitted, (x + (cellw - fitted.width) // 2, y + 10))
        draw.rectangle((x, y, x + cellw - 1, y + cellh - 1), outline=(190, 190, 190), width=2)
    b = io.BytesIO(); canvas.save(b, 'PNG'); return b.getvalue()


def find_marker(sl, label):
    for sh in sl.shapes:
        if N(getattr(sh, 'text', '')) == label:
            return sh, None
        if getattr(sh, 'shape_type', None) == MSO_SHAPE_TYPE.GROUP:
            for child in walk(sh):
                if child is sh:
                    continue
                if N(getattr(child, 'text', '')) == label:
                    return child, sh
    return None, None


def nearby_title(sl, marker, label):
    """Find an ungrouped title near the marker; never create a PowerPoint group."""
    mx, my, mw, mh = box(marker)
    best = None; score = 1e30
    for sh in sl.shapes:
        if sh is marker or not hasattr(sh, 'text_frame'):
            continue
        t = N(getattr(sh, 'text', ''))
        if not t or t == label or len(t) > 80:
            continue
        sx, sy, sw, shh = box(sh)
        if abs(sy - my) < 1.0 * EMU and sx >= mx - 1.0 * EMU and sx <= mx + 3.0 * EMU:
            score0 = abs(sy - my) + abs(sx - (mx + mw))
            if score0 < score:
                best, score = sh, score0
    return best


def move_marker_unit(sl, label, x, y):
    marker, parent = find_marker(sl, label)
    if marker is None:
        return
    if parent is not None:
        # Existing group: move the existing group, do not construct XML.
        bx, by, _, _ = box(parent)
        parent.left += Inches(x) - int(bx)
        parent.top += Inches(y) - int(by)
        return
    # Ungrouped template: move marker and a nearby title independently to the same delta.
    mx, my, _, _ = box(marker)
    dx = Inches(x) - int(mx); dy = Inches(y) - int(my)
    title = nearby_title(sl, marker, label)
    marker.left += dx; marker.top += dy
    if title is not None:
        title.left += dx; title.top += dy


def add_box(sl, name, x, y, w, h, text, size=8):
    sh = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    sh.name = 'AUTO_8D_TEXT_' + name
    set_text(sh, text, size)
    return sh


def add_image(sl, blob, x, y, w, h, name):
    try:
        with PILImage.open(io.BytesIO(blob)) as im:
            iw, ih = im.size
        scale = min((w * EMU) / iw, (h * EMU) / ih)
        ww = iw * scale / EMU; hh = ih * scale / EMU
        sh = sl.shapes.add_picture(io.BytesIO(blob), Inches(x + (w - ww) / 2), Inches(y + (h - hh) / 2), width=Inches(ww), height=Inches(hh))
        sh.name = 'AUTO_8D_IMG_' + name
    except Exception:
        pass


def remove_previous_auto(sl):
    # Only remove shapes created by this program. Template graphics/logos are untouched.
    for sh in list(sl.shapes):
        if str(getattr(sh, 'name', '')).startswith('AUTO_8D_'):
            try:
                sh._element.getparent().remove(sh._element)
            except Exception:
                pass


def section_text(d, key):
    vals = {
        '2D': [d.get('problem')],
        '3D': [d.get('temporary_action'), d.get('customer_response')],
        '4D_CAUSE': [d.get('cause_4d')],
        '4D_LEAK': [d.get('leak_cause'), d.get('system_cause')],
        '5D': [d.get('action_5d')],
        '6D': [d.get('verification_6d')],
    }
    text = '\n'.join(N(v) for v in vals[key] if N(v))
    return text or '검토 중'


def render_zone(sl, key, text, images):
    z = ZONES[key]
    # Text occupies most of the zone; pictures are one controlled collage, never scattered.
    blob = composite(images)
    if blob:
        iw = min(1.65, z['w'] * .36)
        tw = z['w'] - iw - .08
        add_box(sl, key, z['x'], z['y'], tw, z['h'], text, 8)
        add_image(sl, blob, z['x'] + tw + .08, z['y'], iw, z['h'], key)
    else:
        add_box(sl, key, z['x'], z['y'], z['w'], z['h'], text, 8)


def update_page1(prs, d, g):
    team = N(g.get('team'))
    values = {
        '과제명': N(d.get('task_name')),
        '이슈': N(d.get('issue_name')),
        '이슈명': N(d.get('issue_name')),
        '문제/현상': N(d.get('problem')),
    }
    try:
        progress = impl.v305.v303.v301.impl.v29.v29_progress(d)
    except Exception:
        progress = ''
    values['진행'] = N(progress)

    for si in range(min(1,len(prs.slides))):
        sl=prs.slides[si]
        for sh in walk(sl):
            if getattr(sh, 'has_table', False):
                tb = sh.table
                hm = {}
                for c in range(len(tb.columns)):
                    q = C(tb.cell(0, c).text)
                    if '과제명' in q: hm['과제명'] = c
                    elif q == '이슈' or '이슈명' in q: hm['이슈'] = c
                    elif '문제' in q or '현상' in q: hm['문제/현상'] = c
                    elif '진행' in q: hm['진행'] = c
                    elif 'signal' in q: hm['Signal'] = c
                if '이슈' not in hm:
                    continue
                row = 1
                for r in range(1, len(tb.rows)):
                    if not any(N(tb.cell(r, c).text) for c in range(len(tb.columns))):
                        row = r; break
                for key, c in hm.items():
                    if key in values:
                        # Preserve table geometry/format; only replace the requested cell text.
                        tb.cell(row, c).text = values[key]
                        for p in tb.cell(row, c).text_frame.paragraphs:
                            for run in p.runs:
                                run.font.size = Pt(8)
                if 'Signal' in hm:
                    try: base.signal(tb.cell(row, hm['Signal']), base.status(d))
                    except Exception: pass
                break

            t = N(getattr(sh, 'text', ''))
            if team and '팀 주요 논의 사항' in t:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if '팀 주요 논의 사항' in r.text:
                            r.text = re.sub(r'[^\s]*팀\s*주요\s*논의\s*사항', team + '팀 주요 논의 사항', r.text)
                            break


def update_page2(sl, d, g, mode):
    remove_previous_auto(sl)
    if is_new(mode):
        for label, (x, y) in NEW_MARKERS.items():
            move_marker_unit(sl, label, x, y)

    imgs = d.get('_section_images', {}) or {}
    render_zone(sl, '2D', section_text(d, '2D'), imgs.get('2D', []))
    render_zone(sl, '3D', section_text(d, '3D'), imgs.get('3D', []))
    render_zone(sl, '4D_CAUSE', section_text(d, '4D_CAUSE'), imgs.get('4D', []))
    render_zone(sl, '4D_LEAK', section_text(d, '4D_LEAK'), [])
    render_zone(sl, '5D', section_text(d, '5D'), imgs.get('5D', []))
    render_zone(sl, '6D', section_text(d, '6D'), imgs.get('6D', []))

    team = N(g.get('team')); owner = N(g.get('owner'))
    stage = N(g.get('stage') or g.get('occurrence_stage') or d.get('stage'))
    date = N(d.get('occurrence_date') or g.get('occurrence_date'))
    issue = N(d.get('issue_name')); task = N(d.get('task_name'))

    for sh in walk(sl):
        if hasattr(sh, 'text_frame'):
            t = N(getattr(sh, 'text', ''))
            if t.startswith('이슈명'):
                prefix = t.split(':', 1)[0] if ':' in t else '이슈명'
                set_text(sh, prefix + ' : ' + issue, 8)
            elif '과제명_이슈 제목' in t:
                set_text(sh, t.replace('과제명_이슈 제목', task), 8)
            elif '00팀 담당자' in t:
                set_text(sh, f'{team} 담당자 : {owner}', 8)
            elif '발생단계' in C(t) and '발생일자' in C(t):
                # If the value is a separate text box next to the label, fill only a blank nearby box.
                pass
        if getattr(sh, 'has_table', False):
            tb = sh.table
            for r in range(len(tb.rows)):
                for c in range(len(tb.columns)):
                    q = C(tb.cell(r, c).text)
                    if 'signal' == q and c + 1 < len(tb.columns):
                        try: base.signal(tb.cell(r, c + 1), base.status(d))
                        except Exception: pass
                    if '발생단계' in q and c + 1 < len(tb.columns):
                        tb.cell(r, c + 1).text = f'{stage} ({date})'.strip(' ()')

    # Explicit occurrence value: use a text box adjacent to the label when present.
    target = f'{stage} ({date})'.strip(' ()')
    labels = []
    for sh in sl.shapes:
        t = C(getattr(sh, 'text', ''))
        if '발생단계' in t and '발생일자' in t:
            labels.append(sh)
    for lab in labels:
        lx, ly, lw, lh = box(lab)
        best = None; score = 1e30
        for sh in sl.shapes:
            if sh is lab or not hasattr(sh, 'text_frame'):
                continue
            t = N(getattr(sh, 'text', ''))
            sx, sy, sw, shh = box(sh)
            if abs(sy - ly) < .5 * EMU and sx > lx:
                s = sx - (lx + lw)
                if 0 <= s < score:
                    best, score = sh, s
        if best is not None:
            set_text(best, target, 9)
            break


def weekly(src, out, d, g, mode):
    prs = Presentation(src)
    update_page1(prs, d, g)
    if len(prs.slides) > 1:
        # Exactly one detail page is updated. Source 8D slides never become extra weekly pages.
        update_page2(prs.slides[1], d, g, mode)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(out); saved = out
    except PermissionError:
        p = Path(out)
        saved = str(p.with_name(p.stem + '_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + p.suffix))
        prs.save(saved)
    return '주간회의 PPT 업데이트: ' + base.status(d), saved


base.weekly = weekly

if __name__ == '__main__':
    base.App().mainloop()
