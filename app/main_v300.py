# 8D Issue Automation v3.0
# Weekly page-2 rebuild:
# - 2D~6D source content is transferred in full, with related source images.
# - Uses the originally supplied example zones as the default geometry.
# - 4D is split into 발생 원인 / 유출·시스템 원인 exactly as the example mapping.
# - 4D labels are written immediately before the cause content.
# - Text auto-fits without truncation.
# - Layout is externally adjustable through 주간회의_레이아웃.json.
# - 과제명 is extracted from the issue-name metadata area, specifically the value
#   following 고객사, and is written to the issue DB through the existing task field.

import sys, io, json, re, copy, datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import main_v295 as impl
base = impl.base
v29 = impl.v29

# Save the stable v2.9.5 extractor before patching base.extract.
_STABLE_EXTRACT = base.extract

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Inches, Pt
from PIL import Image as PILImage

EMU = 914400

# ---------------------------------------------------------------------------
# Default geometry = the zones from the user's original weekly-meeting example.
# Coordinates are inches from the top-left of the slide.
# These can be changed without editing Python: see 주간회의_레이아웃.json.
# ---------------------------------------------------------------------------
DEFAULT_LAYOUT = {
    "2D": {"x": 0.53, "y": 2.53, "w": 4.76, "h": 0.98},
    "3D": {"x": 0.53, "y": 3.91, "w": 4.76, "h": 1.14},
    "4D_CAUSE": {"x": 0.53, "y": 5.38, "w": 4.76, "h": 1.79},
    "4D_LEAK": {"x": 5.77, "y": 2.47, "w": 4.76, "h": 0.87},
    "5D": {"x": 5.77, "y": 3.74, "w": 4.76, "h": 1.77},
    "6D": {"x": 5.77, "y": 5.91, "w": 4.78, "h": 0.94},
}

LAYOUT_FILE = APP_DIR.parent / "주간회의_레이아웃.json"


def _load_layout():
    data = None
    if LAYOUT_FILE.exists():
        try:
            data = json.loads(LAYOUT_FILE.read_text(encoding="utf-8"))
        except Exception:
            data = None
    if not isinstance(data, dict):
        data = copy.deepcopy(DEFAULT_LAYOUT)
        try:
            LAYOUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
    # Missing/invalid entries fall back to the authoritative defaults.
    out = copy.deepcopy(DEFAULT_LAYOUT)
    for k, v in data.items():
        if k not in out or not isinstance(v, dict):
            continue
        for p in ("x", "y", "w", "h"):
            try:
                out[k][p] = float(v[p])
            except Exception:
                pass
    return out


def _norm(s):
    return str(s or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def _compact(s):
    try:
        return base.compact(s)
    except Exception:
        return re.sub(r"[^0-9A-Za-z가-힣]", "", _norm(s)).lower()


def _box(sh):
    return float(sh.left), float(sh.top), float(sh.width), float(sh.height)


def _pics(sh):
    out = []
    typ = getattr(sh, "shape_type", None)
    if typ == MSO_SHAPE_TYPE.PICTURE:
        out.append(sh)
    elif typ == MSO_SHAPE_TYPE.GROUP:
        for c in sh.shapes:
            out.extend(_pics(c))
    return out


def _shape_image_blob(sh):
    pics = _pics(sh)
    if not pics:
        return None
    if len(pics) == 1:
        try:
            return pics[0].image.blob
        except Exception:
            return None
    gx, gy, gw, gh = _box(sh)
    if gw <= 0 or gh <= 0:
        return None
    W = max(80, int(gw / EMU * 160))
    H = max(60, int(gh / EMU * 160))
    canvas = PILImage.new("RGB", (W, H), "white")
    for p in pics:
        try:
            with PILImage.open(io.BytesIO(p.image.blob)) as im:
                im = im.convert("RGB")
                px, py, pw, ph = _box(p)
                x = max(0, int((px - gx) / EMU * 160))
                y = max(0, int((py - gy) / EMU * 160))
                w = max(1, int(pw / EMU * 160))
                h = max(1, int(ph / EMU * 160))
                im.thumbnail((w, h), PILImage.Resampling.LANCZOS)
                canvas.paste(im, (x, y))
        except Exception:
            pass
    b = io.BytesIO()
    canvas.save(b, "PNG")
    return b.getvalue()


def _image_candidates(prs):
    out = []
    for si, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            typ = getattr(sh, "shape_type", None)
            if typ == MSO_SHAPE_TYPE.GROUP:
                pics = _pics(sh)
                if not pics:
                    continue
                blob = _shape_image_blob(sh)
                if blob:
                    out.append((si, _box(sh), blob, len(pics)))
            elif typ == MSO_SHAPE_TYPE.PICTURE:
                try:
                    with PILImage.open(io.BytesIO(sh.image.blob)) as im:
                        if im.width * im.height >= 5000:
                            out.append((si, _box(sh), sh.image.blob, 1))
                except Exception:
                    pass
    return out


def _section_label(text):
    q = _compact(text)
    if re.search(r"2d", q) or any(x in q for x in ("문제현황", "문제현상", "불량현상")):
        return "2D"
    if re.search(r"3d", q) or any(x in q for x in ("임시조치", "임시대책", "고객대응")):
        return "3D"
    if re.search(r"4d", q) or any(x in q for x in ("원인분석", "발생원인", "유출원인", "시스템원인")):
        return "4D"
    if re.search(r"5d", q) or any(x in q for x in ("개선대책", "개선사항", "영구개선")):
        return "5D"
    if re.search(r"6d", q) or any(x in q for x in ("효과검증", "유효성검증", "효과성검증")):
        return "6D"
    return None


def _center(box):
    x, y, w, h = box
    return x + w / 2.0, y + h / 2.0


def _distance(a, b):
    ax, ay = _center(a)
    bx, by = _center(b)
    return ((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5


def _section_images(prs):
    """Assign every meaningful source image on the 8D PPT to the nearest 2D~6D label.

    A picture is considered relevant when it is on the same slide as a D label.
    If it is inside/near the labeled block it wins over a distant picture. Multiple
    images are retained, not just one representative image.
    """
    labels = []
    for si, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            txt = _norm(getattr(sh, "text", ""))
            sec = _section_label(txt)
            if sec:
                labels.append((si, sec, _box(sh)))

    imgs = _image_candidates(prs)
    result = {"2D": [], "3D": [], "4D": [], "5D": [], "6D": []}
    for si, ibox, blob, count in imgs:
        same = [(sec, lbox, _distance(ibox, lbox)) for ssi, sec, lbox in labels if ssi == si]
        if not same:
            continue
        # Strong preference for actual intersection; otherwise nearest section label.
        intersect = []
        for sec, lbox, dist in same:
            ix = max(0, min(ibox[0] + ibox[2], lbox[0] + lbox[2]) - max(ibox[0], lbox[0]))
            iy = max(0, min(ibox[1] + ibox[3], lbox[1] + lbox[3]) - max(ibox[1], lbox[1]))
            ov = ix * iy
            if ov > 0:
                intersect.append((ov, -dist, sec))
        if intersect:
            intersect.sort(reverse=True)
            sec = intersect[0][2]
        else:
            same.sort(key=lambda x: x[2])
            # Avoid assigning an image that is extremely far from every D marker.
            if same[0][2] > 6.5 * EMU:
                continue
            sec = same[0][0]
        result[sec].append((blob, ibox, count))

    # Remove exact duplicate blobs while retaining all distinct images.
    for sec, vals in result.items():
        seen = set(); clean = []
        for blob, ibox, count in vals:
            key = hash(blob)
            if key in seen:
                continue
            seen.add(key); clean.append((blob, ibox, count))
        result[sec] = clean
    return result


def _extract_task_from_issue_area(path, d):
    """Find 과제명 from the metadata area where Customer is followed by the task value.

    This intentionally does not use the generic '과제명' label as the primary source,
    because the supplied 8D format places the task value immediately after 고객사.
    """
    try:
        prs = Presentation(path)
    except Exception:
        return d.get("task_name", "")

    # 1) Table form: find a cell containing 고객사 and take the next non-label value
    # in the same row, then the next row/cell if necessary.
    for sl in prs.slides:
        for sh in sl.shapes:
            if not getattr(sh, "has_table", False):
                continue
            tb = sh.table
            rows = [[_norm(tb.cell(r, c).text) for c in range(len(tb.columns))] for r in range(len(tb.rows))]
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    if "고객사" not in _compact(val):
                        continue
                    candidates = []
                    # Same row, to the right: most common format.
                    candidates += row[c + 1:]
                    # Same column, following rows.
                    for rr in range(r + 1, min(r + 4, len(rows))):
                        candidates.append(rows[rr][c])
                    # Next cells in the following row(s).
                    for rr in range(r + 1, min(r + 3, len(rows))):
                        candidates += rows[rr][c + 1:]
                    for cand in candidates:
                        q = _compact(cand)
                        if not cand or "고객사" in q or "이슈명" in q or "과제명" in q:
                            continue
                        if q in {"model", "packer", "발생site", "발생일자", "lotno", "발생line"}:
                            continue
                        if len(cand) <= 1:
                            continue
                        # Do not mistake a generic metadata label for the task.
                        if any(x in q for x in ("담당자", "제품타입", "폼팩터", "발생샘플", "개발단계")):
                            continue
                        return cand.strip()

    # 2) Shape/text form: locate a line containing 고객사 and use the text immediately
    # after it. If the same text block contains several lines, the next line is the task.
    for sl in prs.slides:
        for sh in sl.shapes:
            text = _norm(getattr(sh, "text", ""))
            if "고객사" not in text:
                continue
            lines = [x.strip() for x in text.split("\n") if x.strip()]
            for i, line in enumerate(lines):
                if "고객사" not in _compact(line):
                    continue
                if ":" in line or "：" in line:
                    after = re.split(r"[:：]", line, maxsplit=1)[1].strip()
                    if after and not any(k in _compact(after) for k in ("model", "packer", "발생site")):
                        # If the customer value is followed by a separate task line,
                        # use that following line as the task when available.
                        if i + 1 < len(lines) and "과제" not in _compact(after):
                            return lines[i + 1]
                        return after
                if i + 1 < len(lines):
                    return lines[i + 1]
    return d.get("task_name", "")


def _post_extract(path):
    d = _STABLE_EXTRACT(path)
    try:
        task_name = _extract_task_from_issue_area(path, d)
        if task_name:
            d["task_name"] = task_name
    except Exception:
        pass
    try:
        prs = Presentation(path)
        secimgs = _section_images(prs)
        d["_section_images"] = secimgs
        # Representative image remains the 2D-first image for the Excel photo.
        rep = secimgs.get("2D") or secimgs.get("4D") or []
        if rep:
            d["_images"] = [(len(rep[0][0]), 1, 1, rep[0][0])]
        else:
            # Keep v2.9.5 representative image if section mapping found nothing.
            d.setdefault("_images", [])
    except Exception:
        d.setdefault("_section_images", {k: [] for k in ("2D", "3D", "4D", "5D", "6D")})
    return d

base.extract = _post_extract

# ---------------------------------------------------------------------------
# PowerPoint writing helpers
# ---------------------------------------------------------------------------
def _font(run, size, bold=False):
    run.font.name = "맑은 고딕"
    run.font.size = Pt(max(4.0, size))
    run.font.bold = bold
    try:
        r = run._r.get_or_add_rPr()
        r.set("a:latin", "맑은 고딕")
        r.set("a:ea", "맑은 고딕")
        r.set("a:cs", "맑은 고딕")
    except Exception:
        pass


def _write_text(sh, text, size=8, bold=False):
    tf = sh.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.text = _norm(text)
    for run in p.runs:
        _font(run, size, bold)
    return sh


def _fit_text(sh, text, max_size=8, min_size=5, prefix=None):
    text = _norm(text)
    if prefix and text:
        text = prefix + text
    _write_text(sh, text, max_size)
    if not text:
        return max_size
    width = max(0.5, sh.width / EMU)
    height = max(0.2, sh.height / EMU)
    size = float(max_size)
    lines = text.split("\n")
    while size > min_size:
        chars_per_line = max(8, int(width * 10.5 * (8.0 / size)))
        estimated = sum(max(1, int((len(line) + chars_per_line - 1) / chars_per_line)) for line in lines)
        line_height = size / 72.0 * 1.18
        if estimated * line_height <= height - 0.04:
            break
        size -= 0.5
    for p in sh.text_frame.paragraphs:
        for run in p.runs:
            _font(run, size)
    return size


def _delete_auto(sl):
    for sh in list(sl.shapes):
        if str(getattr(sh, "name", "")).startswith("AUTO_8D_"):
            sp = sh._element
            sp.getparent().remove(sp)


def _add_box(sl, key, x, y, w, h):
    sh = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    sh.name = "AUTO_8D_" + key
    return sh


def _add_image_fit(sl, blob, x, y, w, h, name):
    try:
        with PILImage.open(io.BytesIO(blob)) as im:
            iw, ih = im.size
        if iw <= 0 or ih <= 0:
            return None
        scale = min(w * EMU / iw, h * EMU / ih)
        # Keep small images from being enlarged excessively.
        scale = min(scale, 1.0)
        ww = max(0.05, iw * scale / EMU)
        hh = max(0.05, ih * scale / EMU)
        xx = x + max(0, (w - ww) / 2)
        yy = y + max(0, (h - hh) / 2)
        sh = sl.shapes.add_picture(io.BytesIO(blob), Inches(xx), Inches(yy), width=Inches(ww), height=Inches(hh))
        sh.name = "AUTO_8D_IMG_" + name
        return sh
    except Exception:
        return None


def _render_zone(sl, zone_key, layout, text, images, label=None):
    z = layout[zone_key]
    x, y, w, h = z["x"], z["y"], z["w"], z["h"]
    # The label is a separate small box immediately before the actual content.
    label_h = 0.20 if label else 0.0
    if label:
        lb = _add_box(sl, "LABEL_" + zone_key, x, y, min(0.75, w), label_h)
        _fit_text(lb, label, 7, 5)
        y += label_h
        h = max(0.15, h - label_h)

    images = images or []
    # Put text and image(s) in the same authoritative zone. If images exist, reserve
    # the right side for them; this prevents images from spilling into another D zone.
    image_w = 0.0
    if images:
        image_w = min(max(1.0, w * 0.34), 1.55)
        text_w = max(0.5, w - image_w - 0.05)
    else:
        text_w = w

    tx = _add_box(sl, "TEXT_" + zone_key, x, y, text_w, h)
    _fit_text(tx, text, 8, 4.5)

    if images:
        ix = x + text_w + 0.05
        each_h = max(0.15, (h - 0.03 * (len(images) - 1)) / max(1, len(images)))
        # At most three visible images per zone; if the source contains more, combine
        # them vertically into the available space rather than dropping the section.
        for i, item in enumerate(images[:3]):
            blob = item[0]
            _add_image_fit(sl, blob, ix, y + i * (each_h + 0.03), image_w, each_h, zone_key + "_" + str(i))
        if len(images) > 3:
            # Combine the remaining images into one composite thumbnail.
            ims = []
            for item in images[3:]:
                try:
                    with PILImage.open(io.BytesIO(item[0])) as im:
                        ims.append(im.convert("RGB"))
                except Exception:
                    pass
            if ims:
                cw = max(i.width for i in ims); ch = sum(i.height for i in ims)
                canvas = PILImage.new("RGB", (cw, ch), "white")
                yy = 0
                for im in ims:
                    canvas.paste(im, (0, yy)); yy += im.height
                b = io.BytesIO(); canvas.save(b, "PNG")
                _add_image_fit(sl, b.getvalue(), ix, y + max(0, h - each_h), image_w, each_h, zone_key + "_more")


def _combine_4d(d):
    return _norm(d.get("cause_4d")) or "검토 중"


def _combine_leak(d):
    vals = [d.get("leak_cause"), d.get("system_cause")]
    vals = [v for v in vals if _norm(v)]
    return "\n".join(vals) if vals else "검토 중"


def _fill_metadata(sl, d, g):
    title = _norm(d.get("task_name"))
    issue = _norm(d.get("issue_name"))
    for sh in sl.shapes:
        if not hasattr(sh, "text_frame"):
            continue
        txt = _norm(getattr(sh, "text", ""))
        new = txt
        if "과제명_이슈 제목" in new:
            new = new.replace("과제명_이슈 제목", (title + "_" + issue).strip("_"))
        if "이슈명" in new and ":" in new:
            new = new.split(":", 1)[0] + " : " + issue
        if "00팀 담당자 : 000" in new:
            new = new.replace("00팀 담당자 : 000", f"{g.get('team','')} 담당자 : {g.get('owner','')}")
        if new != txt:
            _fit_text(sh, new, 8, 5)


def _fill_page2(sl, d, g, layout):
    _delete_auto(sl)
    _fill_metadata(sl, d, g)
    imgs = d.get("_section_images", {})

    # Keep the template's original marker/arrow shapes. Only normalize their positions
    # when they are explicitly present; the content itself is rendered once below.
    marker_targets = {
        "2D": (0.35, 3.53),
        "3D": (0.35, 4.96),
        "4D": (5.59, 2.34),
        "5D": (5.59, 3.61),
        "6D": (5.59, 5.78),
    }
    for label, (mx, my) in marker_targets.items():
        for sh in sl.shapes:
            if _norm(getattr(sh, "text", "")) == label:
                sh.left = Inches(mx); sh.top = Inches(my)

    _render_zone(sl, "2D", layout, _norm(d.get("problem")) or "검토 중", imgs.get("2D"), "2D 문제 현상:")
    _render_zone(sl, "3D", layout, "\n".join(x for x in (d.get("temporary_action"), d.get("customer_response")) if _norm(x)) or "검토 중", imgs.get("3D"), "3D 임시 조치/고객 대응:")
    _render_zone(sl, "4D_CAUSE", layout, _combine_4d(d), imgs.get("4D"), "4D 발생 원인:")
    _render_zone(sl, "4D_LEAK", layout, _combine_leak(d), [], "4D 유출/시스템 원인:")
    _render_zone(sl, "5D", layout, _norm(d.get("action_5d")) or "검토 중", imgs.get("5D"), "5D 개선 대책:")
    _render_zone(sl, "6D", layout, _norm(d.get("verification_6d")) or "검토 중", imgs.get("6D"), "6D 유효성 검증:")

    # Signal cells on page 2, if present.
    for sh in sl.shapes:
        if not getattr(sh, "has_table", False):
            continue
        tb = sh.table
        for r in range(len(tb.rows)):
            for c in range(len(tb.columns)):
                if "signal" in _compact(tb.cell(r, c).text) and c + 1 < len(tb.columns):
                    base.signal(tb.cell(r, c + 1), base.status(d))


def weekly_v300(src, out, d, g, mode):
    prs = Presentation(src)
    layout = _load_layout()
    st = base.status(d)
    title = _norm(d.get("task_name"))
    issue = _norm(d.get("issue_name"))
    summary = None
    summary_idx = 0

    # Update the existing summary table, preserving its structure.
    for i, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if getattr(sh, "has_table", False):
                header = " ".join(_norm(sh.table.cell(0, c).text) for c in range(len(sh.table.columns)))
                if "과제명" in header and "Signal" in header:
                    summary = sh.table; summary_idx = i; break
        if summary:
            break

    if summary is not None:
        r = None
        for rr in range(1, len(summary.rows)):
            if title and title in _norm(summary.cell(rr, 0).text):
                r = rr; break
        if r is None:
            for rr in range(1, len(summary.rows)):
                if not any(_norm(summary.cell(rr, c).text) for c in range(min(5, len(summary.columns))) if c != 4):
                    r = rr; break
        if r is None:
            r = len(summary.rows) - 1
        vals = [base.task(d), issue, _norm(d.get("problem")), v29.v29_progress(d), "●"]
        for c, val in enumerate(vals):
            v29.v29_set_cell_text(summary.cell(r, c), val, 8)
        base.signal(summary.cell(r, 4), st)

    # Page 2 onward: render only the intended detail page(s). If the template has
    # multiple detail pages, all of them receive the same data without stacking
    # a second renderer on top of the first.
    details = [prs.slides[i] for i in range(summary_idx + 1, len(prs.slides))]
    if not details and len(prs.slides) > 1:
        details = [prs.slides[1]]
    for sl in details:
        _fill_page2(sl, d, g, layout)

    Path(out).parent.mkdir(exist_ok=True)
    try:
        prs.save(out); saved = out
    except PermissionError:
        p = Path(out)
        saved = p.with_name(p.stem + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + p.suffix)
        prs.save(saved)
    return "주간회의 PPT 업데이트: " + st, saved


base.weekly = weekly_v300

if __name__ == "__main__":
    base.App().mainloop()
