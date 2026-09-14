# 8D Issue Automation v3.0.2
# Adds collision-aware page-2 layout on top of v3.0.1.
# If one section needs more vertical space, the following section is moved down
# and its available height is reduced. The overflow is propagated in sequence.

import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import main_v301 as impl
base = impl.base

# Original example geometry is retained as the anchor layout.
# Each column is flowed independently:
# left: 2D -> 3D -> 4D 발생원인
# right: 4D 유출/시스템 -> 5D -> 6D

FLOW = {
    "left": ["2D", "3D", "4D_CAUSE"],
    "right": ["4D_LEAK", "5D", "6D"],
}

MIN_H = {
    "2D": 0.62,
    "3D": 0.70,
    "4D_CAUSE": 0.90,
    "4D_LEAK": 0.58,
    "5D": 0.82,
    "6D": 0.60,
}

# Bottom limits are taken from the original example slide geometry.
BOTTOM = {
    "left": 7.17,
    "right": 6.85,
}


def _text_for(key, d):
    if key == "2D":
        return str(d.get("problem") or "검토 중")
    if key == "3D":
        return "\n".join(x for x in (d.get("temporary_action"), d.get("customer_response")) if str(x or "").strip()) or "검토 중"
    if key == "4D_CAUSE":
        return str(d.get("cause_4d") or "검토 중")
    if key == "4D_LEAK":
        vals = [d.get("leak_cause"), d.get("system_cause")]
        return "\n".join(str(x) for x in vals if str(x or "").strip()) or "검토 중"
    if key == "5D":
        return str(d.get("action_5d") or "검토 중")
    if key == "6D":
        return str(d.get("verification_6d") or "검토 중")
    return ""


def _label_for(key):
    return {
        "2D": "2D 문제 현상:",
        "3D": "3D 임시 조치/고객 대응:",
        "4D_CAUSE": "4D 발생 원인:",
        "4D_LEAK": "4D 유출/시스템 원인:",
        "5D": "5D 개선 대책:",
        "6D": "6D 유효성 검증:",
    }[key]


def _estimate_need(key, text, width, original_h, has_images):
    # Conservative estimate: Korean/ASCII mixed text is wrapped roughly by the
    # available width. Images reserve about one third of the width.
    width_for_text = max(0.65, width * (0.64 if has_images else 0.94))
    chars_per_line = max(7, int(width_for_text * 9.2))
    lines = 0
    for line in str(text or "").splitlines() or [""]:
        lines += max(1, (len(line) + chars_per_line - 1) // chars_per_line)
    # Label + body line height. Extra margin gives PPT rendering a little safety.
    need = 0.20 + lines * 0.145 + 0.06
    return max(original_h, need)


def _flow_layout(layout, d, section_images):
    out = {k: dict(v) for k, v in layout.items()}

    for column, keys in FLOW.items():
        # Preserve the original x/w and first y. The first box remains anchored;
        # every later box starts after the actual used bottom of the previous box.
        start_y = min(out[k]["y"] for k in keys)
        bottom = BOTTOM[column]
        cursor = start_y

        # Calculate desired heights first.
        desired = {}
        for key in keys:
            desired[key] = _estimate_need(
                key,
                _text_for(key, d),
                out[key]["w"],
                out[key]["h"],
                bool(section_images.get("2D" if key == "2D" else "3D" if key == "3D" else "4D" if key.startswith("4D") else key, [])),
            )

        # Sequential collision propagation. When a box grows, the next box moves;
        # when there is insufficient room, the next box is reduced. Any remaining
        # deficit is propagated further downstream instead of causing overlap.
        remaining_keys = list(keys)
        for i, key in enumerate(keys):
            original_y = out[key]["y"]
            h = desired[key]
            out[key]["y"] = cursor

            space_after = bottom - cursor
            min_h = MIN_H[key]
            remaining_min = sum(MIN_H[k] for k in keys[i + 1:])
            available_for_current = max(min_h, space_after - remaining_min)
            actual_h = min(h, available_for_current)
            out[key]["h"] = max(min_h, actual_h)
            cursor = out[key]["y"] + out[key]["h"]

            # If current content cannot fit even at minimum height, it will be
            # auto-fitted by the renderer; importantly, it never overlaps the next zone.

    return out


def _dynamic_fill_page2(sl, d, g, layout):
    # Reuse v3.0.1's single renderer, but feed it a collision-aware layout.
    # No second renderer is called, so boxes cannot stack on top of one another.
    section_images = d.get("_section_images", {}) or {}
    dynamic = _flow_layout(layout, d, section_images)
    impl._fill_page2(sl, d, g, dynamic)


# Replace only the page-2 renderer used by v3.0.1's weekly function.
impl._fill_page2 = _dynamic_fill_page2

if __name__ == "__main__":
    base.App().mainloop()
