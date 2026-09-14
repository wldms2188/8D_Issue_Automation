# 8D Issue Automation v2.9.1 hotfix
# Fixes the v2.9 detail-slide callback recursion without changing its behavior.

import main_v29 as v29
base = v29.base


def fixed_fill_detail_slide(sl, d, g):
    # Use the original v2.8 detail logic directly, then v2.9 geometry/text fixes.
    base.replace_metadata_shapes(sl, d, g)
    base.fill_numbered_boxes(sl, d)
    used = set()
    for tb, _, _, _ in base.table_cells(sl):
        base.fill_metadata_table(tb, d, g)
        base.fill_semantic_table(tb, d, used)
    # Weekly-style blank boxes when numbered annotations are not present.
    v29._fill_weekly_geometry(sl, d)
    # Enforce requested font on all populated text.
    for sh in base.flat(sl):
        if hasattr(sh, 'text_frame'):
            for p in sh.text_frame.paragraphs:
                for run in p.runs:
                    v29.v29_font_run(run, 8)

# v2.9 weekly() resolves this function by name from the imported module.
v29.v29_fill_detail_slide = fixed_fill_detail_slide
base.fill_detail_slide = fixed_fill_detail_slide

if __name__ == '__main__':
    base.App().mainloop()
