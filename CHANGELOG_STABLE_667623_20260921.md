# Stable baseline patch note — 2026-09-21

## Baseline
- Source commit: `667623cc0c7493be88c0c88f8a96352b823c0fa7`
- Working branch: `repair/user-stable-667623-five-fixes`
- Reason: return to the user-confirmed comparatively stable weekly-meeting behavior before the later chain of fixes/regressions.

## Scope: only these five weekly-meeting changes
1. Detail title format
   - `고객사_과제명_발생샘플 샘플_개발단계_발생처 이슈 발생`
   - GUI/input values take priority; baseline fallback remains when a required component cannot be resolved.

2. Summary continuation
   - Search all summary pages.
   - Continue from the LAST page containing the confirmed task.
   - If one row fits, insert directly below the last task row.
   - If it does not fit, clone a summary page immediately after that last matching page.

3. 7D placement
   - Reposition the native 7D circle + 수평전개 unit below the dynamically calculated 6D region.
   - Absolute positioning is recalculated each run to avoid cumulative drift.

4. Summary task-name matching priority
   - (1) exact customer + project
   - (2) exact project
   - (3) only for a similar PowerPoint section that the user chose via “해당 구역 업데이트”, compare that selected section name with actual summary task labels.
   - If no actual summary label matches, do not force fuzzy matching; use the normal new-summary path and write the confirmed customer/project name.
   - When an existing summary label is used, preserve that label verbatim.

5. Clone cleanup
   - On newly cloned summary/detail pages, remove copied picture/media/chart/OLE visual objects.
   - Preserve tables, text, rectangles/frames, lines and other template structure.
   - Existing stable source-8D attachment insertion remains unchanged.

## Explicitly unchanged
- Issue DB / Excel logic
- 8D attachment insertion/deduplication
- Native PowerPoint section creation implementation from the baseline
- Existing extraction logic
- Existing output versioning
- Existing table/frame styles

## Added files
- `app/user_stable_weekly_fixes.py`
- `tests/test_user_stable_five_fixes.py`

The patch is loaded last from `app/main_enterprise_final.py` so it is isolated and can be removed without rewriting the stable core.
