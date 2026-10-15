# Review — mnda (editing: form filling + highlight removal + font fidelity)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true — 2 output pages rendered
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true — fresh `docx read` is byte-identical to the output read view

## Merits
- All cover-page placeholders filled with the task's values: Purpose, Effective Date (June 8, 2026), MNDA Term `2 year(s)`, Term of Confidentiality `3 year(s)` from `Effective Date`, Delaware, New Castle County, Delaware, `None.`. No `[ ... ]` placeholders remain; the only brackets left are the original `[ x ]` / `[    ]` checkboxes.
- Zero `<w:highlight>` in `word/document.xml`; `docx find --highlight any` → no matches.
- Filled runs carry the placeholder's exact rPr (Arial, sz 16, color text1, `w:u single`) minus the highlight, so they look like the original blanks.
- Signature table t2: Print Name / Title / Company / Notice Address / Date for both parties in `c1` and `c3` of `r2`–`r6`, each a single `p0`. Label column, spacer `c2` and signature row `r1` are untouched. Both Date cells are filled.
- Signature values are Arial 9pt, taken from each empty cell's own paragraph-mark rPr (`<w:pPr><w:rPr>… sz 18`), which matches the 9pt labels. This is correct behavior, not drift.
- styles/numbering/settings/theme are unchanged in content (only the XML declaration's line ending differs). The Georgia title, bold labels, gray 7pt sub-labels and all of page 2 render the same as the baseline.

## Demerits
- **minor (tool):** A span `--text` edit combined with `clear: highlight` dropped the placeholder's single underline. The agent caught it and restored it with a second `underline: true` batch, so the final document is correct. This is still a fidelity bug in the tool, and a weaker run could ship the lost underline.
- **minor (UX):** `find "Effective Date"` returned 5 matches, so the agent picked the highlighted span by hand. Spans created by a batch also have to be found again before a follow-up formatting batch.

## Before/after comparison
- **Text:** the diff of the baseline and output read views changes only the t1 value cells (r0–r6 c1) and the t2 c1/c3 cells (r2–r6). Nothing else changed.
- **Visual:** page 1 is identical apart from the filled values and the removed yellow highlight. The Term of Confidentiality row reflows to one line fewer because the value is shorter, which is expected. Page 2 is pixel-identical in content.

## Frictions
- Agent-reported: lost underline after `clear: highlight` (major for the agent), 9pt size in bare cells not documented (minor; the behavior is correct), span locators shift within a paragraph, and the `read --ast --from` usage dump.
- **My read:** the agent got a correct, clean result. It took a moderate slog: one corrective batch and some manual span disambiguation. The underline loss is the real tool issue to fix.
