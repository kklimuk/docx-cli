# Review: mnda (Form filling + highlight removal + font fidelity) — Anthropic docx skill, r3

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (2 output pages rendered)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Merits
- All cover-page placeholders filled with the right values: Purpose, Effective Date June 8, 2026, MNDA Term 2 year(s), Confidentiality 3 year(s) from Effective Date, Delaware, New Castle County, Delaware, and "None." for changes.
- No `[ ... ]` brackets remain (grep of `word/document.xml` finds none). The `[ x ]` / `[    ]` checkboxes are not placeholders and are untouched.
- Zero `<w:highlight>` elements left in `word/document.xml`, and the render shows no yellow.
- Signature table t2 filled in exactly the intended cells: Party 1 in c1, Party 2 in c3, rows r2–r6 (Print Name, Title, Company, Notice Address, Date). Each one is a single `p0`. Label column c0, spacer c2 and signature row r1 are untouched. Both Date cells are filled.
- Fonts are preserved: the Georgia title, bold labels, gray 7pt sub-labels and Arial body all match the baseline. New signature values are 9pt Arial, matching the labels.
- Page 2 (Standard Terms) is visually identical to the baseline. The baseline-vs-output read.md diff shows only the intended value changes.
- `docx read` of the output shows every value, so the changes survive the write→read loop.

## Demerits
- **minor:** Filled cover-page values keep the blank's underline (for example `<u>Delaware</u>` and `<u>June 8, 2026</u>`). This is defensible because the task said to change only the values and the highlight, and the baseline blanks were underlined. The agent flagged this itself as a judgment call.

## Frictions
- The skill's SKILL.md has no form-fill guidance, so the agent wrote its own regex-based XML script.
- `merge_runs.py` failed under `python3 -I` with `ModuleNotFoundError: office`. The agent skipped that step.
- The agent's own assertion bug (`<w:r` also matched `<w:rPr>`) was caught before anything was written. The agent fixed it and reran.
- Read: a fairly smooth path with minor detours. The output is correct.
