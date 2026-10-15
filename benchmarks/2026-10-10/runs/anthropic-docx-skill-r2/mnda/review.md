# Review: mnda (Form filling + highlight removal + font fidelity), competitor arm: Anthropic docx skill (r2)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true. Both output pages rendered.
- **Layout sound:** true. The cover page lost about one line of height because the Term of Confidentiality value is shorter. No overflow and no blank pages. Tables are intact.
- **Formatting preserved:** preserved
- **Survived read loop:** true. `docx read` shows every value. There is no `<mark>` and no placeholder brackets.

## Merits
- All 7 highlighted cover-page placeholders now hold the correct values, with the brackets removed:
  - Purpose: "Evaluating a potential supply-and-integration partnership between the parties."
  - Effective Date: June 8, 2026
  - MNDA Term: 2 year(s)
  - Term of Confidentiality: 3 year(s) from Effective Date
  - Governing Law: Delaware
  - Jurisdiction: New Castle County, Delaware
  - Changes to Standard Terms: None.
- There are zero `<w:highlight>` elements in `word/document.xml`. The only `[` left in the read view belongs to the commonpaper hyperlink markdown.
- In the signature table t2, values are in the right cells: r2–r6 of c1 (Party 1) and c3 (Party 2). Each cell holds a single `p0`. The label column c0, the spacer column c2 and the signature row r1 are untouched. Both dates are filled.
- Font fidelity: the diff of the baseline and output read views changes only the filled cells. The Georgia title, bold 9pt labels, gray 7pt sub-labels and the Arial body are unchanged. Page 2 (Standard Terms) is visually identical.
- Filled values kept the placeholder run's rPr, minus the highlight. Signature values are 9pt, matching the labels.

## Demerits
- **minor**: the cover-page values keep the placeholder's underline (for example `<u>Delaware</u>`, `<u>June 8, 2026</u>`). This is defensible under "change only the values and the highlight". The baseline placeholders were underlined too.

## Frictions
- Reported by the agent:
  - SKILL.md has no recipe for replacing highlighted placeholders that are split across several runs. The agent wrote a custom regex/Python script for it.
  - SKILL.md has no guidance on filling empty table cells. The agent worked out on its own that the paragraph-mark rPr should be copied into a new run.
  - The agent was unsure whether to keep the underline.
  - `merge_runs.py` breaks under `python3 -I` because it imports sibling modules.
  - `timeout` is not installed on macOS.
  - The task text says the signature blocks are highlighted, but they are not.
- My read: the result is correct, but the path was moderately laborious. The skill gives low-level XML access with no form-fill helper, so a weak agent has to write its own run-grouping logic. This time it got that logic right.
