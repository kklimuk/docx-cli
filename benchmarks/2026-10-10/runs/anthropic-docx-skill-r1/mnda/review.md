# Review: mnda (Form filling + highlight removal + font fidelity) — Anthropic docx skill, run r1

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes (2 output pages rendered)
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes

## Merits
- All cover-page placeholders filled correctly with brackets removed: Purpose ("Evaluating a potential supply-and-integration partnership between the parties."), Effective Date (June 8, 2026), MNDA Term "Expires 2 year(s) from Effective Date.", Term of Confidentiality "3 year(s) from Effective Date, ...", Governing Law "Delaware", Jurisdiction "New Castle County, Delaware", Changes "None.".
- No highlight left: `grep w:highlight` on document.xml returns 0; `docx read` shows no `<mark>`.
- Signature table t2: Print Name / Title / Company / Notice Address / Date filled for both parties in c1 and c3 (r2–r6), each a single `p0`. Label column, spacer column c2, and signature row r1 are untouched. Both dates are filled.
- Fonts and layout unchanged: the baseline-vs-output markdown diff covers only the value cells. The Georgia title, bold 9pt labels, gray 7pt sub-labels, Arial body, teal header band, footer, and all of page 2 (Standard Terms) look the same in both renders.
- `docx read` parses the output cleanly.

## Demerits
- (minor, UX only) The agent reached the result through manual regex surgery on the XML and hit three dead ends along the way (see Frictions). The output is not affected.
- Note, not a defect: the only `[` characters left are the `[ x ]` / `[   ]` checkbox markers. These were never highlighted placeholders.

## Frictions
- The skill has no guidance on filling forms or removing highlights, and gives no way to list highlighted runs. The agent had to write regex over the XML itself.
- The task says the signature blocks are highlighted, but they are actually empty cells with no highlight. The skill has no pattern for filling an empty cell while keeping the paragraph-mark formatting.
- pandoc doesn't show highlights, so the agent had to check the XML directly to confirm they were gone.
- Dead ends:
  - A `<w:r` substring check also matched `<w:rPr>`, so the script aborted.
  - The "Notice Address" label cell has more than one paragraph, so label matching failed.
  - A sed command failed on macOS quoting.
- Read: the result is correct, but getting there was a moderate slog that needed hand-written XML scripting. A weaker agent could easily fail here.
