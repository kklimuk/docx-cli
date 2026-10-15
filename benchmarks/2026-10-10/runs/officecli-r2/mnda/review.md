# Review — mnda (OfficeCLI, run r2)

**Scenario:** mnda — form filling + highlight removal + font fidelity (editing bucket)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** — Word rendered 2 output pages |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** — `docx read` output identical to the render-time read view; `docx validate` passes |

## Merits
- All seven cover-page blanks filled with the exact task values and brackets removed: Purpose ("Evaluating a potential supply-and-integration partnership between the parties."), Effective Date (June 8, 2026), MNDA Term (Expires 2 year(s)), Term of Confidentiality (3 year(s) from Effective Date), Governing Law (Delaware), Jurisdiction (New Castle County, Delaware), Changes to Standard Terms (None.).
- The filled cover values keep the placeholder runs' original formatting (Arial, black text1, sz 16, single underline). Only the highlight changed.
- Signature table `t2` filled exactly in `c1`/`c3` of rows `r2`–`r6` for both parties. Each cell has a single `p0`. Label column `c0`, spacer `c2` and signature row `r1` are untouched, and both Date cells are filled. The values are Arial 9pt, which matches the 9pt labels.
- There is no yellow highlight anywhere. The only `<w:highlight>` values in document.xml are `none`.
- Visual before/after: page 1 is identical apart from the filled values and the removed highlight (the cover-page Georgia title, bold labels, gray 7pt sub-labels and header band are all intact). The table compresses slightly because the Confidentiality text now wraps to fewer lines. Page 2 (Standard Terms) is pixel-identical.
- Textual diff: baseline vs output read.md differ only in the `t1` value cells and the `t2` value cells.

## Demerits
- **minor:** The paragraph-level highlight clear stamped an explicit `<w:highlight w:val="none"/>` on about 30 runs, including runs that never had a highlight (the checkbox runs). It is invisible in the render and is XML churn only.
- **minor (not a real drift):** "None." is not underlined. The original `[Fill in, if any.]` placeholder was not underlined either, so this is faithful to the template.
- The template's wording "from Effective Date" (no "the") was kept. That is correct per "change ONLY the placeholder values".

## Frictions
Frictions the agent reported:
- **Major:** In `set --find/--replace` across runs, the replacement silently takes the first matched run's rPr, so the underline is lost when a plain `[` bracket run is first. The agent worked around it by deleting brackets with regex lookaheads (`\[(?=1 year)`) so the value run became the first match.
- **Minor:**
  - `set --prop highlight=none` at paragraph level wrote `none` onto runs that never had a highlight.
  - `query run[highlight=yellow]` returned an incomplete set, so the agent needed a raw XML grep to enumerate the highlights.
  - `get --depth 5` dumped about 37 KB of noise.
  - Filling the empty signature cells required `add --type run`.

**Read:** the output is correct, but the path was a moderate slog. The formatting-adoption trap in find/replace is a genuine tool hazard that could easily have produced silent underline loss for a less careful agent.
