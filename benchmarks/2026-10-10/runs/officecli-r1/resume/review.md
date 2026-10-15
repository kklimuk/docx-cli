# Review: resume (OfficeCLI, run r1)

**Scenario:** resume. Bucket: styling fidelity and drawing preservation.

## Verdict

| Dimension | Result |
|---|---|
| Task success | **partial** |
| Renders correctly (Word opened it) | **true** |
| Layout sound | **false** |
| Formatting preserved | **degraded** |
| Survived read loop | **true** |

On the structural checks the rubric lists, the agent passes almost everything. The data is filled, the notes are gone, Skills & Interests is removed, the drawing is kept, all three headings use Heading 1, and the Heading 1 style definition holds the requested values. Spacing-after and the 0.5in margins are also set. The rendered page is a different story. Every location/date line in the résumé now wraps, and some break in the middle of a word ("Cambrid / ge, MA", "Cambridg / e, MA", "Portland / , OR"). The task's "done" state asks for dates that right-align on a clean one-page résumé, and this page is not that.

## Merits

- **Candidate data is complete and correct.** The read view and render show Priya Raman, the full contact line, Harvard (A.B. in Computer Science, GPA 3.8, Graduating May 2027), Lincoln High School (National Merit Finalist, Graduated June 2023), Northwind Robotics with its 2 bullets, Harvard SEAS with its 1 bullet, and Harvard Robotics Club with its 1 bullet. No template placeholder text is left.
- **The email hyperlink was retargeted properly.** `hyperlinks list` gives `link0` with text `praman@college.harvard.edu` and url `mailto:praman@college.harvard.edu`. It was not flattened to plain text.
- **All `[Note: ...]` text is gone.** `grep -c 'Note:' word/document.xml` returns 0.
- **Skills & Interests is fully removed.** The heading and all 4 prompt lines are gone.
- **The `[drawing]` is kept.** The VML `<w:pict>`/`v:line` rule is still in p4 and renders under the name.
- **Headings are made consistent through the style.** All three headings carry `<w:pStyle w:val="Heading1"/>`. Experience was converted from BodyText. `docx styles --at Heading1` shows size 13pt, color 1F3864, bold true and space-before 12pt. In the render the three headings are the same navy, bold, 13pt and centered.
- **Entry spacing is set in place.** `space-after="6pt"` (`w:after="120"`) is on the contact line (p6), the org/title lines (p19, p20, p24, p25) and the Leadership org/role lines (p29, p30).
- **Margins are 0.5in in all four sections.** All 4 `<w:pgMar>` elements are 720/720/720/720, and the read view shows `docx:page s0 margins="0.5in"`.
- **Tab stops and bullets are preserved.** The tab definitions are unchanged, dates follow a `\t`, and the bullet numbering is intact.
- **The file is valid and opens in Word.** `docx validate` reports valid with 8 parts checked, and Word rendered the page.

## Demerits

1. **MAJOR: dates and locations wrap, some mid-word.** The output render shows "Cambridge, / MA", "Graduating / May 2027", "Portland / , OR" (Portland on its own line, then ", OR" on the next), "San / Francisco, CA", "Cambrid / ge, MA" (Harvard SEAS) and "Cambridg / e, MA" (Robotics Club). In the baseline render every one of these fit on a single line.
   - **Cause:** the template places locations and dates with LEFT tab stops at fixed positions (8621, 9432, 10005 twips and so on). Changing the margins from 0.42in to 0.5in on the left and right made the text area 0.16in narrower, and the new content ("San Francisco, CA", "Graduating May 2027") is longer than the placeholders.
   - **The agent saw the problem and did not fix it.** It noticed the tabs were left-aligned (its own dead-end note) and left them "as instructed." The task says "The dates still right-align… reads as a clean, well-spaced one-page résumé," and this result visibly fails that.
2. **MINOR: inconsistent bolding on dates.** "**Jun 2025** – Aug 2025" is only partly bold (`<w:b/>` on the "Jun 2025" run only), and "**2024 – present**" is fully bold. "Sep 2024 – May 2025" is not bold. The template dates were not bold, so the new text picked up the bold of an adjacent template run.
3. **MINOR: stray line break and large gaps in Education.** The Harvard degree paragraph (p10) still ends with the `<w:br/>` that came before the deleted "Thesis" line. Together with the template's empty spacer paragraphs and the Study Abroad block that was removed, this leaves a large blank gap between the Harvard and Lincoln entries in the render.
4. **MINOR: Leadership org line is still centered.** "Harvard Robotics Club" (p29) keeps the template's `jc=center`, so it sits centered under the heading and its location wraps to a centered line. This came from the template (baseline p37 was also centered), but it adds to how ragged the page looks.
5. **MINOR: redundant direct formatting on Experience.** The Experience heading keeps direct `<w:b/>` runs and a `<w:sz w:val="16"/>` paragraph-mark override from its earlier BodyText state. There is no visible effect, because the style is bold and the run size comes from the style.

## Frictions (from the agent's account)

- **False premise in the task (major).** The template's date tabs are LEFT tabs, not right tabs. The agent correctly diagnosed this but chose not to fix it, and the visible wrapping above is the result. docx-cli's own `read` of the baseline flags this exact problem with a `docx:layout … fix-all="--tabs right"` hint. OfficeCLI gave the agent no such hint, and the agent had no obvious verb for changing tab alignment.
- **`set --find` with zero matches exits 0** and reports `matched 0` with no warning, which is a silent no-op trap.
- **Find/replace across the email hyperlink boundary failed with `internal_error`.** The agent worked around it with `hyperlink[1] url/text set`, which worked well.
- **Margins had to be set on each of the 4 sections separately.** The document-level readback shows only the final section.
- **Readback units are confusing:** 0.5in reads back as 1.27cm.
- **The text and issues views don't show the VML drawing**, so the agent could not confirm it was still there.
- **`view issues` reports noise:** 16 empty-spacer warnings.
- **Stray leading spaces in template runs** needed extra find calls to trim.
- **The agent did not render the page to check it**, and that is how the wrapping defect got through.

**Overall path:** a moderate slog. There were several dead ends (the hyperlink-boundary failure, the silent zero-match set, per-section margins), but the agent reached a structurally correct document. Its one real misjudgment was declining to cure the tab wrapping and not rendering the page to see it.
