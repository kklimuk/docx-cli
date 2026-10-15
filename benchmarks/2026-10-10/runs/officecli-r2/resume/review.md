# Review — resume (Styling fidelity + drawing preservation) — OfficeCLI r2

**Verdict**
- Task success: **partial**
- Renders correctly (Word opened it): **true** (one page rendered)
- Layout sound: **false**
- Formatting preserved: **degraded**
- Survived read loop: **true**

## Merits
- All candidate data filled: Priya Raman; contact line with the hyperlink text and `mailto:` target updated; A.B. in Computer Science, GPA 3.8 / Graduating May 2027; Lincoln High School, Portland, OR / National Merit Finalist / Graduated June 2023; Northwind Robotics (2 bullets), Harvard SEAS (1 bullet), Harvard Robotics Club (1 bullet).
- No `[Note: ...]` text and no template prompts remain. Skills & Interests is gone completely. Study Abroad, Coursework and Thesis lines were dropped because the candidate gave no data for them, which is reasonable.
- The Experience heading is now Heading1. All three headings (p8, p16, p28) are `style: Heading1` and render as identical navy bold headings.
- `docx styles --at Heading1` shows 13pt, color 1F3864, bold, and space-before 12pt, so the change was made on the style definition.
- `spacing.after = 120` is set on p6 (contact line) and on p19/p20/p24/p25/p29/p30 (the org/title/date lines).
- All four `w:pgMar` are 720 on every side; read shows `margins="0.5in"`.
- The `w:pict` drawing (the rule under the name) is still present. Tab stops and bullet numbering are intact. `docx validate` passes.

## Demerits
1. **Major: right-column locations and dates wrap and split words in Word.** In output/page-001.png: "Cambridge, / MA", "Graduating / May 2027", "Portland / , OR", "Graduated / June 2023", "San / Francisco, CA", "Cambrid / ge, MA", "Cambridg / e, MA". The baseline shows every one of these on a single line. The cause is a combination of three things: the template's LEFT tab stops at about 6.5–6.95in, the narrower text width after the right margin went from 0.42in to 0.5in, and the longer real content. The read view flags 10 wrapping lines (`docx:layout … fix-all="edit FILE --at p9-p30 --tabs right"`). The task required "dates still right-align" and a "clean one-page résumé", and this result does not deliver that.
2. **Minor: the "Harvard Robotics Club" line now renders centered.** It picks up the template's direct `align=center` on p29, which shows now that its tab content wraps. The result does not match the left-aligned Experience org lines.
3. **Minor: uneven blank gaps in Education.** There is a large gap before Lincoln High School and another between that heading and its wrapped location.

## Frictions
These are the agent's reports:
- Find/replace was refused across the hyperlink boundary in the contact line. The agent worked around it by setting the hyperlink text and URL separately.
- No verb clears direct paragraph spacing or indent that overrides the style. The agent fell back to a raw XPath remove.
- Removing the trailing break needed an index-based run remove.
- `view pdf` failed with "No exporter plugin found" and there is no page-count stat, so the agent could not render.
- The `view issues` output was noisy and the help text is very large.

**Overall read:** getting to a correct XML/rubric state was a moderate slog. The decisive gap is that OfficeCLI gave no render or layout feedback, so the agent never saw that its margin change plus the longer content broke the right-aligned columns.
