# Review — resume (Styling fidelity + drawing preservation) — OfficeCLI, run r3

## Verdict
| Dimension | Result |
|---|---|
| Task success | **partial** |
| Renders correctly (Word opened it) | **true** (output page-001.png produced) |
| Layout sound | **false** (tab-aligned locations/dates wrap) |
| Formatting preserved | **degraded** |
| Survived read loop | **true** |

On inspection, every rubric check passes. But the final page does not read as the "clean, well-spaced one-page résumé" the task asked for: the margin change makes the right-hand locations and dates wrap mid-phrase, and in two places mid-word.

## Merits
- **Candidate data is complete and correct.** Name; contact line, with both the email hyperlink's text and its mailto target updated; Harvard (A.B. CS, GPA 3.8, Graduating May 2027); Lincoln HS (Portland, OR, National Merit Finalist, Graduated June 2023); Northwind Robotics and Harvard SEAS with all three bullets; Harvard Robotics Club / Project Lead / 2024 – present with its bullet. No template placeholders remain.
- **Helper text and Skills are gone.** No `[Note: …]` text remains anywhere. The Skills & Interests heading and its four prompt lines are removed, and so are the unused Thesis, Coursework and Study Abroad template lines.
- **Headings share one style.** p8, p15 and p27 are all `Heading1` in the AST. Experience was converted from plain bold BodyText.
- **The style was restyled in one change.** `docx styles --at Heading1` shows: size 13pt, color 1F3864, bold true, space-before 12pt. In the render all three headings are matching navy. The agent also cleared the direct spacing and indent overrides on two headings so they truly match.
- **Entry spacing is set.** `spacing.after=120` is on the contact line (p6) and on p18, p19, p23, p24, p28 and p29.
- **Margins are set.** All four `<w:pgMar>` have top/right/bottom/left = 720.
- **Drawing and structure are kept.** The `[drawing]` rule under the name survives (one `w:drawing`). Tab-stop definitions and bullet numbering are preserved, there are no tracked changes, and `docx validate` reports the file valid.

## Demerits
- **MAJOR: locations and dates wrap.** In the output render, "Cambridge, / MA", "Graduating / May 2027", "Portland / , OR", "San / Francisco, CA" and twice "Cambrid / ge, MA" break onto separate lines.
  - **Cause:** going from 0.42in left/right margins to 0.5in shrank the text width from 7.66in to 7.5in. That leaves about 0.55in after the LEFT tabs at 6.948in.
  - **Evidence:** `read` flags `docx:layout … wrap="10 lines"` on p9–p29. The baseline render had none of this.
  - **Effect:** it fails the task's "dates still right-align" and "clean one-page résumé" outcome. The agent didn't render to check after changing the margins.
- **minor: the Harvard Robotics Club line is centered.** p28 keeps the template's `align=center`, so once its tab content wrapped, the org name renders centered with "Cambridg / e, MA" under it.
- **minor: date bolding is inconsistent.** "Jun 2025 – Aug 2025" and "2024 – present" are bold; "Sep 2024 – May 2025" is not.
- **minor: leftover empty paragraphs.** Empty p11/p14 leave a large gap in Education, and empty paragraphs p31–p34 remain at the end.

## Frictions
These are the frictions and dead ends the agent reported:
- **Hyperlink boundary:** find/replace refused any match that crossed the email hyperlink, so the agent split the replacement into pieces and updated the hyperlink separately.
- **Heading overrides:** direct `w:spacing before` on the Leadership heading and `w:ind` on the Experience heading shadowed the style. They could only be removed with raw-set.
- **Margin units:** the margin readback is in cm, so the 0.5in result is not obvious.
- **Outline view:** it didn't show the restyled Experience heading.
- **Unfilled template lines:** the task didn't say what to do with template entries that had no candidate data (Study Abroad, Coursework); the agent deleted them.

**Read:** the agent reached a rubric-compliant document, but only after some slog (hyperlink splitting, raw-set fallbacks). It stopped one step short: it never looked at the rendered page, so the visible wrapping that the margin change caused went unnoticed. A tool that warns about, or automatically fixes, tab overflow when margins change would have caught it.
