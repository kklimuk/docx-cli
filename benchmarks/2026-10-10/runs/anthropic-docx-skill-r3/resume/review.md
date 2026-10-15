# Review: resume (Styling fidelity + drawing preservation), competitor arm (Anthropic docx skill, r3)

## Verdict
- **Task success:** success. Every rubric pass condition is met. The remaining defects are cosmetic.
- **Renders correctly (Word opened it):** true. Word rendered one output page.
- **Layout sound:** true. Everything fits on one page, nothing is pushed off the page, there are no blank pages, the drawing rule is under the name, and the section breaks are intact.
- **Formatting preserved:** degraded. The date and location column is ragged, and one role line has an odd indent (details under Demerits).
- **Survived read loop:** true. `docx read`, `--ast` and `docx styles --at Heading1` all show the intended state.

## Rubric check
| Check | Result | Evidence |
|---|---|---|
| Candidate data filled | PASS | The render and read.md show the name, contact line (mailto link updated), Harvard / Lincoln HS, Northwind, Harvard SEAS and Harvard Robotics Club. All roles, dates and bullets match task.md. |
| [Note: ...] removed | PASS | read.md has no `[Note`. The Thesis, Relevant Coursework and Study Abroad placeholder lines were removed. |
| Skills & Interests removed | PASS | The heading and all four prompt lines are absent. |
| [drawing] survives | PASS | p4 still contains `[drawing]`, and the rule under the name is visible in the render. |
| Date tab stops / `\t` | PASS (with caveat) | The date paragraphs keep their tab stops, and each date follows a tab. The agent changed the stops from left to right (for example p9 right@9489, p10 right@9360, p20 right@8621, p30 right@8622). |
| All three headings Heading1 | PASS | The AST shows `style: Heading1` on p8, p16 (Experience, previously BodyText) and p28. |
| Style definition restyled once | PASS | `docx styles --at Heading1` shows size 13pt, color 1F3864, bold true, space-before 12pt. The render shows three identical navy bold headings. |
| Entry spacing after >= 6pt | PASS | `spacing.after=120` is set on p6 (contact), p19, p20, p24, p25, p29 and p30. |
| Margins 0.5in | PASS | All of s0 to s3 have marginTop, marginRight, marginBottom and marginLeft = 720. read.md shows `margins="0.5in"`. |

## Merits
- All candidate content is correct and complete. The hyperlink rel target was updated to `mailto:praman@college.harvard.edu` so it matches the visible text.
- The Experience heading was moved to Heading1, and the Heading1 style definition was edited once. All three headings render identically in navy, 13pt and bold.
- The agent spotted a direct `spacing before` override on the Leadership heading that would have broken consistency, and removed it.
- The VML `[drawing]` and the multi-section structure (continuous breaks, the 3-column s1) are untouched.
- Margins were applied to every sectPr. The result is a clean one-page document.
- The Skills section and all `[Note]` helper text are gone.

## Demerits
- **minor/major (cosmetic): the date and location column is ragged and stops short of the right margin.** The right tab stops land at different positions: 9489, 9432, 9360, 8621 and 8622 twips, against a text width of 10800 twips. In the render, "Graduating May 2027" ends at about x=1050, "Jun 2025 – Aug 2025" and "2024 – present" at about x=973, and the locations at about x=1063–1117. None of them reaches the margin at about x=1210. task.md asked that "the dates still right-align". Each date is right-aligned to its own stop, but the column does not line up and does not look professional.
- **minor: the "Software Engineering Intern" line is indented 174 twips, while the other entry lines use 119/120.** This is visible in the render as a slight indent compared with "Research Assistant" and "Project Lead". It was inherited from the template, but it was not normalized.
- **minor: there is a large vertical gap between the Harvard and Lincoln High School entries.** The leftover empty paragraphs (p11 and p12) from the removed Thesis and Study Abroad lines produce visible dead space in the Education section.

## Frictions
The agent reported these frictions and dead ends:
- The skill has no style inspect/edit verb, so styles.xml had to be hand-edited.
- The template's "right-aligned" date tabs are actually LEFT tabs. Converting them in place gave mid-line right stops.
- Direct pPr overrides (spacing and jc) silently defeat a style edit.
- After the margins were narrowed, the location tabs overflowed and wrapped ("San / Francisco"), which the first render showed.
- The Robotics Club org line had a stray jc=center and was missing its indent.
- The skill does not mention hyperlink rels, and merge_runs.py gives no paragraph index.

Read: the path was a slog. Everything was raw XML surgery across document.xml, styles.xml and the rels, with several render-and-fix iterations. The agent recovered well, and only the tab-stop positioning stayed imperfect.
