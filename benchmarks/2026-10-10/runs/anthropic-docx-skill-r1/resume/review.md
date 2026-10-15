# Review: resume (Styling fidelity + drawing preservation) — Anthropic docx skill, r1

**Scenario:** resume · **Bucket:** editing / styling fidelity + drawing preservation · **Arm:** competitor (Anthropic docx skill, haiku)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (one-page output PNG) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

## Rubric check
- Candidate data filled: PASS. Name, contact (mailto retargeted to praman@college.harvard.edu), both schools, both jobs with bullets, and the Robotics Club entry are all present. No placeholders remain.
- `[Note: ...]` helpers gone: PASS. None appear in the read or the render.
- Skills & Interests removed: PASS. The heading and all four prompt lines are gone.
- `[drawing]` survives: PASS. It is at p4, and the rule under the name shows in the render.
- Date tabs: PASS in intent. Dates sit after `\t` and right-align in the render. The agent replaced the template's left tabs (left@6.5–6.9in, which would wrap at 0.5in margins) with `right@10800`.
- Heading 1 consistency: PASS. p8, p15 and p27 are all `style: Heading1`. `docx styles --at Heading1` shows color 1F3864, 13pt, bold and space-before 12pt, set once in styles.xml. The render shows three identical navy bold headings.
- Entry spacing: PASS. `after: 120` is on p6 (contact), p18/p19, p23/p24 and p28/p29.
- Margins: PASS. All four sectPrs have pgMar 720/720/720/720, and `docx:page margins="0.5in"`.

## Merits
- Every pass condition in the rubric is met, and the render is a clean one-page résumé.
- The heading restyle went into the style definition, not onto each heading.
- The section structure (including the 3-column continuous section), bullet numbering and the drawing were preserved.
- The agent spotted and cured the tab overflow at the new margins on its own initiative.

## Demerits
- **Minor:** Stray trailing periods were copied from the brief: "Graduating May 2027.", "Graduated June 2023." and "GPA 3.8." (render, read.md p10/p13).
- **Minor:** "Harvard Robotics Club" (p28) lost its template centering and has no left indent, so it sits about 0.08in left of the other entry lines (visible in the render).
- **Minor:** "Software Engineering Intern" (p19) keeps the template's 0.121in indent against 0.083in elsewhere, a slight visible offset. The agent noticed it and left it.
- **Minor:** The tab stops were changed (left → right@10800), which departs from "keep exactly as they are". The change is justified because the original left tabs would wrap at 0.5in margins, and the intent (right-aligned dates) is satisfied.

## Frictions
Agent-reported:
- **No guidance on tabs after a margin change (major).** Left tabs overflowed and had to be found by rendering, then cured by hand.
- **No helper for paragraph spacing or style edits (major).** The agent hand-edited XML with lxml and worked out schema child order (spacing before ind, outlineLvl before rPr) itself.
- **Margins are per sectPr (minor).** There are four sectPrs, and the skill doesn't mention it.
- **soffice.py returned exit 1 after writing the PDF (minor).** It also printed "Task policy set failed".
- **validate.py fails under `python3 -I` (minor).** The run hit ModuleNotFoundError: helpers.
- **Leftovers (minor).** Column widths in the 3-col section exceed the new text width, and header distance 1094 is larger than the 720 top margin. Neither is visible in the render.

**Read:** The path was a moderate slog. The skill provides no verbs for styles, spacing or margins, so the agent did correct but manual XML surgery and needed renders to find the tab problem. Misleading exit codes from soffice.py and validate.py would trip a weaker agent. The output itself is correct.
