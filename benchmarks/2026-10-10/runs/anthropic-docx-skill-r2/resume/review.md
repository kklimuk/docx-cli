# Review: resume (Styling fidelity + drawing preservation), Anthropic docx skill arm, r2

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (one page)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Merits
- The Heading1 style definition was restyled once. `docx styles --at Heading1` shows 13pt, color 1F3864, bold, and 12pt space-before.
- Education, Experience and Leadership & Activities all use `Heading1`. Experience was converted from bold-centered BodyText. The agent stripped the direct pPr overrides so all three render as identical navy headings.
- All candidate data is filled in, and the mailto hyperlink rel was updated as well. No `[Note: ...]` text remains. Skills & Interests and the empty optional blocks are gone.
- `[drawing]` (the rule under the name) is preserved.
- space-after 6pt (120 twips) is set on the contact line and every org/title line (p18/p19/p23/p24/p28/p29).
- All four sections have margins of 720 twips on every side (`docx:page margins="0.5in"`).
- The page renders cleanly: dates and locations right-align without wrapping, and the bullets are intact.

## Demerits
- **minor**: The Education date lines were trimmed to "May 2027." / "June 2023.". They lose the "Graduating"/"Graduated" wording and keep a stray period.
- **minor**: The org/location lines' left tabs were changed to right tabs at 10800. The task said to keep the tab stops, but this cured the wrapping the new margins caused. The date lines keep their left tabs and end about 0.35in short of the margin.
- **minor**: The "Harvard Robotics Club" line sits slightly left of its sibling lines. This is a template indent quirk the agent left alone.

## Frictions
The agent reported these gaps in the skill:
- No guidance on editing styles.xml or on direct formatting that overrides a style.
- No guidance that margins are per section (four sectPr blocks).
- No helper for setting paragraph text, and merge_runs.py side effects plus a `-I` import failure.
- Left tab stops wrapped after the margin change.
- Hyperlink rels are not mentioned.
- `timeout` is missing on macOS.

Read: the output is correct, but the path was a slog of hand-written lxml editing.
