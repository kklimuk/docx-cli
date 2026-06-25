# Grading rubric — resume (JUDGE ONLY)

The agent never sees this file. It is the ground-truth definition of "correct" for
the request in `task.md`. Judge the finished document against these checks.

## Pass conditions

- **Candidate data filled** — name, contact line, all education entries (school,
  location, degree/GPA, date), all experience entries (org, location, role, date,
  bullets), and the leadership entry (org, location, role, date, bullets) match the
  values in `brief.md`. No template placeholder text remains.
- **[Note: ...] helpers gone** — no `[Note: ...]` bracketed helper text remains
  anywhere in the document (`docx read` or `docx read --ast` shows none).
- **[drawing] element survives** — the `[drawing]` element is still present in the
  document body; it must not have been deleted.
- **Date tab stops preserved** — right-aligning tab stops on date fields survive;
  `docx read --ast` shows the paragraph still carries the tab stop definition and the
  date text is placed after a `\t` character.
- **Heading consistency via style definition:**
  - All three section headings (Education, Experience, Leadership & Activities) use
    the **Heading 1** style — `docx read --ast` shows `style: "Heading1"` on each.
  - The Experience heading in particular must be Heading 1, not plain bold centered
    text (that was its state before editing).
  - The restyle was applied to the **style definition** (one `docx styles` call with
    `--at Heading1`), NOT formatted onto each heading individually. `docx styles --at
    Heading1` shows `color: 1F3864`, `size: 13pt` (26 half-points), `bold: true`,
    `space-before: 12pt` (240 twips).
- **Entry spacing** — experience and leadership entry lines (org/title/date lines)
  and the contact line carry `spacing.after` set to at least 6pt (120 twips).
  `docx read --ast` shows `spacing: { after: 120 }` (or higher) on those paragraphs.
- **Page margins** — all four margins set to 0.5 inch (720 twips). `docx read --ast`
  shows the trailing section's `marginTop`, `marginRight`, `marginBottom`, `marginLeft`
  all equal 720. `docx read` shows `docx:page … margins="0.5in …"`.

## How to verify

- `docx read FILE` — confirm candidate text landed, no `[Note: ...]` survives, section
  headings look uniformly styled and navy, dates right-align, entries have breathing
  room, page geometry note shows 0.5in margins.
- `docx read FILE --ast` — check `style: "Heading1"` on all three section headings,
  `spacing.after` on entry paragraphs, section `marginTop/Right/Bottom/Left = 720`.
- `docx styles FILE --at Heading1` — confirm `color: 1F3864`, `size: 26` (half-points)
  or equivalent 13pt representation, `bold: true`, `spaceBefore: 240` (twips) or
  `space-before: 12pt`.
- Render the before and after with `docx render` — three matching navy bold headings,
  breathing room between entries, tight margins, no helper text, drawing still present.
