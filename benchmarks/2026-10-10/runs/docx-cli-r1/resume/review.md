# Review — resume (Styling fidelity + drawing preservation)

**Bucket:** editing · **Run:** docx-cli-r1 (haiku)

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (1 output page rendered) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

## Verification performed

- `docx styles resume.docx --at Heading1` → `size 13pt`, `color 1F3864`, `bold true`, `space-before 12pt`, so the style definition was restyled.
- `docx read --ast`: p8 Education, p16 Experience and p28 Leadership & Activities all have `style: "Heading1"`. In the baseline, Experience was `BodyText` with bold, centred text.
- Spacing: p6 (contact line), p19/p20, p24/p25 and p29/p30 all have `spacing.after: 120`.
- Margins: every section (s0–s3) has `marginTop/Right/Bottom/Left = 720`. The read view's head shows `docx:page s0 margins="0.5in"`.
- Tabs: the date and location lines (p9, p20, p30, …) carry `tabStops: [{align: right, pos: 10800}]`, and the date text comes after the tab run.
- `[drawing]` is still at p4. A `find "Note"` returns no matches, and none of the Skills & Interests lines remain.
- `docx validate` reports `valid — 8 parts checked`.

## Merits

- All candidate data is filled correctly, including the name, the contact line with its mailto link, both education entries, both experience entries with their bullets, and the leadership entry with its bullet. No template placeholders remain.
- The Experience heading was converted to Heading1. The style definition was then restyled in one `styles set` call, and the render shows three matching navy, bold headings.
- The `[Note: …]` helper text is gone. The Skills & Interests section was removed in full, as requested.
- The drawing (the rule under the name) was preserved, and the render still shows the horizontal rule.
- Entry spacing was set in place without retyping the text. The 0.5in margins apply to all sections.
- Dates right-align cleanly in the render. In the baseline the dates sat on left tabs that nearly overflowed. The `--margins` auto-reflow moved them to right tabs at the text edge, which gives the result the task wanted.
- The output is a clean single page and validates against the schema.

## Demerits

- **minor**: The Leadership & Activities heading (p28) has a direct `space-before=240` instead of inheriting it from the style. The template's direct 0.05pt override was replaced, not cleared. The heading looks identical, but it is a direct override, not purely the style.
- **minor**: The Harvard Robotics Club line (p29) is still centre-aligned, which is a direct property from the template. In the render it starts slightly left of the other org lines. Software Engineering Intern (p20) also keeps a slightly larger left indent. Both issues come from the template and were not introduced by the agent, but the polish request would ideally have caught them.
- **minor**: Leftover blank spacer paragraphs (p11/p12 and the empty ones around the deleted Thesis, Study Abroad and Skills lines) add uneven vertical gaps. The gap after the Harvard degree line is visibly large. This is cosmetic; the page still fits easily on one page.
- **minor (debatable)**: The tab stops changed from left to right alignment. The task said to keep the tabs as they are, but it also described them as tabs that right-align the dates. The visible result matches that intent and the rubric's "right-aligning tab stops survive" check.

## Frictions

- The agent reported that `sections --margins 0.5` silently converted 10 left tabs to right tabs. The change was beneficial, but the agent was surprised by it. **This is a tool UX issue:** the margins ack should name the tab reflow.
- Dead end: `edit --at p28 --clear space-before` failed with USAGE because `--clear` does not accept paragraph spacing. The agent worked around it with a direct `--space-before 12`. **This is a tool gap:** `--clear` cannot remove direct paragraph properties that shadow a style.
- The agent found the direct spacing override on the heading only through `--ast`. Neither the read view nor `styles --at` showed that the override shadowed the style.
- `timeout` is not available on macOS, so the agent re-ran `render` without it. This is an environment issue, not a tool issue.
- Overall, the path was mostly smooth. One batched edit handled all 30 text changes and one `styles set` handled the restyle. The only small slog was the heading's direct spacing.
