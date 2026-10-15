# Review — resume (Styling fidelity + drawing preservation)

**Bucket:** editing — résumé styling / style-definition restyle / drawing preservation

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (1 output page PNG) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** (`docx read` matches the harness read view byte-for-byte) |

## Merits
- **All candidate data filled.** Name, contact line (address · mailto email link · phone), Harvard (A.B. in Computer Science, GPA 3.8, Graduating May 2027), Lincoln High School (Portland, OR, National Merit Finalist, Graduated June 2023), both experience entries with all three bullets, and the leadership entry with its bullet all match the brief. No template placeholder text is left.
- **[Note: …] helpers removed.** None remain in `read`. The optional template lines with no candidate data (Thesis, Relevant Coursework, Study Abroad) were removed instead of left as prompts.
- **Skills & Interests section fully removed**: the heading and all four prompt lines are gone.
- **`[drawing]` preserved** at p4, and the horizontal rule under the name still renders.
- **Heading consistency:** the AST shows `style: "Heading1"` on p8 Education, p16 Experience, and p28 Leadership & Activities. Experience was BodyText bold centered in the baseline (p19).
- **Style definition restyled:** `docx styles --at Heading1` returns size 13pt, color 1F3864, bold true, space-before 12pt. The render shows three matching navy bold headings.
- **Entry spacing:** `spacing.after: 120` is set on the contact line p6 and on the entry lines p19, p20, p24, p25, p29, and p30.
- **Margins:** every section (s0–s3) has marginTop/Right/Bottom/Left = 720, and `read` shows `docx:page s0 margins="0.5in"`. The baseline's mixed per-section geometry (`varies="by-section"`) is now uniform.
- **Dates right-align:** every entry line carries `tabStops: [{align: right, pos: 10800}]`, with the date after the tab. The baseline's `docx:layout` overflow warnings are gone, and in the render every date sits flush on the right margin.
- Bullet lists are preserved (ListParagraph with bullets in the render). `docx validate` passes. There are no tracked changes or comments.

## Demerits
- **minor (pre-existing, not introduced):** some alignment quirks carry over from the template. "Harvard Robotics Club" (p29, align=center with no left indent) sits a few points left of the other entry lines, and "Software Engineering Intern" (indent 0.121in) sits slightly right. Both match the baseline's own indents on those paragraphs, so they are not agent defects. The agent also didn't clean them up.
- **minor (pre-existing):** the empty 3-column continuous section (s1, p18) still adds a gap under the Experience heading. It is template structure and was preserved as the task asked.
- **minor:** the tab stops changed from LEFT to RIGHT. `sections --margins` did this automatically, the "fix, don't hint" auto-reflow. The result matches what the task described, right-aligned dates, and the render improves on it. The agent noted that the change came with little notice.

## Frictions
- **Agent dead end (major UX / possible tool bug):** a batch `edit` entry `{"at":"p19","style":"Heading1","clear":"bold"}` acked "edit 30 changes", but by the agent's account the paragraph kept the BodyText style. The agent recovered with a standalone `edit --at p16 --style Heading1`. If the report is accurate, `--style` combined with `--clear` in a batch entry is silently dropped while the ack still counts the entry. That is a no-op-reported-as-success bug worth reproducing.
- The tab conversion done by `sections --margins` (10 left→right tabs) surfaced only as one line in the ack.
- Batch `--dry-run` previews list locators only.
- The agent skipped the one-page render check. The render confirms the document fits on one page.
- **Overall:** the result is fully correct. The path was mostly smooth, with one significant dead end (the silent style drop in a batch) that the agent caught by re-verifying.
