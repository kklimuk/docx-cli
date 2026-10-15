# Review: contract-markup (OfficeCLI, r2), bucket: editing / legal redlining + commenting

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** (all 5 rubric conditions pass) |
| Renders correctly (Word opened it) | **true**: 3 output pages rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

Render note: "OfficeCLI close exited 1 ... nothing to close" is benign. The disk copy was current and every render and read succeeded.

## Rubric check
1. **Changes tracked.** `docx track-changes list` shows 22 revisions (12 logical) on p7, t0:r1c2, p18, p22 and p24. PASS.
2. **At least 3 tracked text edits.** All four egregious terms are redlined:
   - §3: Net 90 at sole discretion becomes Net 30 from receipt of a correct invoice. The schedule cell changed with it.
   - §6: five years becomes one year, worldwide becomes United States, any industry becomes the same services, and $50,000 becomes $1,000.
   - §8: adds "except to the extent caused by the negligence or willful misconduct of the Company…".
   - §9: the cap is now mutual, the "$100" and "Contractor's liability … is unlimited" clauses are deleted, and the indirect-damages waiver is mutual.
   PASS.
3. **§9 line spacing.** p24 has `line-spacing="1.5"` in read, and the render shows it visibly opened up. PASS. The change is untracked, which the rubric allows.
4. **At least 4 anchored comments.** There are 9 comments, all with anchors. They cover §4 (p10), §7 (p20), §11 (p28), §12 (p30) and §13 (p32), plus §3, §6, §8 and §9. PASS.
5. **IP sub-points in lower-roman.** In `numbering.xml`, abstractNum 2 ilvl 1 is `lowerRoman` with `(%2)`. Word renders (i)–(iv), and the top-level 1–15 numbering is unchanged. PASS. docx-cli's `--ast` list object has no `format` field for this list, so the rubric's grep can't confirm it. That is a gap in docx-cli's reader, not an OfficeCLI defect. I checked the XML directly.

## Merits
- Clean, surgical redlines at the phrase level. They show up as proper Word deletion/insertion balloons.
- The agent caught the duplicate "Net 90" in the payment schedule table and redlined it too.
- The comments are specific and professional. They come from author "Mercer Design Studio" (initials MDS) and are anchored to the clause bodies.
- The before/after read.md diff touches only p7, t0:r1c2, p18, p22 and p24, all tracked. No wording changed silently.
- Header, footer, table styling, the signature block and the 3-page layout all match the baseline.

## Demerits
- **minor:** The numeral change edits the shared list-level definition, untracked. The task allows this ("however is cleanest"), and only the four IP items use that level, so nothing else changed.
- **minor:** The revised §9 cap is mutual but still set at one month's fees. That is fair in both directions, but the protection is thin. This is a drafting choice, not a rubric failure.
- **minor:** The agent's summary says the §9 heading also got 1.5 spacing. Only the body paragraph p24 has it.

## Frictions
- OfficeCLI `batch` reported success, exit 0, "(0 matched)", for four `set --find` entries that matched nothing. They targeted the auto-numbered (a)–(d) labels. This is a real tool UX bug: weak agents read a zero exit code as done.
- List labels can't be reached by find/replace. The only route is editing the numbering definition, which is untracked.
- The numbering-level selector is off by one: `get` prints lvl[2] for ilvl=1, but `set` needs level[1], and the error message gives no hint.
- `view issues` crashed on the pre-existing duplicate Heading1 styleId.
- The agent truncated its own batch output with `head -60`, so it couldn't see items 11–14.

Overall the path was a moderate slog, mostly from the false-success batch and the confusing numbering indices. The agent still recovered to a fully correct result.
