# Review — mnda (OfficeCLI, run r1)

**Scenario:** mnda — Form filling + highlight removal + font fidelity
**Bucket:** officecli-r1 (OfficeCLI comparison, haiku agent)

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (output page-001/002 PNGs present and inspected) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** (a fresh `docx read` matches `renders/output/read.md` byte for byte) |

## Merits

- **Every cover-page placeholder is filled with the right value and has no brackets:**
  - Purpose: "Evaluating a potential supply-and-integration partnership between the parties."
  - Effective Date: "June 8, 2026"
  - MNDA Term: "2 year(s)"
  - Term of Confidentiality: "3 year(s) from Effective Date"
  - Governing Law: "Delaware"
  - Jurisdiction: "New Castle County, Delaware"
  - Changes to Standard Terms: "None."

  In the output read view, the only `[ ... ]` left are the original `[ x ]` / `[    ]` checkbox glyphs, which were never placeholders.
- **The blank underline was kept on the filled values.** `<u>` wraps every value that was underlined before, such as `<u>June 8, 2026</u>` and `<u>Delaware</u>`. "None." was not underlined in the baseline and is still not underlined.
- **The signature table is filled in exactly the intended cells.** Party 1 values are in `t2:r2c1` through `t2:r6c1` and Party 2 values in `t2:r2c3` through `t2:r6c3`. Each destination has exactly one `<w:p>`. Spacer column c2, label column c0 and signature row r1 are empty and untouched. Both Date cells read "June 8, 2026".
- **The new signature text matches the document.** It is Arial 9pt (`sz=18`), the same size and face as the bold row labels next to it.
- **No yellow highlight is left.** All 26 baseline runs with `w:highlight="yellow"` are gone, and none remains anywhere in document.xml.
- **Fonts, sizes and colors are unchanged everywhere else.** I compared the run properties of every text run with the baseline. The only differences are on the edited runs, where yellow became none, plus the 10 new signature runs. No untouched run's formatting changed.
  - styles.xml, theme1.xml, numbering.xml and settings.xml are semantically identical to the baseline (canonical XML compare). The byte differences are serialization only.
  - The Georgia title, gray 7pt sub-labels, bold 9pt labels and teal Standard-Terms defined terms all look the same as before in the renders.
- **The visual before/after shows only the intended changes.** Page 2 (Standard Terms) is unchanged. Page 1 differs only in the filled values and the removed yellow. The Term of Confidentiality row is now one line shorter because the value text is shorter, which is expected reflow.

## Demerits

- **Minor: the cleared highlight was left as `<w:highlight w:val="none"/>` instead of being removed.** It is on 8 runs (Purpose, Effective Date, 2 year(s), 3 year(s), Effective Date, Delaware, New Castle County, None.). It renders with no highlight, so the requested behavior is met. A strict grep for `w:highlight` would still find it.
- **Minor: the paragraph marks of the filled signature cells gained `rFonts Arial` / `sz 18`.** The agent reported this itself. It has no visual effect.

## Frictions

The agent reported these frictions and dead ends:

- **Bracket-run formatting took over the replacement (major).** A find/replace whose match started in a bracket run used that run's formatting for the whole replacement, which dropped the blank underline. The agent worked around it by replacing the underlined inner text first and then removing the brackets in separate find calls.
- **A paragraph-level `highlight=none` would have touched every run in the paragraph.** The agent switched to setting it per find match.
- **The brief's mention of highlighted signature blanks was misleading.** The signature cells were empty and not highlighted, so the agent had to discover the empty cells itself.
- **Resident-mode save problems.** Reading the zip directly showed stale content until a save. One `save` call exceeded the 120s timeout while reporting "No pending changes".
- **Re-serialization churn.** Attribute-order changes appeared on untouched runs. They are semantically harmless; I confirmed this.

**Assessment:** the agent reached a clean, fully correct result, but the path was a moderate slog. Run-boundary formatting in find/replace forced multi-step workarounds, and the resident save stalled. The tool output itself is correct; the friction is in the ergonomics of cross-run replace and the save/timeout behavior.
