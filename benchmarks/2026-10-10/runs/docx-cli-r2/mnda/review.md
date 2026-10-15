# Review: mnda (editing bucket: form filling, highlight removal, font fidelity)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: 2 output pages rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true**: the live `docx read` output is identical to the output read.md |

## Merits
- All 8 cover-page placeholders are filled with the task values:
  - Purpose
  - Effective Date: June 8, 2026
  - Expires 2 year(s)
  - 3 year(s) from Effective Date
  - Delaware
  - New Castle County, Delaware
  - None.
- No `[...]` placeholders remain. The only `[` left are the template's original `[ x ]` and `[    ]` checkbox markers.
- There is zero highlight in the file: `grep '<w:highlight'` on document.xml finds nothing, and `find --highlight any` returns no matches.
- The signature table t2 is filled in the intended cells:
  - Party 1 values are in c1 and Party 2 values are in c3, for rows r2 to r6 (Print Name, Title, Company, Notice Address, Date).
  - Each of those cells holds exactly one paragraph, `p0`.
  - The label column c0, the spacer column c2 and the signature row r1 are unchanged.
  - Both dates are filled.
- Fonts are preserved:
  - The baseline vs. output read.md diff shows only the filled values changing.
  - Page 2 is visually identical to the baseline.
  - The Georgia title, the bold labels, the gray 7pt sub-labels and the Arial body are unchanged.
  - The cover-page values keep the placeholder look: Arial 8pt, black, underlined.
  - The signature values are Arial 9pt, which matches their labels.
- `docx validate` passes (11 parts checked).

## Demerits
- **minor (tool UX):** A span replacement with `edit --at pN:S-E --text` inherited the formatting of the first run in the span. That run was the un-underlined `[` bracket, so the value's underline was silently dropped. The agent needed a second batch to underline the value spans again. The final XML is correct (`<w:u w:val="single"/>` is on the filled runs), so this is a cost of reaching the result, not a defect in the output.
- **minor (discoverability):** `edit --help` shows no example of "fill a bracketed blank, clear the highlight, keep the value's run formatting." The agent worked it out by trial and by inspecting the AST.

## Frictions
- **Agent-reported:**
  - The span fill lost the underline (a major issue from the agent's point of view).
  - It was unclear what size a bare-cell fill would get.
  - It thought the size shown on the PARTY 1/2 headers in the read view had drifted. I could not reproduce this: the baseline and output read.md files do not differ on those lines.
  - The fill-plus-unhighlight recipe was hard to discover.
- **Dead ends:**
  - The task text says the signature blanks are highlighted, but the source has no highlight there.
  - Hand-counting offsets was error-prone, so the agent used `find` to get them.
- **Assessment:** The signature table was a smooth fill. The cover-page blanks were a moderate slog because span replacement inherits bracket formatting, which forced a second underline pass. The tool should keep the formatting of the run that holds the value, not the bracket's.
