# mnda (OfficeCLI, r3): form filling, highlight removal, and font fidelity

**Bucket:** editing / form-fill

## Verdict
- Task success: **success**
- Renders correctly (Word opened it): **true**
- Layout sound: **true**
- Formatting preserved: **preserved**
- Survived read loop: **true** (a fresh `docx read` matches the output read view byte for byte)

## Merits
- All cover-page placeholders are filled with the exact values: Purpose, Effective Date (June 8, 2026), MNDA Term `2 year(s)`, Term of Confidentiality `3 year(s)` from `Effective Date`, Delaware, New Castle County, Delaware, and `None.`. No `[...]` brackets remain.
- The global `1 year(s)` double-match was handled correctly. The two terms carry different values.
- The signature table t2 is fully populated: r2–r6 × c1/c3 hold Dana Okafor / Marcus Bell, their titles, both companies, both emails, and June 8, 2026 twice. Each cell is a single `p0`. The label column c0, spacer c2 and signature row r1 are untouched.
- There is no yellow highlight anywhere. The XML has only `w:highlight w:val="none"`, and the render is clean.
- Font fidelity: the baseline and output read views differ only in the filled cells. The Georgia title, 9pt bold labels, 7pt gray `#8C8D8E` sub-labels and Arial body are all unchanged. Signature values are Arial 9pt, matching the cells' paragraph-mark rPr. Page 2 is visually identical to the baseline.

## Demerits
- **minor**: The highlight was cleared by writing `<w:highlight w:val="none"/>` onto every run and paragraph mark in each edited paragraph (31 occurrences), including runs that were never highlighted, such as "The laws of the State of " and ".". It is visually equivalent and valid, but it adds XML noise.

## Frictions
- Cross-run find/replace merged runs and took the first run's formatting, which silently dropped the underline on the purpose value. The agent caught this and redid the edit with smaller finds.
- Resident mode kept stale in-memory state after the agent copied a file over the open document.
- Clearing highlight needed a separate paragraph-wide `set --prop highlight=none`.
- Filling an empty cell needed `add --type run` with explicit font.latin, font.cs, size and size.cs props, and the help doesn't document this.
- Table `get --depth` returned about 30KB of output.

**Read:** The output is correct, but getting there was a slog with several dead ends. The silent formatting loss on cross-run replace is a real OfficeCLI bug that a less careful agent would have shipped.
