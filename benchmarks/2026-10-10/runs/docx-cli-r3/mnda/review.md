# Review — mnda (Form filling + highlight removal + font fidelity), docx-cli r3

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (2 output pages rendered)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true (fresh `docx read` is identical to the output read.md; `validate`: valid, 11 parts)

## Merits
- All eight cover placeholders hold the task values: Purpose sentence, June 8, 2026, "Expires 2 year(s) from Effective Date.", "3 year(s) from Effective Date, but in the case of trade secrets…", "The laws of the State of Delaware.", "…located in New Castle County, Delaware.", "None.". No `[ ... ]` placeholders remain; the only brackets left are the original `[ x ]` / `[    ]` checkboxes.
- Zero highlight: `find --highlight any` returns no matches, and `word/document.xml` has 0 `<w:highlight>` elements.
- Signature table t2: Print Name, Title, Company, Notice Address and Date are filled for both parties in c1/c3, rows r2–r6. Each cell has a single p0. Labels (c0), the spacer (c2) and the signature row (r1) are untouched.
- Font fidelity: the filled cover runs keep Arial 8pt, color 000000. The Georgia 22pt title, the bold 9pt labels, the gray 7pt sub-labels and all of page 2 are unchanged in both the render and the read diff. `theme1.xml` is byte-identical, and `styles.xml` differs only by a newline after the XML declaration.
- Efficient path: two `edit --batch` calls (one using `--clear highlight`), dry-run first, then verified with find, diff and validate.

## Demerits
- **minor**: the placeholder underline was dropped on the filled values. The baseline shows `<mark>[<u>1 year(s)</u>]</mark>`, and the output shows plain `2 year(s)`, with no underline in the page-001 render. The underline arguably belongs to the blank's styling, but the agent said nothing asked for it to go and the tool removed it silently.
- Observation, not a defect: signature fills render at 9pt Arial, which matches the 9pt labels in those rows. The 8pt base was not used. The cells were empty, so no existing text drifted.

## Frictions
- The agent reported three minor frictions: (1) `edit --text` over an underlined, highlighted span silently dropped the underline; (2) nothing tells the agent what font and size a bare-cell fill inherits; (3) batch confirmations ("edit 8 changes") don't list locators.
- Dead ends: one wrong relative path to the binary, with no other effect.
- My read: the path was smooth and close to the intended one. The frictions are real but small UX gaps, and the top one is span replacement dropping underline without saying so.
