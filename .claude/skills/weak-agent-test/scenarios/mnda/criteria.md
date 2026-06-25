# Grading rubric — mnda (JUDGE ONLY)

The agent never sees this file. It is the ground-truth definition of "correct" for
the request in `task.md`. Judge the finished document against these checks.

## Pass conditions

- **Cover-page placeholders filled** with the values from `brief.md`: both party
  names, the signatories + titles, notice emails, the purpose, Effective Date
  (June 8, 2026), MNDA term (2 years), confidentiality term (3 years), Delaware
  governing law, New Castle County jurisdiction, and the signature-block dates. No
  `[ ... ]` bracket placeholders remain anywhere.
- **Signature table populated** for BOTH parties — Print Name, Title, Company,
  Notice Address, and Date.
- **Zero yellow highlight** anywhere in the document. The placeholders were
  highlighted; once filled, no highlight may remain (check the runs, not just the
  visual — a filled value that kept its `<w:highlight>` fails).
- **Font fidelity preserved** — the Georgia title, the bold field labels, the small
  gray sub-labels, and the Arial body all keep their original font, size, and color.
  The before/after render is identical EVERYWHERE except the filled-in values and the
  removed highlight. Any font/size/color drift on untouched text is a failure.

## How to verify

- `docx read FILE` (and `--ast` for run-level detail) — confirm the values landed and
  no `[...]` or highlight survives.
- Compare the BASELINE render against the OUTPUT render: only the intended cells
  changed; all other formatting is intact.
