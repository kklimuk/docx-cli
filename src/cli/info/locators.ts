import { describeForms, LOCATOR_FORMS, type LocatorFormKey } from "@core";
import { EXIT, tryParseArgs, writeStdout } from "../respond";

const HELP = `docx info locators — print the locator grammar reference

Usage:
  docx info locators [options]

Options:
  --json       Print as JSON
  -h, --help   Show this help

Examples:
  docx info locators
  docx info locators --json | jq '.entityLocators'
`;

// One source of truth: the form groups below name keys in @core's LOCATOR_FORMS
// registry, and BOTH the text reference (renderReference) and the JSON reference
// (renderJson, keyed by `jsonKey`) are rendered from it — so this doc can't drift
// from the parser, from each command's help, or between its own two projections.
const GROUPS: { title: string; jsonKey: string; keys: LocatorFormKey[] }[] = [
	{
		title: "Block locators",
		jsonKey: "blockLocators",
		keys: ["paragraph", "table", "section", "cellParagraph"],
	},
	{
		title: "Span locators (characters within a single paragraph)",
		jsonKey: "spanLocators",
		keys: ["span", "cellSpan"],
	},
	{
		title: "Range locators (across blocks)",
		jsonKey: "rangeLocators",
		keys: ["blockRange", "crossSpan"],
	},
	{
		title: "Entity locators",
		jsonKey: "entityLocators",
		keys: [
			"comment",
			"image",
			"hyperlink",
			"equation",
			"trackedChange",
			"footnote",
			"endnote",
			"header",
			"footer",
			"cell",
		],
	},
	{
		title: 'Table-structure locators (for the "docx tables" verbs)',
		jsonKey: "tableStructureLocators",
		keys: ["tableRow", "tableColumn", "cellRange"],
	},
];

function renderReference(): string {
	const sections = GROUPS.map((group) => {
		const lines = describeForms(group.keys, "  ");
		return `${group.title}:\n${lines}`;
	}).join("\n\n");

	const examples = [
		"paragraph",
		"span",
		"blockRange",
		"crossSpan",
		"section",
		"comment",
		"footnote",
		"endnote",
		"cell",
		"cellParagraph",
		"cellSpan",
	]
		.map((key) => {
			const form = LOCATOR_FORMS[key as LocatorFormKey];
			return `  ${form.example.padEnd(15)} -> ${form.summary}`;
		})
		.join("\n");

	return `LOCATOR GRAMMAR

${sections}

How a locator is passed:
  --at LOCATOR            address an existing thing to act on (edit, delete,
                         comments, footnotes/endnotes, images, hyperlinks,
                         tables, track-changes); insert --at accepts a bare cell
  --after / --before     where to place a new block, or a block at a cell's
                         start/end boundary (insert)
  --from / --to          a top-level block slice to render (read)
  --anchor PHRASE        find text, then anchor to it (comments add)
  positional [LOCATOR]   an optional slice to scope (wc)

Notation:
  Uppercase letters are numeric indices (N, R, C, S, E, K). Character offsets
  are 0-based, start inclusive, end exclusive (p3:5-20 is chars 5..19).
  Don't hand-count offsets — run \`docx find FILE "phrase"\` to get the exact
  span locator (e.g. p3:5-20) and paste it straight into --at.
  eqN surfaces run.latex (reconstructed LaTeX) + run.display (inline vs $$…$$).
  tcN covers run-level <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>, <w:sectPrChange>,
  table-structural revisions, and paragraph-mark ins/del.
  Cell locators chain to any depth: t0:r2c1:t0:r0c0:p0 is the first paragraph of
  cell (0,0) of the first table nested inside cell (2,1) of table t0.
  A bare CELL content target is conservative: edit requires exactly one direct
  paragraph; insert rejects merged/grid-shifted cells. Use CELL:pK for a complex
  cell or exact paragraph targeting.

Table locators are one family — each "docx tables" verb takes the shape that
matches what it reshapes (they look different because they address different
things, not because the grammar is inconsistent):
  tN:rR              a whole row     (insert-row position, delete-row)
  tN:cC              a whole column  (insert-column position, delete-column)
  tN:rRcC            a single cell   (unmerge; edit --at CELL; insert --at,
                                      --before, or --after CELL when simple/unmerged)
  tN:rR1cC1-rR2cC2   a rectangular region, top-left (R1,C1) → bottom-right
                     (R2,C2)  (merge)
  tN                 the whole table (set-widths, borders, insert-row/column)

Examples:
${examples}

Discovering ids:
  pN / tN / sN     docx read FILE --ast   (or the <!-- pN --> comments in the
                   default Markdown render)
  cN               docx comments list FILE
  imgN             docx images list FILE
  linkN            docx hyperlinks list FILE
  fnN / enN        docx footnotes list FILE / docx endnotes list FILE
  hdrN / ftrN      docx headers list FILE / docx footers list FILE
  tcN              docx track-changes list FILE
  eqN              docx read FILE --ast   (EquationRun.id)

Notes:
  Block ids are positional and shift after structural edits — re-read between
  non-trivial mutations. Commands that mint a new id print it (one per line);
  re-read to recover any others.
`;
}

function renderJson(): unknown {
	const group = (keys: LocatorFormKey[]) =>
		Object.fromEntries(
			keys.map((key) => [
				key,
				{
					syntax: LOCATOR_FORMS[key].syntax,
					example: LOCATOR_FORMS[key].example,
					summary: LOCATOR_FORMS[key].summary,
				},
			]),
		);

	return {
		// Derived from GROUPS (same list the text reference renders), so a new
		// LOCATOR_FORMS entry lands in both projections from one edit.
		...Object.fromEntries(
			GROUPS.map((formGroup) => [formGroup.jsonKey, group(formGroup.keys)]),
		),
		notation: {
			placeholders: "Uppercase letters are numeric indices (N, R, C, S, E, K).",
			offsets: "Character offsets are 0-based, start inclusive, end exclusive.",
		},
		flags: {
			"--at":
				"address an existing thing to act on; insert --at accepts a bare simple cell",
			"--after / --before":
				"where to place a new block, or at a bare cell's start/end boundary (insert)",
			"--from / --to": "a top-level block slice to render (read)",
			"--anchor": "find text, then anchor to it (comments add)",
			"positional [LOCATOR]": "an optional slice to scope (wc)",
		},
		discoveringIds: {
			"pN/tN/sN": "docx read FILE --ast",
			cN: "docx comments list FILE",
			imgN: "docx images list FILE",
			linkN: "docx hyperlinks list FILE",
			"fnN/enN": "docx footnotes list FILE / docx endnotes list FILE",
			"hdrN/ftrN": "docx headers list FILE / docx footers list FILE",
			tcN: "docx track-changes list FILE",
			eqN: "docx read FILE --ast (EquationRun.id)",
		},
		notes: [
			"Block ids are positional and shift after structural edits.",
			"Re-read between non-trivial mutations.",
			"Commands that mint a new id print it (one per line) so you can chain.",
			"Bare-cell edit requires one direct paragraph; bare-cell insert rejects merged/grid-shifted cells. Use CELL:pK for precision.",
		],
	};
}

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	if (parsed.values.json) {
		await writeStdout(`${JSON.stringify(renderJson(), null, 2)}\n`);
		return EXIT.OK;
	}

	await writeStdout(renderReference());
	return EXIT.OK;
}
