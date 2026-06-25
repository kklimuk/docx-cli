import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import {
	collectTrackedChangeRecords,
	renderTrackedChangeTable,
} from "./list-view";

const HELP = `docx track-changes list — inventory every revision wrapper

Usage:
  docx track-changes list FILE [options]

Options:
  --json        Emit the full revision objects as a JSON array (default: a
                text table, one LOGICAL change per line)
  -h, --help    Show this help

Lists every revision wrapper with stable tcN ids — run-level <w:ins>, <w:del>,
<w:moveFrom>, <w:moveTo>; section-property revisions <w:sectPrChange>; and
paragraph-mark <w:ins>/<w:del> markers (a self-closing element inside
<w:pPr><w:rPr> that tracks a paragraph break itself). moveFrom/moveTo halves
of the same logical move appear as separate entries; their kind tells them
apart.

Output: a text table, ONE LOGICAL CHANGE per line —

  rev0  replace  p7   "Net 90 from the Company" → "Net 30 from the Company"
  tc4   format   p22  spacing.line →360
  tc7   delete   p27  "Personal Guarantee."

The leftmost token is the handle you pass to \`accept\`/\`reject --at\` (repeatable,
e.g. \`--at rev0 --at tc4\`). A del+ins REPLACE pair on the same paragraph is
collapsed onto ONE line under its shared revN handle, so accepting/rejecting the
whole logical change is a single call — addressing the two tcN halves separately
forces a re-list, because tcN ids renumber after each accept. All --at targets in
one call resolve against the pre-mutation tree, so the renumbering never bites a
batch.

--json gives a bare JSON array of { id, kind, author, date, revisionId, blockId,
text } sorted by id (document order). Each item's "id" (e.g. tc0) is its granular
handle; paired halves additionally carry "group": "revN". kind is one of: "ins",
"del", "moveFrom", "moveTo", "sectPrChange", "pPrChange", "rowIns", "rowDel",
"cellIns", "cellDel", "tblGridChange", "tblPrChange", "tcPrChange",
"checkboxToggle". Paragraph-mark entries have kind "ins"/"del" with text "" —
their blockId is the owning paragraph's pN. Table-structural entries
(rowIns/rowDel/cellIns/cellDel and the property revisions
tblGridChange/tblPrChange/tcPrChange) have text "" and blockId set to the owning
table's tN. checkboxToggle entries surface a Word checkbox content control's
tracked toggle (☐↔☒): metadata comes from the inner <w:ins> (the new glyph);
reject restores the prior glyph and flips the w14:checked attribute back.
Structural inserts/deletes of a checkbox (Word emits
<w:customXmlDel/InsRangeStart/End> around the SDT) round-trip through the
XmlNode tree but aren't yet enumerated as a dedicated kind.

Property revisions (kind="sectPrChange" or "pPrChange") additionally include
{ prior, current } objects from before and after the tracked edit, so agents
can see the diff without re-reading XML: sectPrChange carries the section's
columns/sectionType; pPrChange carries the paragraph's
style/alignment/spacing/indent.

Examples:
  docx track-changes list doc.docx
  docx track-changes accept doc.docx --at rev0 --at tc4
  docx track-changes list doc.docx --json | jq '.[] | select(.kind == "del")'
  docx track-changes list doc.docx --json | jq '.[] | select(.kind | test("move"))'
  docx track-changes list doc.docx --json | jq '.[] | select(.kind == "pPrChange") | .prior, .current'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			help: { type: "boolean", short: "h" },
			json: { type: "boolean" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const records = collectTrackedChangeRecords(document);

	if (parsed.values.json) {
		await respond(records);
		return EXIT.OK;
	}

	await writeStdout(renderTrackedChangeTable(records));
	return EXIT.OK;
}
