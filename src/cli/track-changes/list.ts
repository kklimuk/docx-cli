import type { SectionProperties, TrackedChange } from "@core";
import { flattenParagraphs, readSectionProperties } from "@core";
import { TrackChanges } from "@core/track-changes";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { revisionGroups } from "./groups";

const HELP = `docx track-changes list — inventory every revision wrapper

Usage:
  docx track-changes list FILE [options]

Options:
  -h, --help    Show this help

Lists every revision wrapper with stable tcN ids — run-level <w:ins>, <w:del>,
<w:moveFrom>, <w:moveTo>; section-property revisions <w:sectPrChange>; and
paragraph-mark <w:ins>/<w:del> markers (a self-closing element inside
<w:pPr><w:rPr> that tracks a paragraph break itself). moveFrom/moveTo halves
of the same logical move appear as separate entries; their kind tells them
apart.

Output: a bare JSON array of { id, kind, author, date, revisionId, blockId,
text } sorted by id (document order). Each item's "id" (e.g. tc0) is its
addressable handle — pass it to \`accept\`/\`reject --at tcN\`. When a del and an
ins are an adjacent REPLACE pair on the same paragraph, BOTH carry a shared
"group": "revN" — accept/reject the whole logical change in one call with
\`--at revN\` instead of accepting each half separately (tcN ids renumber after
each single accept, so the revN handle avoids the re-list ping-pong). Errors print
{code, error, hint?} with a nonzero exit. kind is one of: "ins", "del", "moveFrom",
"moveTo", "sectPrChange", "rowIns", "rowDel", "cellIns", "cellDel",
"tblGridChange", "tblPrChange", "tcPrChange", "checkboxToggle". Paragraph-mark
entries have kind "ins"/"del" with text "" — their blockId is the owning
paragraph's pN. Table-structural entries (rowIns/rowDel/cellIns/cellDel and
the property revisions tblGridChange/tblPrChange/tcPrChange) have text "" and
blockId set to the owning table's tN. checkboxToggle entries surface a Word
checkbox content control's tracked toggle (☐↔☒): metadata comes from the
inner <w:ins> (the new glyph); reject restores the prior glyph and flips
the w14:checked attribute back. Structural inserts/deletes of a checkbox
(Word emits <w:customXmlDel/InsRangeStart/End> around the SDT) round-trip
through the XmlNode tree but aren't yet enumerated as a dedicated kind.

Section-property revisions (kind="sectPrChange") additionally include
{ prior, current } objects with the section's columns/sectionType from
before and after the tracked edit, so agents can see the diff without
re-reading XML.

Examples:
  docx track-changes list doc.docx
  docx track-changes list doc.docx | jq '.[] | select(.kind == "del")'
  docx track-changes list doc.docx | jq '.[] | select(.kind | test("move"))'
  docx track-changes list doc.docx | jq '.[] | select(.kind == "sectPrChange") | .prior, .current'
`;

type TrackedChangeRecord = TrackedChange & {
	blockId: string;
	text: string;
	prior?: SectionProperties;
	current?: SectionProperties;
	/** `revN` when this change is one half of a del+ins replace pair; absent for
	 *  solo changes. `accept/reject --at revN` acts on both halves at once. */
	group?: string;
};

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			help: { type: "boolean", short: "h" },
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

	const byId = new Map<string, TrackedChangeRecord>();
	for (const paragraph of flattenParagraphs(document.body.blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text" || !run.trackedChange) continue;
			const change = run.trackedChange;
			const existing = byId.get(change.id);
			if (existing) {
				existing.text += run.text;
				continue;
			}
			byId.set(change.id, {
				...change,
				blockId: paragraph.id,
				text: run.text,
			});
		}
	}

	// The AST loop above only sees text-bearing run-level changes. Everything
	// else — empty wrappers (e.g. <w:ins> wrapping only <w:del>), section /
	// table-property revisions, checkbox toggles, and standalone note-body
	// edits — comes from the single tracked-change inventory the reader built
	// (TrackChanges.list reads document.trackedChangeReferences; no re-walk). kind,
	// author, date and revisionId are already resolved on each record.
	for (const change of new TrackChanges(document).list()) {
		if (byId.has(change.id)) continue;
		const record: TrackedChangeRecord = {
			id: change.id,
			kind: change.kind,
			author: change.author,
			date: change.date,
			revisionId: change.revisionId,
			blockId: change.blockId,
			text: "",
		};
		if (change.kind === "sectPrChange") {
			// Live siblings (parent array) carry the post-edit values; the
			// snapshot inside the change marker carries the prior values.
			const liveSiblings = change.parent.filter(
				(child) => child !== change.node,
			);
			record.current = readSectionProperties(liveSiblings);
			const snapshot = change.node.findChild("w:sectPr");
			record.prior = snapshot ? readSectionProperties(snapshot.children) : {};
		}
		byId.set(change.id, record);
	}

	const sorted = [...byId.values()].sort(
		(a, b) => trackedChangeIndex(a.id) - trackedChangeIndex(b.id),
	);

	// Tag the two halves of each del+ins replace with a shared `revN` so an agent
	// can accept/reject the logical change in ONE call (`accept --at revN`) instead
	// of the id-renumbering ping-pong of accepting each half separately.
	const { revOf } = revisionGroups(sorted);
	for (const record of sorted) {
		const group = revOf.get(record.id);
		if (group) record.group = group;
	}

	await respond(sorted);
	return EXIT.OK;
}

function trackedChangeIndex(id: string): number {
	const match = id.match(/^tc(\d+)$/);
	return match?.[1] ? Number(match[1]) : 0;
}
