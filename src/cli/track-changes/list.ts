import type {
	SectionProperties,
	TrackedChange,
	TrackedChangeKind,
} from "@core";
import { flattenParagraphs, readSectionProperties } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import { collectTrackedChanges } from "./apply";

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

Output: JSON array of { id, kind, author, date, revisionId, blockId, text }
sorted by id (document order). kind is one of: "ins", "del", "moveFrom",
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
};

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const byId = new Map<string, TrackedChangeRecord>();
	for (const paragraph of flattenParagraphs(view.doc.blocks)) {
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

	// Empty wrappers (e.g. <w:ins> containing only <w:del>) carry no text runs
	// so the loop above misses them. Pull them from the reference map so the
	// inventory stays in sync with what `resolveTrackedChange` can address.
	for (const [id, reference] of view.trackedChangeReferences) {
		if (byId.has(id)) continue;
		// Table-structural markers carry an explicit kind (their tag is
		// ambiguous: a row revision is a <w:ins>/<w:del> like a run-level one).
		const kind = reference.kind ?? trackedChangeKindForTag(reference.node.tag);
		if (!kind) continue;
		// For `checkboxToggle` the reference.node is the SDT; metadata lives on
		// the inner `<w:ins>` (the new glyph). For other kinds the node IS the
		// metadata-carrying element.
		const metadataNode =
			kind === "checkboxToggle"
				? (reference.node.findChild("w:sdtContent")?.findChild("w:ins") ??
					reference.node)
				: reference.node;
		const record: TrackedChangeRecord = {
			id,
			kind,
			author: metadataNode.getAttribute("w:author") ?? "",
			date: metadataNode.getAttribute("w:date") ?? "",
			revisionId: metadataNode.getAttribute("w:id") ?? "",
			blockId: reference.blockId,
			text: "",
		};
		if (kind === "sectPrChange") {
			// Live siblings (parent array) carry the post-edit values; the
			// snapshot inside the change marker carries the prior values.
			const liveSiblings = reference.parent.filter(
				(child) => child !== reference.node,
			);
			record.current = readSectionProperties(liveSiblings);
			const snapshot = reference.node.findChild("w:sectPr");
			record.prior = snapshot ? readSectionProperties(snapshot.children) : {};
		}
		byId.set(id, record);
	}

	// Body-only note revisions (footnote/endnote edits) aren't reachable from
	// the AST or the reference map (both walk document.xml only). Pull them
	// from the apply walker so `list` and `apply --at tcN` agree. Body-side
	// revisions paired to a doc-body reference are hidden by that walker, so
	// only standalone body-only edits show up here.
	for (const change of collectTrackedChanges(view)) {
		if (byId.has(change.id)) continue;
		byId.set(change.id, {
			id: change.id,
			kind: change.kind,
			author: change.author,
			date: change.date,
			revisionId: change.node.getAttribute("w:id") ?? "",
			blockId: "",
			text: "",
		});
	}

	const sorted = [...byId.values()].sort(
		(a, b) => trackedChangeIndex(a.id) - trackedChangeIndex(b.id),
	);

	await respond(sorted);
	return EXIT.OK;
}

function trackedChangeIndex(id: string): number {
	const match = id.match(/^tc(\d+)$/);
	return match?.[1] ? Number(match[1]) : 0;
}

function trackedChangeKindForTag(tag: string): TrackedChangeKind | null {
	if (tag === "w:ins") return "ins";
	if (tag === "w:del") return "del";
	if (tag === "w:moveFrom") return "moveFrom";
	if (tag === "w:moveTo") return "moveTo";
	if (tag === "w:sectPrChange") return "sectPrChange";
	return null;
}
