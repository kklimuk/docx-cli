import type { TrackedChange, TrackedChangeKind } from "@core";
import { flattenParagraphs } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

const HELP = `docx track-changes list — inventory every revision wrapper

Usage:
  docx track-changes list FILE [options]

Options:
  -h, --help    Show this help

Lists every <w:ins>, <w:del>, <w:moveFrom>, and <w:moveTo> wrapper with
stable tcN ids. moveFrom/moveTo halves of the same logical move appear
as separate entries (one for each side); their kind tells them apart.

Output: JSON array of { id, kind, author, date, revisionId, blockId, text }
sorted by id (document order). kind is one of: "ins", "del", "moveFrom",
"moveTo".

Examples:
  docx track-changes list doc.docx
  docx track-changes list doc.docx | jq '.[] | select(.kind == "del")'
  docx track-changes list doc.docx | jq '.[] | select(.kind | test("move"))'
`;

type TrackedChangeRecord = TrackedChange & {
	blockId: string;
	text: string;
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
		const kind = trackedChangeKindForTag(reference.node.tag);
		if (!kind) continue;
		byId.set(id, {
			id,
			kind,
			author: reference.node.getAttribute("w:author") ?? "",
			date: reference.node.getAttribute("w:date") ?? "",
			revisionId: reference.node.getAttribute("w:id") ?? "",
			blockId: reference.blockId,
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
	return null;
}
