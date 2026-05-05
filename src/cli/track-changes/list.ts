import type { TrackedChange } from "@core";
import { flattenParagraphs } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

const HELP = `docx track-changes list — list every <w:ins>/<w:del> with metadata

Usage:
  docx track-changes list FILE [options]

Options:
  -h, --help    Show this help

Output: JSON array of { id, kind, author, date, revisionId, blockId, text }
sorted by id (document order).

Examples:
  docx track-changes list doc.docx
  docx track-changes list doc.docx | jq '.[] | select(.kind == "del")'
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
		byId.set(id, {
			id,
			kind: reference.node.tag === "w:ins" ? "ins" : "del",
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
