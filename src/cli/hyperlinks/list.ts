import { type Block, type Hyperlink, iterateBlocks } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx hyperlinks list — print hyperlink manifest as JSON

Usage:
  docx hyperlinks list FILE [options]

Options:
  -h, --help        Show this help

Output:
  A bare JSON array of hyperlink objects: { id, url (or anchor), tooltip (if
  set), text (display text), blockId (the paragraph containing the link) }. Each
  item's "id" (e.g. link0) is its addressable handle — pass it to
  \`hyperlinks replace/delete --at\`. Errors print {code, error, hint?} with a
  nonzero exit.

Examples:
  docx hyperlinks list doc.docx | jq -c '.[] | {id, url, text}'
`;

type HyperlinkEntry = {
	id: string;
	url?: string;
	anchor?: string;
	tooltip?: string;
	text: string;
	blockId: string;
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

	const entries = new Map<string, HyperlinkEntry>();
	collectHyperlinks(document.body.blocks, entries);
	await respond([...entries.values()]);
	return EXIT.OK;
}

function collectHyperlinks(
	blocks: Block[],
	entries: Map<string, HyperlinkEntry>,
): void {
	for (const block of iterateBlocks(blocks)) {
		if (block.type !== "paragraph") continue;
		for (const run of block.runs) {
			if (run.type !== "text" || !run.hyperlink) continue;
			addToEntry(entries, run.hyperlink, block.id, run.text);
		}
	}
}

function addToEntry(
	entries: Map<string, HyperlinkEntry>,
	hyperlink: Hyperlink,
	blockId: string,
	text: string,
): void {
	const existing = entries.get(hyperlink.id);
	if (existing) {
		existing.text += text;
		return;
	}
	const entry: HyperlinkEntry = {
		id: hyperlink.id,
		text,
		blockId,
	};
	if (hyperlink.url) entry.url = hyperlink.url;
	if (hyperlink.anchor) entry.anchor = hyperlink.anchor;
	if (hyperlink.tooltip) entry.tooltip = hyperlink.tooltip;
	entries.set(hyperlink.id, entry);
}
