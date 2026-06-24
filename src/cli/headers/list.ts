import type { MarginalKind } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { marginalNoun } from "./shared";

function helpFor(kind: MarginalKind): string {
	const noun = marginalNoun(kind);
	const prefix = kind === "header" ? "hdr" : "ftr";
	return `docx ${noun} list — print existing ${noun} as JSON

Usage:
  docx ${noun} list FILE [options]

Options:
  -h, --help          Show this help

Output:
  A bare JSON array, one object per ${kind} reference:
  { id, kind, type, sectionId, text }. \`id\` (e.g. ${prefix}0) is the positional
  handle; \`type\` is default | first | even; \`text\` shows fields as tokens
  ({page}, {pages}, {date}, {time}, {styleref:NAME}, {filename}, {title}, {author};
  {time} is read-only — no authoring flag mints it).
  A document-wide ${kind} appears once per section it governs. Empty array if none.

Examples:
  docx ${noun} list doc.docx
  docx ${noun} list doc.docx | jq '.[] | select(.type=="first")'
`;
}

export async function runListMarginals(
	args: string[],
	kind: MarginalKind,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{ help: { type: "boolean", short: "h" } },
		helpFor(kind),
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(helpFor(kind));
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", helpFor(kind));

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const list =
		kind === "header" ? document.body.headers : document.body.footers;
	await respond(list);
	return EXIT.OK;
}
