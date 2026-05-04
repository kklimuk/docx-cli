import { enrichImageHashes, openDocView, PkgError } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx read — print AST as JSON

Usage:
  docx read FILE [options]

Options:
  --text       Render as human-readable text instead of JSON (NOT YET IMPLEMENTED)
  -h, --help   Show this help

Examples:
  docx read input.docx
  docx read input.docx | jq '.blocks[] | select(.type == "paragraph")'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				text: { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (e) {
		return fail("USAGE", e instanceof Error ? e.message : String(e), HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (e) {
		if (e instanceof PkgError) {
			if (e.code === "FILE_NOT_FOUND") return fail("FILE_NOT_FOUND", e.message);
			if (e.code === "NOT_A_ZIP") return fail("NOT_A_ZIP", e.message);
		}
		throw e;
	}

	await enrichImageHashes(view);
	await respond(view.doc);
	return EXIT.OK;
}
