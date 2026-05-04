import { saveDocView } from "@core";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	writeStdout,
} from "../respond";

const HELP = `docx delete — remove a block at a locator

Usage:
  docx delete FILE [options]

Locator (required):
  --at LOCATOR      Block to remove (e.g., p3, t0)

  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be removed; do not write the file
  -h, --help        Show this help

Examples:
  docx delete doc.docx --at p3
  docx delete doc.docx --at t0
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
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

	const at = parsed.values.at as string | undefined;
	if (!at) return fail("USAGE", "Missing --at LOCATOR", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, at);
	if (typeof blockRef === "number") return blockRef;

	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "delete",
			dryRun: true,
			path,
			locator: at,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	blockRef.parent.splice(targetIndex, 1);
	await saveDocView(view, outputPath);

	await respond({
		ok: true,
		operation: "delete",
		path: outputPath ?? path,
		locator: at,
	});
	return EXIT.OK;
}
