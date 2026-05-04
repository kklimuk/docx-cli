import {
	LocatorResolveError,
	openDocView,
	PkgError,
	resolveBlock,
	saveDocView,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx delete — remove a block at a locator

Usage:
  docx delete FILE [options]

Locator (required):
  --at LOCATOR      Block to remove (e.g., p3, t0)

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

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (openError) {
		if (openError instanceof PkgError) {
			if (openError.code === "FILE_NOT_FOUND") {
				return fail("FILE_NOT_FOUND", openError.message);
			}
			if (openError.code === "NOT_A_ZIP") {
				return fail("NOT_A_ZIP", openError.message);
			}
		}
		throw openError;
	}

	let blockRef: ReturnType<typeof resolveBlock>;
	try {
		blockRef = resolveBlock(view, at);
	} catch (resolveError) {
		if (resolveError instanceof LocatorResolveError) {
			return fail("BLOCK_NOT_FOUND", resolveError.message);
		}
		throw resolveError;
	}

	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "delete",
			dryRun: true,
			path,
			locator: at,
		});
		return EXIT.OK;
	}

	blockRef.parent.splice(targetIndex, 1);
	await saveDocView(view);

	await respond({ ok: true, operation: "delete", path, locator: at });
	return EXIT.OK;
}
