import {
	type Block,
	enrichImageHashes,
	type ImageRun,
	openDocView,
	PkgError,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx images list — print image manifest as JSON

Usage:
  docx images list FILE [options]

Options:
  -h, --help        Show this help

Examples:
  docx images list doc.docx | jq -c '.[] | {id, contentType, hash}'
`;

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

	await enrichImageHashes(view);

	const images: ImageRun[] = [];
	collectImages(view.doc.blocks, images);
	await respond(images);
	return EXIT.OK;
}

function collectImages(blocks: Block[], out: ImageRun[]): void {
	for (const block of blocks) {
		if (block.type === "paragraph") {
			for (const run of block.runs) {
				if (run.type === "image") out.push(run);
			}
			continue;
		}
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					collectImages(cell.blocks, out);
				}
			}
		}
	}
}
