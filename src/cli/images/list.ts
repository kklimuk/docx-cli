import { type Block, enrichImageHashes, type ImageRun } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

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
