import { join } from "node:path";
import {
	type Block,
	enrichImageHashes,
	type ImageRun,
	openDocView,
	PkgError,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx images extract — dump image bytes to a directory

Usage:
  docx images extract FILE --to DIR [options]

Required:
  --to DIR          Output directory (created if missing)

Optional:
  --id IMG_ID       Extract a single image (default: extract all)
  -h, --help        Show this help

Files are named <hash>.<ext> where hash is the sha256 of the image bytes
and ext is derived from contentType. Returns a manifest mapping image ids
to written paths.

Examples:
  docx images extract doc.docx --to ./media
  docx images extract doc.docx --to ./media --id img2
`;

const EXTENSION_BY_CONTENT_TYPE: Record<string, string> = {
	"image/png": "png",
	"image/jpeg": "jpg",
	"image/jpg": "jpg",
	"image/gif": "gif",
	"image/svg+xml": "svg",
	"image/webp": "webp",
	"image/tiff": "tif",
	"image/bmp": "bmp",
	"image/x-emf": "emf",
	"image/x-wmf": "wmf",
	"image/vnd.microsoft.icon": "ico",
};

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				to: { type: "string" },
				id: { type: "string" },
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

	const outputDir = parsed.values.to as string | undefined;
	if (!outputDir) return fail("USAGE", "Missing --to DIR", HELP);

	const targetId = parsed.values.id as string | undefined;

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

	const allImages: ImageRun[] = [];
	collectImages(view.doc.blocks, allImages);

	const targets = targetId
		? allImages.filter((image) => image.id === targetId)
		: allImages;

	if (targetId && targets.length === 0) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	const manifest: { id: string; path: string; bytes: number }[] = [];
	const seenHashes = new Set<string>();
	for (const image of targets) {
		const reference = view.imageById.get(image.id);
		if (!reference) continue;
		const extension = extensionFor(image.contentType, reference.partName);
		const fileName = `${image.hash}.${extension}`;
		const outputPath = join(outputDir, fileName);

		if (!seenHashes.has(image.hash)) {
			const bytes = await view.pkg.readBytes(reference.partName);
			await Bun.write(outputPath, bytes);
			seenHashes.add(image.hash);
		}
		manifest.push({
			id: image.id,
			path: outputPath,
			bytes: (await Bun.file(outputPath).arrayBuffer()).byteLength,
		});
	}

	await respond({ ok: true, operation: "images.extract", path, manifest });
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

function extensionFor(contentType: string, partName: string): string {
	const fromType = EXTENSION_BY_CONTENT_TYPE[contentType.toLowerCase()];
	if (fromType) return fromType;
	const fromName = partName.split(".").pop()?.toLowerCase();
	return fromName ?? "bin";
}
