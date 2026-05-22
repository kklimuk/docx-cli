import { join } from "node:path";
import { enrichImageHashes, flattenImageRuns } from "@core";
import { extensionForImageMime } from "@core/image";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	await enrichImageHashes(view);

	const allImages = flattenImageRuns(view.doc.blocks);

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

function extensionFor(contentType: string, partName: string): string {
	const fromType = extensionForImageMime(contentType);
	if (fromType) return fromType;
	const fromName = partName.split(".").pop()?.toLowerCase();
	return fromName ?? "bin";
}
