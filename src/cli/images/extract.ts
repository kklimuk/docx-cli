import { join } from "node:path";
import { flattenImageRuns } from "@core";
import { extensionForImageMime, Images } from "@core/image";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx images extract — dump image bytes to a directory

Usage:
  docx images extract FILE --to DIR [options]

Required:
  --to DIR          Output directory (created if missing)

Optional:
  --at imgN         Extract a single image (e.g., img0; default: extract all)
  -h, --help        Show this help

Files are named <hash>.<ext> where hash is the sha256 of the image bytes
and ext is derived from contentType.

Output:
  Prints a bare JSON array — the manifest, one entry per extracted image:
  {id, path, bytes}. Errors print {code, error, hint?} with a nonzero exit.
  Discover ids with \`docx images list FILE\`.

Examples:
  docx images extract doc.docx --to ./media
  docx images extract doc.docx --to ./media --at img2
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			to: { type: "string" },
			at: { type: "string" },
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

	const outputDir = parsed.values.to as string | undefined;
	if (!outputDir) return fail("USAGE", "Missing --to DIR", HELP);

	const targetId = parsed.values.at as string | undefined;

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	await new Images(document).enrichHashes();

	const allImages = flattenImageRuns(document.body.blocks);

	const targets = targetId
		? allImages.filter((image) => image.id === targetId)
		: allImages;

	if (targetId && targets.length === 0) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	const manifest: { id: string; path: string; bytes: number }[] = [];
	const seenHashes = new Set<string>();
	for (const image of targets) {
		const reference = document.body.imageById.get(image.id);
		if (!reference) continue;
		const extension = extensionFor(image.contentType, reference.partName);
		const fileName = `${image.hash}.${extension}`;
		const outputPath = join(outputDir, fileName);

		if (!seenHashes.has(image.hash)) {
			const bytes = await document.pkg.readBytes(reference.partName);
			await Bun.write(outputPath, bytes);
			seenHashes.add(image.hash);
		}
		manifest.push({
			id: image.id,
			path: outputPath,
			bytes: (await Bun.file(outputPath).arrayBuffer()).byteLength,
		});
	}

	await respond(manifest);
	return EXIT.OK;
}

function extensionFor(contentType: string, partName: string): string {
	const fromType = extensionForImageMime(contentType);
	if (fromType) return fromType;
	const fromName = partName.split(".").pop()?.toLowerCase();
	return fromName ?? "bin";
}
