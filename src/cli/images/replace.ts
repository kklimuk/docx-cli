import { getPageContentHeightEmu, getPageContentWidthEmu } from "@core";
import { findContainingParagraph } from "@core/comments";
import {
	applyReplacementExtent,
	fitReplacementExtent,
	type ImageSource,
	ImageSourceError,
	Images,
	imageFormatForExtension,
	loadImageSource,
	type ReplacementSizing,
	SUPPORTED_IMAGE_FORMATS,
} from "@core/image";
import { addAuditComment } from "../audit-comment";
import type { RawValues } from "../insert/index";
import { formatImageSize } from "../read/annotations";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { parseImageSize } from "./add";

const HELP = `docx images replace — swap an image's picture, refitting its frame

Usage:
  docx images replace FILE --at imgN --with PATH [options]

Examples:
  docx images replace doc.docx --at img2 --with ./new-photo.png
  docx images replace doc.docx --at img0 --with ./diagram.jpg
  docx images replace doc.docx --at img0 --with ./logo.png --width 1.5in

Required:
  --at imgN         image id to replace (e.g. img0)
  --with PATH       New image file (format keyed off the file extension:
                    png/jpg/gif/webp/bmp/tiff/svg/emf/wmf/ico)

Sizing (default: keep the displayed WIDTH, recompute the height from the new
image's aspect ratio, scaled down to fit the text area if needed — the new
picture is never stretched into the old one's box):
  --width LENGTH    Display width: bare inches (2) or in/cm/mm/pt/px/tw.
                    Alone, the height follows the new image's aspect ratio.
  --height LENGTH   Display height, same units. Alone, the width follows.
                    Both given: honored as-is.
  --keep-size       Keep the old box exactly (only when the new image has the
                    same shape)

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

The replacement may use a different format from the original — the embedded
part is renamed and re-registered automatically. Any crop on the old picture
is dropped (it was cut for the old image). Every place that shows the same
picture (a shared media part) is refitted, each from its own width. A format
whose pixel size isn't read (emf/wmf/webp/bmp/tiff/ico) keeps the old box —
the confirmation says so; pass --width and --height to size it.

Image swaps can't be recorded as tracked changes, so when track-changes is on
an audit comment is anchored to each place the swapped image appears.

Output:
  Prints a one-line confirmation naming the new size on success (exit 0),
  e.g. "images.replace img0 (1x4in)" — replace mints no new id (the image
  keeps its imgN). --verbose prints {ok:true, operation, path, imageId, size,
  partName, mimeType, bytes}. Errors print {code, error, hint?} with
  a nonzero exit. Discover ids with \`docx images list FILE\`.
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			with: { type: "string" },
			width: { type: "string" },
			height: { type: "string" },
			"keep-size": { type: "boolean" },
			author: { type: "string" },
			...SAVE_FLAGS,
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at imgN", HELP);

	const sourcePath = parsed.values.with as string | undefined;
	if (!sourcePath) return fail("USAGE", "Missing --with PATH", HELP);

	const sizing = await parseSizing(parsed.values);
	if (typeof sizing === "number") return sizing;

	const sourceFile = Bun.file(sourcePath);
	if (!(await sourceFile.exists())) {
		return fail("FILE_NOT_FOUND", `Replacement file not found: ${sourcePath}`);
	}

	// Key off the file extension, not Bun's sniffed MIME — Bun reports
	// `image/emf`/`image/x-ms-bmp` where the docx (and extract) use
	// `image/x-emf`/`image/bmp`, so extension → canonical Word MIME is what lets
	// emf/wmf/bmp/ico round-trip.
	const sourceExtension = sourcePath.split(".").pop() ?? "";
	const format = imageFormatForExtension(sourceExtension);
	if (!format) {
		return fail(
			"USAGE",
			`Unsupported replacement image type: .${sourceExtension}`,
			`Supported: ${SUPPORTED_IMAGE_FORMATS}.`,
		);
	}
	const newMimeType = format.mimeType;
	const newExtension = format.extension;

	// Route the replacement through the shared resolver (same as `docx images add`)
	// so it gets the same safety: SVG sanitization (scripts/handlers/animation
	// stripped), HEIC→JPEG transcode, and the SSRF guard on remote srcs. We keep
	// the extension/MIME resolved above for part naming; the bytes and pixel
	// dimensions come back.
	let source: ImageSource;
	try {
		source = await loadImageSource(sourcePath);
	} catch (error) {
		if (error instanceof ImageSourceError) {
			return fail("IMAGE_SOURCE", error.message);
		}
		throw error;
	}
	const bytes = source.bytes;

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const reference = document.body.imageById.get(targetId);
	if (!reference) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	// The bytes are shared by every drawing on this relationship, so every one
	// of them is refitted (each from its own displayed width). Plan all fits
	// BEFORE any mutation, so a sizing error writes nothing.
	const occurrences = new Images(document).list();
	const textArea = {
		widthEmu: getPageContentWidthEmu(document),
		heightEmu: getPageContentHeightEmu(document),
	};
	const fits = [];
	for (const [index, occurrence] of occurrences.entries()) {
		if (occurrence.relationshipId !== reference.relationshipId) continue;
		const drawing = occurrence.run.findChild("w:drawing");
		if (!drawing) continue;
		const extent = fitReplacementExtent(drawing, source, sizing, textArea);
		if (!extent) {
			return fail(
				"USAGE",
				`Can't keep the aspect ratio: no pixel size could be read from .${sourceExtension}`,
				"Pass both --width and --height to size it explicitly.",
			);
		}
		fits.push({ id: `img${index}`, occurrence, drawing, extent });
	}
	const targetFit = fits.find((fit) => fit.id === targetId);
	if (!targetFit) {
		return fail(
			"IMAGE_NOT_FOUND",
			`Image reference is stale (no drawing for ${targetId})`,
		);
	}
	const size = formatImageSize(
		targetFit.extent.widthEmu,
		targetFit.extent.heightEmu,
	);
	const sizeNote = targetFit.extent.sizeKnown
		? undefined
		: `kept ${size}: size unknown for .${sourceExtension} — pass --width and --height if it's a different shape`;

	const originalPartName = reference.partName;
	const newPartName = renameExtension(originalPartName, newExtension);
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "images.replace",
			dryRun: true,
			path,
			imageId: targetId,
			size,
			from: { partName: originalPartName, mimeType: reference.contentType },
			to: { partName: newPartName, mimeType: newMimeType },
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	for (const fit of fits) applyReplacementExtent(fit.drawing, fit.extent);

	const originalMimeType = reference.contentType;

	if (newPartName === originalPartName) {
		document.pkg.writeBytes(originalPartName, bytes);
	} else {
		document.pkg.writeBytes(newPartName, bytes);
		document.relationships.setImageTarget(
			reference.relationshipId,
			relativeTargetFor(newPartName),
			newPartName,
			newMimeType,
		);
		// Repoint first, then delete the old part only if no OTHER relationship
		// still targets it — identical images are often deduped to one shared
		// part, and deleting it would dangle the sibling rId.
		if (
			!document.relationships.hasTarget(relativeTargetFor(originalPartName))
		) {
			document.pkg.deletePart(originalPartName);
		}
		document.contentTypes.registerExtension(newExtension, newMimeType);
		// Body's `imageById` reference (this `reference` IS the live entry) —
		// sibling view, so the call above can't sync it; we patch in place.
		reference.partName = newPartName;
		reference.contentType = newMimeType;
	}

	if (document.isTrackChangesEnabled()) {
		const authorFlag = parsed.values.author as string | undefined;
		const message = `image replaced: ${originalPartName} (${originalMimeType}) → ${newPartName} (${newMimeType}, ${bytes.length} bytes), now ${size}`;
		for (const { occurrence } of fits) {
			const drawingRun = occurrence.run;
			const paragraph = findContainingParagraph(
				document.documentTree,
				drawingRun,
			);
			if (!paragraph) continue;
			addAuditComment(
				document,
				{ kind: "run", paragraph, run: drawingRun },
				message,
				authorFlag,
			);
		}
	}

	await document.save(outputPath);

	await respondAck(
		{
			ok: true,
			operation: "images.replace",
			path: outputPath ?? path,
			imageId: targetId,
			size,
			...(sizeNote ? { note: sizeNote } : {}),
			partName: newPartName,
			mimeType: newMimeType,
			bytes: bytes.length,
		},
		sizeNote && `↳ ${sizeNote}`,
	);
	return EXIT.OK;
}

/** `--width`/`--height` (shared grammar with `images add`) or `--keep-size`,
 *  which can't be combined — it would silently win over an explicit size. */
async function parseSizing(
	values: RawValues,
): Promise<ReplacementSizing | number> {
	const size = await parseImageSize(values);
	if (typeof size === "number") return size;
	if (!values["keep-size"]) return size;
	if (size.widthInches !== undefined || size.heightInches !== undefined) {
		return fail(
			"USAGE",
			"--keep-size can't be combined with --width/--height",
			"Use --keep-size to keep the old box, or --width/--height to set a new one.",
		);
	}
	return { keepSize: true };
}

function renameExtension(partName: string, newExtension: string): string {
	const dotIndex = partName.lastIndexOf(".");
	const slashIndex = partName.lastIndexOf("/");
	if (dotIndex > slashIndex) {
		return `${partName.slice(0, dotIndex + 1)}${newExtension}`;
	}
	return `${partName}.${newExtension}`;
}

function relativeTargetFor(partName: string): string {
	return partName.startsWith("word/")
		? partName.slice("word/".length)
		: partName;
}
