import { resolveAuthor, resolveDate } from "@core";
import { Comments, findContainingParagraph } from "@core/comments";
import {
	Images,
	imageFormatForExtension,
	SUPPORTED_IMAGE_FORMATS,
} from "@core/image";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx images replace — swap an image's bytes

Usage:
  docx images replace FILE --at IMG_ID --with PATH [options]

Required:
  --at IMG_ID       Existing image to replace (e.g., img0)
  --with PATH       New image file (any image MIME type)

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

If the replacement uses a different format from the original, the part is
renamed (extension changes), the relationship Target is rewritten, and
[Content_Types].xml gets a Default entry for the new extension if needed.

When track-changes is on, an audit comment is anchored to each drawing that
referenced the swapped image since OOXML has no native tracked-change form
for image replacement.

Examples:
  docx images replace doc.docx --at img2 --with ./new-photo.png
  docx images replace doc.docx --at img0 --with ./diagram.svg
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			with: { type: "string" },
			author: { type: "string" },
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
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
	if (!targetId) return fail("USAGE", "Missing --at IMG_ID", HELP);

	const sourcePath = parsed.values.with as string | undefined;
	if (!sourcePath) return fail("USAGE", "Missing --with PATH", HELP);

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

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const reference = document.body.imageById.get(targetId);
	if (!reference) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	const originalPartName = reference.partName;
	const newPartName = renameExtension(originalPartName, newExtension);
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "images.replace",
			dryRun: true,
			path,
			imageId: targetId,
			from: { partName: originalPartName, mimeType: reference.contentType },
			to: { partName: newPartName, mimeType: newMimeType },
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const bytes = new Uint8Array(await sourceFile.arrayBuffer());
	const originalMimeType = reference.contentType;

	if (newPartName === originalPartName) {
		document.pkg.writeBytes(originalPartName, bytes);
	} else {
		document.pkg.writeBytes(newPartName, bytes);
		document.relationships.setTarget(
			reference.relationshipId,
			relativeTargetFor(newPartName),
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
		reference.partName = newPartName;
		reference.contentType = newMimeType;
	}

	if (document.isTrackChangesEnabled()) {
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		const body = `[docx-cli] image replaced: ${originalPartName} (${originalMimeType}) → ${newPartName} (${newMimeType}, ${bytes.length} bytes)`;
		const drawingRuns = new Images(document)
			.list()
			.filter((hit) => hit.relationshipId === reference.relationshipId)
			.map((hit) => hit.run);
		for (const drawingRun of drawingRuns) {
			const paragraph = findContainingParagraph(
				document.documentTree,
				drawingRun,
			);
			if (!paragraph) continue;
			new Comments(document).addAudit(
				{ kind: "run", paragraph, run: drawingRun },
				{ body, author, date },
			);
		}
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "images.replace",
		path: outputPath ?? path,
		imageId: targetId,
		partName: newPartName,
		mimeType: newMimeType,
		bytes: bytes.length,
	});
	return EXIT.OK;
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
