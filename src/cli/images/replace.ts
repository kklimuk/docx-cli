import {
	isTrackChangesEnabled,
	resolveAuthor,
	resolveDate,
	saveDocView,
} from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { emitAuditComment, findContainingParagraph } from "../comments/helpers";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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

const EXTENSION_BY_MIME: Record<string, string> = {
	"image/png": "png",
	"image/jpeg": "jpeg",
	"image/gif": "gif",
	"image/webp": "webp",
	"image/bmp": "bmp",
	"image/tiff": "tiff",
	"image/svg+xml": "svg",
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
				at: { type: "string" },
				with: { type: "string" },
				author: { type: "string" },
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

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at IMG_ID", HELP);

	const sourcePath = parsed.values.with as string | undefined;
	if (!sourcePath) return fail("USAGE", "Missing --with PATH", HELP);

	const sourceFile = Bun.file(sourcePath);
	if (!(await sourceFile.exists())) {
		return fail("FILE_NOT_FOUND", `Replacement file not found: ${sourcePath}`);
	}

	const newMimeType = sourceFile.type;
	const newExtension = EXTENSION_BY_MIME[newMimeType];
	if (!newExtension) {
		return fail(
			"USAGE",
			`Unsupported replacement image type: ${newMimeType}`,
			"Supported: png, jpeg, gif, webp, bmp, tiff, svg, emf, wmf, ico.",
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = view.imageById.get(targetId);
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
		view.pkg.writeBytes(originalPartName, bytes);
	} else {
		view.pkg.writeBytes(newPartName, bytes);
		view.pkg.deletePart(originalPartName);
		updateRelationshipTarget(
			view.relationshipsTree,
			reference.relationshipId,
			relativeTargetFor(newPartName),
		);
		ensureContentTypeDefault(view.contentTypesTree, newExtension, newMimeType);
		reference.partName = newPartName;
		reference.contentType = newMimeType;
	}

	if (isTrackChangesEnabled(view)) {
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		const body = `[docx-cli] image replaced: ${originalPartName} (${originalMimeType}) → ${newPartName} (${newMimeType}, ${bytes.length} bytes)`;
		const drawingRuns = findDrawingRunsForRelationship(
			view.documentTree,
			reference.relationshipId,
		);
		for (const drawingRun of drawingRuns) {
			const paragraph = findContainingParagraph(view.documentTree, drawingRun);
			if (!paragraph) continue;
			emitAuditComment(
				view,
				{ kind: "run", paragraph, run: drawingRun },
				{ body, author, date },
			);
		}
	}

	await saveDocView(view, outputPath);

	await respond({
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

function findDrawingRunsForRelationship(
	documentTree: XmlNode[],
	relationshipId: string,
): XmlNode[] {
	const matches: XmlNode[] = [];
	function walk(node: XmlNode): void {
		if (node.tag === "w:r" && runReferencesImage(node, relationshipId)) {
			matches.push(node);
			return;
		}
		for (const child of node.children) walk(child);
	}
	for (const root of documentTree) walk(root);
	return matches;
}

function runReferencesImage(run: XmlNode, relationshipId: string): boolean {
	for (const child of run.children) {
		if (child.tag !== "w:drawing") continue;
		const blip = child.findDescendant("a:blip");
		if (!blip) continue;
		const embed = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link");
		if (embed === relationshipId) return true;
	}
	return false;
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

function updateRelationshipTarget(
	relationshipsTree: XmlNode[],
	relationshipId: string,
	newTarget: string,
): void {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (!relationships) return;
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		if (child.getAttribute("Id") === relationshipId) {
			child.setAttribute("Target", newTarget);
			return;
		}
	}
}

function ensureContentTypeDefault(
	contentTypesTree: XmlNode[],
	extension: string,
	mimeType: string,
): void {
	const types = XmlNode.findRoot(contentTypesTree, "Types");
	if (!types) return;
	for (const child of types.children) {
		if (child.tag !== "Default") continue;
		if (child.getAttribute("Extension")?.toLowerCase() === extension) return;
	}
	types.children.push(
		new XmlNode("Default", { Extension: extension, ContentType: mimeType }),
	);
}
