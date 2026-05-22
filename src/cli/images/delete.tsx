import {
	convertTextToDelText,
	createRevisionAllocator,
	Del,
	type DocView,
	isRelationshipReferenced,
	isTrackChangesEnabled,
	removeRelationship,
	resolveAuthor,
	resolveDate,
	saveDocView,
	type TrackedMeta,
} from "@core";
import { collectImageRuns } from "@core/image";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

const HELP = `docx images delete — remove an embedded image

Usage:
  docx images delete FILE --at IMG_ID [options]

Required:
  --at IMG_ID       Existing image to remove (e.g., img0)

Optional:
  --author NAME     Author for the tracked deletion when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Removes the inline drawing (and its containing run). If the underlying media
part is no longer referenced, both it and the relationship are pruned.

When track-changes is on, the drawing's run is wrapped in <w:del> instead — a
real tracked deletion (Word reverts it on reject, removes it on accept). The
media part is always kept (rejecting must be able to restore the image); an
accepted deletion leaves it as a harmless unreferenced part rather than pruning.
Image ids are positional and shift after a delete — re-read before the next
mutation.

Examples:
  docx images delete doc.docx --at img0
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				author: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
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

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at IMG_ID", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = view.imageById.get(targetId);
	if (!reference) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	// `imgN` is positional in document order; collectImageRuns yields hits in the
	// same order `read` assigns ids, so hit[N] is imgN (two ids can share one
	// media part / relationship, so we resolve the occurrence, not just the rId).
	const occurrences = collectImageRuns(view);
	const ordinal = imageOrdinal(targetId);
	const occurrence = ordinal === null ? undefined : occurrences[ordinal];
	if (!occurrence) {
		return fail(
			"IMAGE_NOT_FOUND",
			`Image reference is stale (no drawing for ${targetId})`,
		);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "images.delete",
			dryRun: true,
			path,
			imageId: targetId,
			partName: reference.partName,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const runIndex = occurrence.parent.indexOf(occurrence.run);
	if (runIndex === -1) {
		return fail(
			"IMAGE_NOT_FOUND",
			`Image reference is stale (parent does not contain it): ${targetId}`,
		);
	}

	let pruned = false;
	if (isTrackChangesEnabled(view)) {
		// Real tracked deletion: wrap the run in <w:del>. Keep the media part —
		// rejecting the change restores the image, so pruning now would orphan it.
		const allocator = createRevisionAllocator(view);
		const meta: TrackedMeta = {
			author: resolveAuthor(parsed.values.author as string | undefined),
			date: resolveDate(),
			revisionId: allocator.next(),
		};
		const deleted = (
			<Del meta={meta}>{convertTextToDelText(occurrence.run)}</Del>
		);
		occurrence.parent.splice(runIndex, 1, deleted);
	} else {
		occurrence.parent.splice(runIndex, 1);
		view.imageById.delete(targetId);
		pruned = pruneIfUnreferenced(view, occurrence.relationshipId);
	}

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "images.delete",
		path: outputPath ?? path,
		imageId: targetId,
		partName: reference.partName,
		pruned,
	});
	return EXIT.OK;
}

function imageOrdinal(id: string): number | null {
	const match = id.match(/^img(\d+)$/);
	if (!match) return null;
	return Number(match[1]);
}

/** Drop the relationship and media part only when the rId is referenced NOWHERE
 * in the document — scanning every attribute, not just drawings, so we never
 * orphan a reference we don't model (a VML `<v:imagedata>` fallback sharing the
 * rId, an OLE object, a `<w:background>`). Better to leave a harmless orphan
 * part than to dangle an rId and corrupt the file. Returns whether we pruned. */
function pruneIfUnreferenced(view: DocView, relationshipId: string): boolean {
	if (isRelationshipReferenced(view.documentTree, relationshipId)) return false;

	const reference = view.imagesByRelationshipId.get(relationshipId);
	removeRelationship(view.relationshipsTree, relationshipId);
	if (reference && view.pkg.hasPart(reference.partName)) {
		view.pkg.deletePart(reference.partName);
	}
	view.imagesByRelationshipId.delete(relationshipId);
	return true;
}
