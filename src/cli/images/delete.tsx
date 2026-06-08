import { convertTextToDelText, Del, type Document, TrackChanges } from "@core";
import { Images } from "@core/image";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx images delete — remove an embedded image

Usage:
  docx images delete FILE --at imgN [options]

Required:
  --at imgN         image id to remove (e.g. img0)

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

Output:
  Silent on success (exit 0) — delete mints no new id. --verbose prints
  {ok:true, operation, path, imageId, partName, pruned}. Errors print
  {code, error, hint?} with a nonzero exit. Discover ids with
  \`docx images list FILE\`.

Examples:
  docx images delete doc.docx --at img0
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			author: { type: "string" },
			track: { type: "boolean" },
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
	if (!targetId) return fail("USAGE", "Missing --at imgN", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const reference = document.body.imageById.get(targetId);
	if (!reference) {
		return fail("IMAGE_NOT_FOUND", `Image not found: ${targetId}`);
	}

	// `imgN` is positional in document order; collectImageRuns yields hits in the
	// same order `read` assigns ids, so hit[N] is imgN (two ids can share one
	// media part / relationship, so we resolve the occurrence, not just the rId).
	const occurrences = new Images(document).list();
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
	if (resolveTracked(document, parsed.values.track)) {
		// Real tracked deletion: wrap the run in <w:del>. Keep the media part —
		// rejecting the change restores the image, so pruning now would orphan it.
		const meta = new TrackChanges(document).mintMeta(
			parsed.values.author as string | undefined,
		);
		const deleted = (
			<Del meta={meta}>{convertTextToDelText(occurrence.run)}</Del>
		);
		occurrence.parent.splice(runIndex, 1, deleted);
	} else {
		occurrence.parent.splice(runIndex, 1);
		document.body.imageById.delete(targetId);
		pruned = pruneIfUnreferenced(document, occurrence.relationshipId);
	}

	await document.save(outputPath);

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
function pruneIfUnreferenced(
	document: Document,
	relationshipId: string,
): boolean {
	const reference =
		document.relationships.imagesByRelationshipId.get(relationshipId);
	const pruned = document.relationships.removeIfUnreferenced(
		relationshipId,
		document.documentTree,
	);
	if (!pruned) return false;
	if (reference && document.pkg.hasPart(reference.partName)) {
		document.pkg.deletePart(reference.partName);
	}
	return true;
}
