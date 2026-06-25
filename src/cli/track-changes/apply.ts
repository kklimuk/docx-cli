import { TrackChanges, TrackedChangeNotFoundError } from "@core/track-changes";
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
import {
	expandRevisionTargets,
	revisionGroups,
	UnknownRevisionError,
} from "./groups";
import { remainingTrackedChangesBlock } from "./list-view";

const HELP = `docx track-changes apply — finalize: accept AND reject in one atomic call

Usage:
  docx track-changes apply FILE (--accept H ... | --reject H ...) [options]

A document review ends in a finalize: accept the changes you want, reject the
rest. Doing that as separate \`accept\` and \`reject\` calls is a trap — tcN/revN
ids renumber after every accept/reject, so the SECOND command addresses a
moved target ("tc5 not found", or worse, silently the wrong change). \`apply\`
takes BOTH decisions at once and resolves EVERY handle against the original
pre-mutation tree, so nothing renumbers mid-operation and the file is never
left half-finalized (there is no undo).

Handles are the same ones \`track-changes list\` prints: a tcN, or a revN that
covers both halves of a del+ins replace pair. Both flags repeat.

Targets (at least one required):
  --accept H        A handle (tcN or revN) to accept. Repeat for several
                    (--accept rev0 --accept rev1 --accept tc4).
  --reject H        A handle (tcN or revN) to reject. Repeat for several.

A handle may not appear in both lists. Unknown handles error before anything is
written. Leftover changes you name in neither list stay tracked (apply finalizes
only what you address); the confirmation re-lists them so you can see what's left.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0); when changes remain it
  also re-lists them with their renumbered handles. --verbose prints
  {ok:true, operation, path, applied}. --dry-run previews. Errors print
  {code, error, hint?} with a nonzero exit. Discover handles with
  \`docx track-changes list FILE\`.

Examples:
  docx track-changes apply doc.docx --accept rev0 --accept rev1 --accept tc4 \\
                                    --reject rev2 --reject tc7
  docx track-changes apply doc.docx --reject rev0 --dry-run
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			accept: { type: "string", multiple: true },
			reject: { type: "string", multiple: true },
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

	const acceptRaw = (parsed.values.accept as string[] | undefined) ?? [];
	const rejectRaw = (parsed.values.reject as string[] | undefined) ?? [];
	if (acceptRaw.length === 0 && rejectRaw.length === 0) {
		return fail("USAGE", "Specify --accept and/or --reject (handle)", HELP);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const trackChanges = new TrackChanges(document);
	// Expand any revN handles to their member tcNs against the ORIGINAL list, so
	// both decisions address the same pre-mutation numbering.
	const groups = revisionGroups(trackChanges.list());
	let accepts: string[];
	let rejects: string[];
	try {
		accepts = expandRevisionTargets(acceptRaw, groups);
		rejects = expandRevisionTargets(rejectRaw, groups);
	} catch (error) {
		if (error instanceof UnknownRevisionError) {
			return fail(
				"TRACKED_CHANGE_NOT_FOUND",
				error.message,
				"Run 'docx track-changes list FILE' — revN handles appear as the `group` field on paired changes.",
			);
		}
		throw error;
	}

	// A change can't be both accepted and rejected — catch it (including via a
	// revN whose half is named on the other side) before mutating anything.
	const rejectSet = new Set(rejects);
	const conflict = accepts.find((id) => rejectSet.has(id));
	if (conflict) {
		return fail(
			"USAGE",
			`${conflict} is named in both --accept and --reject`,
			HELP,
		);
	}

	const outputPath = parsed.values.output as string | undefined;

	try {
		if (parsed.values["dry-run"]) {
			// Sort the merged preview into DOCUMENT order (by tcN index) so the
			// preview matches the real apply ack, which applyResolvedTargets emits
			// in document order — otherwise the same call previews accepts-then-
			// rejects but acks in doc order, a confusing mismatch.
			const previewApplied = [
				...trackChanges.preview(accepts, "accept"),
				...trackChanges.preview(rejects, "reject"),
			].sort((a, b) => trackedChangeIndex(a.id) - trackedChangeIndex(b.id));
			await respond({
				operation: "track-changes.apply",
				dryRun: true,
				path,
				...(outputPath ? { output: outputPath } : {}),
				applied: previewApplied,
			});
			return EXIT.OK;
		}

		const applied = trackChanges.apply(accepts, rejects);

		await document.save(outputPath);

		await respondAck({
			ok: true,
			operation: "track-changes.apply",
			path: outputPath ?? path,
			applied,
		});

		// Re-list any changes the agent addressed in neither list, with their now-
		// renumbered handles (see remainingTrackedChangesBlock). Quiet path only —
		// the JSON ack is the machine contract. Only write when something remains;
		// `writeStdout("")` would DISCARD the ack just printed (Bun's stdout sink
		// drops a prior write on an empty one).
		if (!parsed.values.verbose) {
			const remaining = await remainingTrackedChangesBlock(
				outputPath ?? path,
				"apply",
			);
			if (remaining) await writeStdout(remaining);
		}
		return EXIT.OK;
	} catch (error) {
		if (error instanceof TrackedChangeNotFoundError) {
			return fail(
				"TRACKED_CHANGE_NOT_FOUND",
				error.message,
				"Run 'docx track-changes list FILE' to see available handles.",
			);
		}
		throw error;
	}
}

/** Numeric suffix of a `tcN` id — its document-order position (the reader
 *  assigns tcN in walk order). Used to sort the dry-run preview into the same
 *  order the real apply ack uses. */
function trackedChangeIndex(id: string): number {
	const match = id.match(/^tc(\d+)$/);
	return match?.[1] ? Number(match[1]) : 0;
}
