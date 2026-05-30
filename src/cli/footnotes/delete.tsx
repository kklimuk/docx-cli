import {
	type NoteKind,
	noteConfig,
	removeNoteReferences,
	wrapNoteBodyAsDeleted,
	wrapNoteReferencesAsDeleted,
} from "@core/notes";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "@core/track-changes";
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

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	return `docx ${verb} delete — remove a ${kind} body and every reference to it

Usage:
  docx ${verb} delete FILE --id ${idPrefix}N [options]

Required:
  --id ID              ${capitalize(kind)} id (e.g., ${idPrefix}1). The
                       \`${idPrefix}\` prefix is optional.

Optional:
  --author NAME        Author for tracked deletions (default: $DOCX_AUTHOR
                       env, fallback "docx-cli"). Ignored when tracking off.
  -o, --output PATH    Write to PATH instead of overwriting FILE.
  --dry-run            Print what would be removed; do not write the file.
  -v, --verbose        Print the success ack JSON (default: silent on success).
  -h, --help           Show this help.

Examples:
  docx ${verb} delete doc.docx --id ${idPrefix}3
`;
}

function capitalize(word: string): string {
	return word.charAt(0).toUpperCase() + word.slice(1);
}

export async function runDeleteNote(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			id: { type: "string" },
			author: { type: "string" },
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
		},
		helpFor(kind),
	);
	if (typeof parsed === "number") return parsed;

	const help = helpFor(kind);
	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const idInput = parsed.values.id as string | undefined;
	if (!idInput)
		return fail("USAGE", `Missing --id ${noteConfig(kind).idPrefix}N`, help);

	const config = noteConfig(kind);
	const numericId = idInput.startsWith(config.idPrefix)
		? idInput.slice(config.idPrefix.length)
		: idInput;
	const idLabel = `${config.idPrefix}${numericId}`;

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const notesView =
		kind === "footnote" ? document.footnotes : document.endnotes;
	const reference = notesView?.findByNumericId(numericId);
	if (!reference) {
		return fail("BLOCK_NOT_FOUND", `${capitalize(kind)} not found: ${idLabel}`);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: `${kind}s.delete`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			id: idLabel,
		});
		return EXIT.OK;
	}

	const tracked = document.isTrackChangesEnabled();
	if (tracked) {
		const allocator = new TrackChanges(document).createAllocator();
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		// Three coupled <w:del> revisions per Word's empirical shape — body
		// reference run, body content, paragraph-mark. Different revision ids,
		// shared author/date. Accept/reject pairs them via the footnote id
		// (the `w:id` on <w:footnoteReference>, not the revision id).
		const refMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		const bodyMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		wrapNoteReferencesAsDeleted(
			document.documentTree,
			kind,
			numericId,
			refMeta,
		);
		wrapNoteBodyAsDeleted(reference.node, bodyMeta);
	} else {
		const index = reference.parent.indexOf(reference.node);
		if (index !== -1) reference.parent.splice(index, 1);
		removeNoteReferences(document.documentTree, kind, numericId);
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: `${kind}s.delete`,
		path: outputPath ?? path,
		id: idLabel,
	});
	return EXIT.OK;
}

export async function run(args: string[]): Promise<number> {
	return runDeleteNote(args, "footnote");
}
