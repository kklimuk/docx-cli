import { saveDocView } from "@core";
import {
	findNoteByNumericId,
	type NoteKind,
	noteConfig,
	removeNoteReferences,
	wrapNoteBodyAsDeleted,
} from "@core/notes";
import { XmlNode } from "@core/parser";
import {
	createRevisionAllocator,
	isTrackChangesEnabled,
	resolveAuthor,
	resolveDate,
	type TrackedMeta,
} from "@core/track-changes";
import { Del } from "@core/track-changes/emit";
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
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				id: { type: "string" },
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
		return fail("USAGE", message, helpFor(kind));
	}

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

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = findNoteByNumericId(view, kind, numericId);
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

	const tracked = isTrackChangesEnabled(view);
	if (tracked) {
		const allocator = createRevisionAllocator(view);
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
		wrapReferencesAsDeleted(view.documentTree, kind, numericId, refMeta);
		wrapNoteBodyAsDeleted(reference.node, bodyMeta);
	} else {
		const index = reference.parent.indexOf(reference.node);
		if (index !== -1) reference.parent.splice(index, 1);
		removeNoteReferences(view.documentTree, kind, numericId);
	}

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: `${kind}s.delete`,
		path: outputPath ?? path,
		id: idLabel,
	});
	return EXIT.OK;
}

/** Find every `<w:r>` in `documentTree` that contains a reference to
 *  `(kind, numericId)` and wrap it in `<w:del meta>`. Mirrors the pattern in
 *  `removeNoteReferences` but wraps instead of removing — the run stays
 *  intact, the wrapper records the revision. The post-pass GC in apply.ts is
 *  what eventually removes both reference and body on accept.
 *
 *  Skips runs already inside a `<w:del>` (idempotent under repeated calls;
 *  also avoids double-wrap if the agent runs `footnotes delete fnN` twice). */
function wrapReferencesAsDeleted(
	documentTree: XmlNode[],
	kind: NoteKind,
	numericId: string,
	meta: TrackedMeta,
): void {
	const document = XmlNode.findRoot(documentTree, "w:document");
	if (!document) return;
	const config = noteConfig(kind);
	walkAndWrap(document, config.referenceTag, numericId, meta);
}

function walkAndWrap(
	node: XmlNode,
	referenceTag: string,
	numericId: string,
	meta: TrackedMeta,
): void {
	if (node.tag === "w:del") return; // already inside a tracked deletion
	const replaced: XmlNode[] = [];
	for (const child of node.children) {
		if (
			child.tag === "w:r" &&
			runContainsReference(child, referenceTag, numericId)
		) {
			replaced.push(<Del meta={meta}>{child}</Del>);
			continue;
		}
		walkAndWrap(child, referenceTag, numericId, meta);
		replaced.push(child);
	}
	node.children = replaced;
}

function runContainsReference(
	run: XmlNode,
	referenceTag: string,
	numericId: string,
): boolean {
	for (const child of run.children) {
		if (
			child.tag === referenceTag &&
			child.getAttribute("w:id") === numericId
		) {
			return true;
		}
	}
	return false;
}

export async function run(args: string[]): Promise<number> {
	return runDeleteNote(args, "footnote");
}
