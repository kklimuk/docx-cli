import {
	insertNoteReferenceAtOffset,
	NoteBody,
	type NoteKind,
	NoteOffsetOutOfRangeError,
	NoteReferenceRun,
	noteConfig,
	TrackedNoteBody,
} from "@core/notes";
import { sumRunBearingTextLength, XmlNode } from "@core/parser";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "@core/track-changes";
import { Ins } from "@core/track-changes/emit";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	return `docx ${verb} add — author a new ${kind} anchored to a paragraph offset

Usage:
  docx ${verb} add FILE --at pN[:offset] --text TEXT [options]

Anchor:
  --at pN              Append the reference at the end of paragraph pN.
  --at pN:offset       Insert the reference at character offset within pN.

Required:
  --text TEXT          ${capitalize(kind)} body text (single paragraph).

Optional:
  -o, --output PATH    Write to PATH instead of overwriting FILE.
  --dry-run            Print what would be added; do not write the file.
  -v, --verbose        Print the success ack JSON (default: silent on success).
  -h, --help           Show this help.

Examples:
  docx ${verb} add doc.docx --at p3 --text "See p.42 for the long form."
  docx ${verb} add doc.docx --at p0:12 --text "Citation needed."

Notes:
  The ${kind} body's id appears in the AST as "${idPrefix}N" (used by
  \`docx ${verb} delete\` / \`edit\` / \`list\`). If the document has no
  ${kind}s yet, ${verb}.xml is provisioned with Word's reserved
  separator/continuationSeparator boilerplate.
`;
}

function capitalize(word: string): string {
	return word.charAt(0).toUpperCase() + word.slice(1);
}

export async function runAddNote(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			text: { type: "string" },
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

	const anchorInput = parsed.values.at as string | undefined;
	const text = parsed.values.text as string | undefined;
	if (!anchorInput) return fail("USAGE", "Missing --at pN[:offset]", help);
	if (!text) return fail("USAGE", "Missing --text TEXT", help);

	const anchor = parseAnchor(anchorInput);
	if (!anchor) {
		return fail(
			"INVALID_LOCATOR",
			`--at expects pN or pN:offset, got "${anchorInput}"`,
			help,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const paragraphRef = await resolveBlockOrFail(document, anchor.blockId);
	if (typeof paragraphRef === "number") return paragraphRef;

	const existingNotes =
		kind === "footnote" ? document.footnotes : document.endnotes;
	const numericId = existingNotes?.nextId() ?? "1";
	const config = noteConfig(kind);
	const idLabel = `${config.idPrefix}${numericId}`;
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: `${kind}s.add`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			id: idLabel,
			at: anchorInput,
		});
		return EXIT.OK;
	}

	// Tracking is doc-level: when `<w:trackChanges/>` is set in settings.xml,
	// Word wraps BOTH the body reference run AND the entire body content in
	// `<w:ins>` (different revision ids, shared author/date). Mirroring Word
	// exactly is what makes accept/reject in Word render this CLI's edits
	// correctly — empirically validated against `/tmp/fn-probe/add.docx`.
	const tracked = document.isTrackChangesEnabled();
	let refMeta: TrackedMeta | undefined;
	let bodyMeta: TrackedMeta | undefined;
	if (tracked) {
		const allocator = new TrackChanges(document).createAllocator();
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		refMeta = { author, date, revisionId: allocator.next() };
		bodyMeta = { author, date, revisionId: allocator.next() };
	}

	const baseRun = (
		<NoteReferenceRun config={config} id={numericId} />
	) as ReturnType<typeof NoteReferenceRun>;
	const referenceRun = refMeta
		? ((<Ins meta={refMeta}>{baseRun}</Ins>) as ReturnType<
				typeof NoteReferenceRun
			>)
		: baseRun;

	const targetOffset =
		anchor.offset ?? sumRunBearingTextLength(paragraphRef.node.children);

	try {
		insertNoteReferenceAtOffset(paragraphRef.node, targetOffset, referenceRun);
	} catch (error) {
		if (error instanceof NoteOffsetOutOfRangeError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	const notesView =
		kind === "footnote"
			? document.ensureFootnotes()
			: document.ensureEndnotes();
	const notesRoot = XmlNode.findRoot(notesView.tree, config.rootTag);
	if (!notesRoot) throw new Error(`expected <${config.rootTag}> root`);
	notesRoot.children.push(
		bodyMeta ? (
			<TrackedNoteBody
				config={config}
				id={numericId}
				text={text}
				meta={bodyMeta}
			/>
		) : (
			<NoteBody config={config} id={numericId} text={text} />
		),
	);
	notesView.ensureNoteStyles(document.ensureStyles());

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: `${kind}s.add`,
		path: outputPath ?? path,
		id: idLabel,
		at: anchorInput,
	});
	return EXIT.OK;
}

/** Parse `pN` (anchor at end of paragraph) or `pN:offset` (anchor at a
 *  character offset). Anything else returns null. Unlike `parseLocator`,
 *  we accept a bare single-offset suffix because there's no `pN:S-E` form
 *  here — we always insert at a single point. */
function parseAnchor(
	input: string,
): { blockId: string; offset?: number } | null {
	const trimmed = input.trim();
	const bareMatch = trimmed.match(/^p(\d+)$/);
	if (bareMatch) {
		return { blockId: `p${bareMatch[1]}` };
	}
	const offsetMatch = trimmed.match(/^p(\d+):(\d+)$/);
	if (offsetMatch) {
		return {
			blockId: `p${offsetMatch[1]}`,
			offset: Number(offsetMatch[2]),
		};
	}
	return null;
}

export async function run(args: string[]): Promise<number> {
	return runAddNote(args, "footnote");
}
