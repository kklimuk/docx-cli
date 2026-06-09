import { type Document, describeForms, MarkdownImport, type Run } from "@core";
import type { NotesView } from "@core/ast/document/notes";
import { RunElement } from "@core/blocks";
import { findTextSpans } from "@core/find";
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
import { parseRunsArg } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveTracked,
	respond,
	respondMinted,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const ANCHOR_FORMS = describeForms(
	["paragraph", "cellParagraph"],
	"                       ",
);

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	const article = kind === "endnote" ? "an endnote" : "a footnote";
	return `docx ${verb} add — author ${article} anchored to a paragraph offset

Usage:
  docx ${verb} add FILE --at LOCATOR --text TEXT [options]
  docx ${verb} add FILE --at LOCATOR --runs JSON [options]
  docx ${verb} add FILE --at LOCATOR --markdown TEXT [options]

Anchor (one required):
  --at LOCATOR         Where to anchor the reference. One of:
${ANCHOR_FORMS}
                       Append the reference at the END of the paragraph, OR add
                       a single 0-based character offset to insert at a point:
                       pN:offset or tN:rRcC:pK:offset (note-specific point form;
                       there is no S-E span here). See \`docx info locators\`.
  --anchor PHRASE      Find PHRASE (same matcher as \`docx find\`) and drop the
                       reference right AFTER it — no character-offset math.
                       Errors if it matches more than once without --occurrence.
  --occurrence N       With --anchor: pick the Nth match (1-indexed, default 1).

Body (exactly one required):
  --text TEXT          ${capitalize(kind)} body text (single paragraph).
  --runs JSON          Custom runs as a Run[] JSON array (one paragraph).
  --markdown TEXT      GFM markdown body (may produce multiple paragraphs).
  --markdown-file PATH Same as --markdown, but read from PATH (- for stdin).

Optional:
  --author NAME        Author for tracked attribution (default: $DOCX_AUTHOR
                       env, fallback "Reviewer"). Ignored when tracking off.
  -o, --output PATH    Write to PATH instead of overwriting FILE.
  --dry-run            Print what would be added; do not write the file.
  -v, --verbose        Print the success ack JSON (default: prints the new id).
  -h, --help           Show this help.

Output:
  Prints the new id (${idPrefix}N), one per line. Address it later with
  \`--at ${idPrefix}N\` on \`${verb} edit\` / \`delete\`. --verbose prints the full ack
  {ok:true, operation, path, id, at}. Errors print {code, error, hint?} with a
  nonzero exit. Discover ids with \`docx ${verb} list FILE\`.

Examples:
  docx ${verb} add doc.docx --at p3 --text "See p.42 for the long form."
  docx ${verb} add doc.docx --anchor "$4.2M in Q3 revenue" --text "Source: audited close."
  docx ${verb} add doc.docx --at p0:12 --text "Citation needed."
  docx ${verb} add doc.docx --at t0:r1c2:p0 --text "Cell-anchored note."
  docx ${verb} add doc.docx --at p3 --runs '[{"type":"text","text":"Bold","bold":true}]'
  docx ${verb} add doc.docx --at p3 --markdown $'First para.\\n\\nSecond para.'

Notes:
  The ${kind} body's id appears in the AST as "${idPrefix}N" (used by
  \`docx ${verb} delete\` / \`edit\` / \`list\`). If the document has no
  ${kind}s yet, ${verb}.xml is provisioned with Word's reserved
  separator/continuationSeparator boilerplate. In a --markdown body, links are
  preserved (their relationship is written into the note part's own rels);
  images are dropped. Rich bodies (--runs / --markdown) are unsupported under
  track-changes — use --text, or turn tracking off.
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
			anchor: { type: "string" },
			occurrence: { type: "string" },
			text: { type: "string" },
			runs: { type: "string" },
			markdown: { type: "string" },
			"markdown-file": { type: "string" },
			author: { type: "string" },
			track: { type: "boolean" },
			...SAVE_FLAGS,
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

	const atInput = parsed.values.at as string | undefined;
	const anchorPhrase = parsed.values.anchor as string | undefined;
	const anchorCount = (atInput ? 1 : 0) + (anchorPhrase ? 1 : 0);
	if (anchorCount === 0) {
		return fail("USAGE", "Specify --at LOCATOR or --anchor PHRASE", help);
	}
	if (anchorCount > 1) {
		return fail("USAGE", "--at and --anchor are mutually exclusive", help);
	}

	const text = parsed.values.text as string | undefined;
	const runsJson = parsed.values.runs as string | undefined;
	const markdown = parsed.values.markdown as string | undefined;
	const markdownFile = parsed.values["markdown-file"] as string | undefined;

	const bodyCount =
		(text !== undefined ? 1 : 0) +
		(runsJson !== undefined ? 1 : 0) +
		(markdown !== undefined ? 1 : 0) +
		(markdownFile !== undefined ? 1 : 0);
	if (bodyCount === 0) {
		return fail(
			"USAGE",
			"Specify exactly one of --text, --runs, --markdown, or --markdown-file",
			help,
		);
	}
	if (bodyCount > 1) {
		return fail(
			"USAGE",
			"--text, --runs, --markdown, and --markdown-file are mutually exclusive",
			help,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	// Resolve the anchor to a { blockId, offset? }. `--anchor PHRASE` finds the
	// phrase (same matcher as `docx find`) and drops the reference right AFTER
	// it — no character-offset arithmetic. `--at` keeps the explicit locator.
	let blockId: string;
	let offset: number | undefined;
	let anchorLabel: string;
	if (anchorPhrase !== undefined) {
		const occurrenceRaw = parsed.values.occurrence as string | undefined;
		const occurrence = occurrenceRaw === undefined ? 1 : Number(occurrenceRaw);
		if (!Number.isInteger(occurrence) || occurrence < 1) {
			return fail(
				"USAGE",
				`--occurrence must be a positive integer (1-indexed), got "${occurrenceRaw}"`,
			);
		}
		const matches = findTextSpans(document.body, anchorPhrase).matches;
		if (matches.length === 0) {
			return fail(
				"MATCH_NOT_FOUND",
				`anchor not found: ${JSON.stringify(anchorPhrase)}`,
			);
		}
		if (matches.length > 1 && occurrenceRaw === undefined) {
			return fail(
				"USAGE",
				`anchor matches ${matches.length} times; pass --occurrence N (1..${matches.length}) to disambiguate`,
			);
		}
		const match = matches[occurrence - 1];
		if (!match) {
			return fail(
				"MATCH_NOT_FOUND",
				`only ${matches.length} match(es) for anchor; --occurrence ${occurrence} is out of range`,
			);
		}
		blockId = match.blockId;
		offset = match.end; // reference mark goes after the matched phrase
		anchorLabel = `${match.blockId}:${match.end}`;
	} else {
		const anchor = parseAnchor(atInput as string);
		if (!anchor) {
			return fail(
				"INVALID_LOCATOR",
				`--at expects pN[:offset] or tN:rRcC:pK[:offset], got "${atInput}"`,
				help,
			);
		}
		blockId = anchor.blockId;
		offset = anchor.offset;
		anchorLabel = atInput as string;
	}

	const paragraphRef = await resolveBlockOrFail(document, blockId);
	if (typeof paragraphRef === "number") return paragraphRef;

	const existingNotes =
		kind === "footnote" ? document.footnotes : document.endnotes;
	const numericId = existingNotes?.nextId() ?? "1";
	const config = noteConfig(kind);
	const idLabel = `${config.idPrefix}${numericId}`;
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: `${kind}s.add`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			id: idLabel,
			at: anchorLabel,
		});
		return EXIT.OK;
	}

	const tracked = resolveTracked(document, parsed.values.track);
	const isRichBody = text === undefined;
	if (tracked && isRichBody) {
		return fail(
			"USAGE",
			"rich note bodies (--runs/--markdown) under track-changes are not supported yet — use --text, or turn tracking off",
			help,
		);
	}

	// Provision the notes part up front (idempotent) so a markdown body's
	// hyperlink rels mint into THIS part's own rels, not the document's — a
	// `<w:hyperlink r:id>` spliced into footnotes.xml resolves against
	// word/_rels/footnotes.xml.rels (a dangling rId otherwise corrupts the file).
	const notesView =
		kind === "footnote"
			? document.ensureFootnotes()
			: document.ensureEndnotes();

	// Build the body content (runs for the first paragraph + any extra sibling
	// paragraphs for a multi-paragraph markdown body). Resolved before any
	// reference-side mutation so a malformed --runs / bad markdown fails cleanly.
	const body = await buildNoteBody(document, notesView, {
		text,
		runsJson,
		markdown,
		markdownFile,
	});
	if (typeof body === "number") return body;

	// Tracking is doc-level: when `<w:trackChanges/>` is set in settings.xml,
	// Word wraps BOTH the body reference run AND the entire body content in
	// `<w:ins>` (different revision ids, shared author/date). Mirroring Word
	// exactly is what makes accept/reject in Word render this CLI's edits
	// correctly — empirically validated against `/tmp/fn-probe/add.docx`.
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
		offset ?? sumRunBearingTextLength(paragraphRef.node.children);

	try {
		insertNoteReferenceAtOffset(paragraphRef.node, targetOffset, referenceRun);
	} catch (error) {
		if (error instanceof NoteOffsetOutOfRangeError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	const notesRoot = XmlNode.findRoot(notesView.tree, config.rootTag);
	if (!notesRoot) throw new Error(`expected <${config.rootTag}> root`);
	notesRoot.children.push(
		bodyMeta ? (
			// Tracked path is `--text`-only (rich bodies were rejected above), so
			// the verified single-paragraph Word shape is preserved exactly.
			<TrackedNoteBody
				config={config}
				id={numericId}
				text={text ?? ""}
				meta={bodyMeta}
			/>
		) : (
			<NoteBody
				config={config}
				id={numericId}
				text={text}
				runs={body.runs}
				paragraphs={body.paragraphs}
			/>
		),
	);
	notesView.ensureNoteStyles(document.ensureStyles());

	await document.save(outputPath);

	// The minted id is the addressable handle (`--at ${idLabel}`) the agent
	// can't reconstruct, so it prints by default — one per line — upgrading to
	// the full ack under --verbose.
	await respondMinted([idLabel], {
		ok: true,
		operation: `${kind}s.add`,
		path: outputPath ?? path,
		id: idLabel,
		at: anchorLabel,
	});
	return EXIT.OK;
}

type ResolvedBody = { runs?: XmlNode[]; paragraphs?: XmlNode[] };

/** Resolve the note body content from whichever flag was passed. For `--text`
 *  the convenience path is left to `NoteBody` (returns no explicit runs). For
 *  `--runs` we build `<w:r>` siblings via the shared run emitter. For markdown
 *  we run `MarkdownImport.blocks` and split the result: the first block's runs
 *  lead the note's first paragraph (after the back-ref) and any further blocks
 *  become sibling `<w:p>`. Returns a fail() exit code on malformed input. */
async function buildNoteBody(
	document: Document,
	notesView: NotesView,
	flags: {
		text?: string;
		runsJson?: string;
		markdown?: string;
		markdownFile?: string;
	},
): Promise<ResolvedBody | number> {
	if (flags.text !== undefined) return {};

	if (flags.runsJson !== undefined) {
		const runs = await parseRunsArg(flags.runsJson);
		if (typeof runs === "number") return runs;
		return { runs: runsToXml(runs) };
	}

	const source =
		flags.markdown !== undefined
			? flags.markdown
			: await readMarkdownSource(flags.markdownFile as string);
	if (typeof source === "number") return source;

	let blocks: XmlNode[];
	try {
		// Note bodies route hyperlink rels into the notes part's OWN rels and
		// drop images — a `<w:hyperlink r:id>` / `<w:drawing r:embed>` lands in
		// footnotes.xml, which resolves rIds against word/_rels/footnotes.xml.rels,
		// not the document's (a dangling rId there is the "unreadable content" bug).
		blocks = await new MarkdownImport(document).blocks(source, {
			relationships: notesView.ensureRelationships(),
			stripImages: true,
		});
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to parse --markdown body: ${message}`);
	}

	return splitMarkdownBlocks(blocks);
}

/** Split markdown-produced blocks into the note body shape: the first
 *  paragraph's runs (everything inside the first `<w:p>` except its `<w:pPr>`)
 *  lead the note's first paragraph; subsequent blocks become sibling
 *  paragraphs. Non-paragraph leading blocks (a table, code, …) keep their own
 *  `<w:p>`/block as a sibling and the first note paragraph carries only the
 *  back-ref. */
function splitMarkdownBlocks(blocks: XmlNode[]): ResolvedBody {
	const [first, ...rest] = blocks;
	if (!first) return {};
	if (first.tag !== "w:p") {
		return { paragraphs: blocks };
	}
	const runs = first.children.filter((child) => child.tag !== "w:pPr");
	return { runs, paragraphs: rest.length > 0 ? rest : undefined };
}

/** Convert a `Run[]` to fresh `<w:r>` `XmlNode`s via the shared run emitter,
 *  dropping any run kind the emitter can't materialize (image/equation/…). */
function runsToXml(runs: Run[]): XmlNode[] {
	const out: XmlNode[] = [];
	for (const run of runs) {
		const element = RunElement({ run });
		if (element) out.push(element);
	}
	return out;
}

async function readMarkdownSource(path: string): Promise<string | number> {
	try {
		return path === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(path).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --markdown-file ${path}: ${message}`,
		);
	}
}

/** Parse `pN` / `pN:offset` or a cell-nested `tN:rRcC:pK` / `tN:rRcC:pK:offset`
 *  anchor. Routes through the standard locator so the cell-paragraph forms
 *  resolve the same way the rest of the CLI does; the optional trailing
 *  `:offset` is a single 0-based character point (note-specific — there is no
 *  `S-E` span form here). Anything else returns null. */
function parseAnchor(
	input: string,
): { blockId: string; offset?: number } | null {
	const trimmed = input.trim();

	// Plain paragraph anchors: pN or pN:offset.
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

	// Cell-nested paragraph anchors: tN:rRcC:pK or tN:rRcC:pK:offset. We pull a
	// trailing `:offset` (a single integer point) off the end if present, then
	// hand the remainder to `resolveBlockOrFail` as a cell-paragraph locator.
	if (!trimmed.startsWith("t")) return null;
	const cellOffsetMatch = trimmed.match(/^(.*:p\d+):(\d+)$/);
	if (cellOffsetMatch?.[1]) {
		return { blockId: cellOffsetMatch[1], offset: Number(cellOffsetMatch[2]) };
	}
	if (/:p\d+$/.test(trimmed)) {
		return { blockId: trimmed };
	}
	return null;
}

export async function run(args: string[]): Promise<number> {
	return runAddNote(args, "footnote");
}
