import { type Document, MarkdownImport, type Run } from "@core";
import type { NotesView } from "@core/ast/document/notes";
import { RunElement } from "@core/blocks";
import { w } from "@core/jsx";
import { type NoteKind, noteConfig, wrapNoteBodyAsEdited } from "@core/notes";
import type { XmlNode } from "@core/parser";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "@core/track-changes";
import { parseRunsArg } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	const article = kind === "endnote" ? "an endnote" : "a footnote";
	return `docx ${verb} edit — replace ${article}'s body

Usage:
  docx ${verb} edit FILE --at ${idPrefix}N --text TEXT [options]
  docx ${verb} edit FILE --at ${idPrefix}N --runs JSON [options]
  docx ${verb} edit FILE --at ${idPrefix}N --markdown TEXT [options]

Target:
  --at ${idPrefix}N              ${capitalize(kind)} id (e.g. ${idPrefix}0); the ${idPrefix} prefix is optional.
                       See \`docx info locators\`.

Body (exactly one required):
  --text TEXT          New body text. Replaces the current paragraph(s).
  --runs JSON          Custom runs as a Run[] JSON array (one paragraph).
  --markdown TEXT      GFM markdown body (may produce multiple paragraphs).
  --markdown-file PATH Same as --markdown, but read from PATH (- for stdin).

Optional:
  --author NAME        Author for tracked attribution (default: $DOCX_AUTHOR
                       env, fallback "Reviewer"). Ignored when tracking off.
  -o, --output PATH    Write to PATH instead of overwriting FILE.
  --dry-run            Print what would change; do not write the file.
  -v, --verbose        Print the success ack JSON (default: a one-line confirmation).
  -h, --help           Show this help.

Output:
  Prints a one-line confirmation on success; --verbose prints {ok:true, operation, path, id}. Errors
  print {code, error, hint?} with a nonzero exit. Discover ids with
  \`docx ${verb} list FILE\`.

Examples:
  docx ${verb} edit doc.docx --at ${idPrefix}2 --text "Updated citation."
  docx ${verb} edit doc.docx --at ${idPrefix}2 --runs '[{"type":"text","text":"X","italic":true}]'
  docx ${verb} edit doc.docx --at ${idPrefix}2 --markdown $'First.\\n\\nSecond.'

Notes:
  In a --markdown body, links are preserved (their relationship is written into
  the note part's own rels); images are dropped. Rich bodies (--runs /
  --markdown) are unsupported under track-changes — use --text, or turn it off.
`;
}

function capitalize(word: string): string {
	return word.charAt(0).toUpperCase() + word.slice(1);
}

export async function runEditNote(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
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

	const idInput = parsed.values.at as string | undefined;
	const config = noteConfig(kind);
	if (!idInput) return fail("USAGE", `Missing --at ${config.idPrefix}N`, help);

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

	const numericId = idInput.startsWith(config.idPrefix)
		? idInput.slice(config.idPrefix.length)
		: idInput;
	const idLabel = `${config.idPrefix}${numericId}`;

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const notesView =
		kind === "footnote" ? document.footnotes : document.endnotes;
	const reference = notesView?.findByNumericId(numericId);
	if (!reference || !notesView) {
		return fail("BLOCK_NOT_FOUND", `${capitalize(kind)} not found: ${idLabel}`);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: `${kind}s.edit`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			id: idLabel,
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

	if (tracked) {
		// Tracked replace: `<w:ins>NEW</w:ins><w:del>OLD</w:del>` inside the
		// existing body paragraph, with `<w:footnoteRef/>` and leading
		// whitespace bare. Matches Word's empirical shape from
		// `/tmp/fn-probe/edit.docx` (ins precedes del in document order).
		// Word allocates distinct revision ids for the two sides — we match
		// that for empirical parity even though the OOXML spec allows shared
		// ids. `--text`-only: rich bodies were rejected above so the verified
		// single-run shape is preserved exactly.
		const allocator = new TrackChanges(document).createAllocator();
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		const insMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		const delMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		wrapNoteBodyAsEdited(reference.node, ` ${text ?? ""}`, insMeta, delMeta);
	} else {
		// Untracked: replace the note's content with a fresh body. The
		// `<w:footnote>` / `<w:endnote>` wrapper itself stays — anything that
		// referenced the original (back-references, custom marker attributes)
		// keeps pointing here. Only the inner paragraph content is rewritten.
		const body = await buildNoteBody(document, notesView, {
			text,
			runsJson,
			markdown,
			markdownFile,
		});
		if (typeof body === "number") return body;
		reference.node.children = [
			<NoteParagraph config={config} text={text} runs={body.runs} />,
			...(body.paragraphs ?? []),
		];
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: `${kind}s.edit`,
		path: outputPath ?? path,
		id: idLabel,
	});
	return EXIT.OK;
}

type ResolvedBody = { runs?: XmlNode[]; paragraphs?: XmlNode[] };

/** Resolve the replacement body content from whichever flag was passed. For
 *  `--text` the convenience path is left to `NoteParagraph`. For `--runs` we
 *  build `<w:r>` siblings via the shared run emitter. For markdown we run
 *  `MarkdownImport.blocks` and split: the first block's runs lead the note's
 *  first (back-ref-bearing) paragraph, further blocks become sibling `<w:p>`. */
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
		// Route hyperlink rels into the note part's OWN rels and drop images —
		// the blocks are spliced into footnotes.xml/endnotes.xml, whose rIds
		// resolve against word/_rels/<part>.xml.rels, not the document's. Without
		// this a note-body link is a dangling rId (Word's "unreadable content").
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

/** Split markdown-produced blocks into the note body shape — see the twin in
 *  `add.tsx`. The first paragraph's runs lead the back-ref paragraph; the rest
 *  become sibling paragraphs. A non-paragraph leading block keeps its own
 *  block as a sibling and the back-ref paragraph carries only the numeral. */
function splitMarkdownBlocks(blocks: XmlNode[]): ResolvedBody {
	const [first, ...rest] = blocks;
	if (!first) return {};
	if (first.tag !== "w:p") {
		return { paragraphs: blocks };
	}
	const runs = first.children.filter((child) => child.tag !== "w:pPr");
	return { runs, paragraphs: rest.length > 0 ? rest : undefined };
}

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

/** The note's first body paragraph — carries the back-reference marker
 *  (`<w:footnoteRef/>` / `<w:endnoteRef/>`) Word renders as the body numeral.
 *  `text` is the single-run convenience (a leading-space `<w:t>`); `runs` are
 *  explicit pre-built `<w:r>` siblings (runs wins if both are given). */
function NoteParagraph({
	config,
	text,
	runs,
}: {
	config: ReturnType<typeof noteConfig>;
	text?: string;
	runs?: XmlNode[];
}) {
	const BodyRef = config.kind === "footnote" ? w.footnoteRef : w.endnoteRef;
	const bodyRuns: XmlNode[] = runs ?? [
		<w.r>
			<w.t {...{ "xml:space": "preserve" }}>{` ${text ?? ""}`}</w.t>
		</w.r>,
	];
	return (
		<w.p>
			<w.pPr>
				<w.pStyle w-val={config.textStyle} />
			</w.pPr>
			<w.r>
				<w.rPr>
					<w.rStyle w-val={config.referenceStyle} />
				</w.rPr>
				<BodyRef />
			</w.r>
			{bodyRuns}
		</w.p>
	);
}

export async function run(args: string[]): Promise<number> {
	return runEditNote(args, "footnote");
}
