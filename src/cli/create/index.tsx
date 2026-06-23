import {
	Document,
	literalParagraphs,
	MarkdownImport,
	MarkdownImportError,
	type XmlNode,
} from "@core";
import { buildBlankPackage } from "@core/create";
import {
	EXIT,
	fail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx create — create a new minimal .docx

Usage:
  docx create FILE [options]

Options:
  --title TEXT       Document title
  --author TEXT      Document author (default: $DOCX_AUTHOR)
  --text TEXT        Seed first paragraph with this text. One content source
                     only (mutex with --text-file / --from).
  --text-file PATH   Seed the body with LITERAL multi-paragraph text from PATH
                     (use "-" for stdin), NOT parsed as markdown — every
                     character lands verbatim, each newline starts a new
                     paragraph. Use for prose GFM would corrupt ("3. note"
                     stays "3.", *x* / [t](u) / bare URLs / {++x++} untouched).
                     One content source only.
  --from PATH        Seed the body with parsed markdown from PATH (use "-" for
                     stdin). One content source only. Uses the same markdown
                     dialect as 'docx insert --markdown'. This is the canonical
                     way to build a whole .docx from a markdown file.
                     Footnote/endnote bodies keep bold/italic + hyperlinks;
                     footnote labels renumber to [^fnN] on import. (Under
                     track-changes, note bodies flatten to plain text.)
  --force            Overwrite if FILE already exists
  --dry-run          Print what would be created; do not write the file
  -v, --verbose      Print the success ack JSON (default: a one-line confirmation)
  -h, --help         Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  bytes, blocks}. --dry-run prints {operation, dryRun:true, path, ...} and
  writes nothing. Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx create out.docx
  docx create out.docx --title "Spec" --author "Claude" --text "First paragraph."
  docx create out.docx --from draft.md
  cat draft.md | docx create out.docx --from -
  docx create out.docx --text-file reviewer-notes.txt
  cat notes.txt | docx create out.docx --text-file -

For a doc that opens with a code block, chain create with insert:
  docx create out.docx
  docx insert out.docx --after p0 --code-file snippet.py --language python
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			title: { type: "string" },
			author: { type: "string" },
			text: { type: "string" },
			"text-file": { type: "string" },
			from: { type: "string" },
			force: { type: "boolean" },
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
	if (!path) {
		return fail("USAGE", "Missing FILE argument", HELP);
	}

	const text = parsed.values.text as string | undefined;
	const fromPath = parsed.values.from as string | undefined;
	const textFilePath = parsed.values["text-file"] as string | undefined;
	const contentSources = [text, fromPath, textFilePath].filter(
		(value) => value !== undefined,
	);
	if (contentSources.length > 1) {
		return fail(
			"USAGE",
			"Pass at most one of --text, --text-file, --from",
			"--text seeds one paragraph; --text-file seeds literal multi-paragraph text; --from parses a markdown file into the body.",
		);
	}

	const dryRun = Boolean(parsed.values["dry-run"]);
	// `create`'s positional FILE *is* the destination (unlike the mutators,
	// which edit an existing FILE and use -o to write elsewhere) — so there's no
	// -o here; the positional is the output.
	const destination = path;

	if ((await Bun.file(destination).exists()) && !parsed.values.force) {
		return fail(
			"USAGE",
			`File already exists: ${destination}`,
			"Pass --force to overwrite.",
		);
	}

	if (dryRun) {
		await respond({
			operation: "create",
			dryRun: true,
			path: destination,
			...(text !== undefined ? { text } : {}),
			...(textFilePath !== undefined ? { textFile: textFilePath } : {}),
			...(fromPath !== undefined ? { from: fromPath } : {}),
		});
		return EXIT.OK;
	}

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const title = (parsed.values.title as string | undefined) ?? "";

	// Seed with --text (or empty) first. If --from was passed, we re-open the
	// just-written package, replace the placeholder paragraph with parsed
	// markdown blocks, and save again. We don't try to do the markdown parse
	// inline against the raw template tree because MarkdownImport needs a
	// full Document (relationships, content-types, lazy-provisioned views) to
	// register footnote bodies / hyperlinks / images — all of which only come
	// alive once the package is on disk and re-opened. We write the blank
	// package straight to `destination` (the --output target if given, else
	// FILE), then re-open and re-save that same file for the markdown pass.
	const pkg = buildBlankPackage({ path: destination, title, author, text });
	await pkg.save();

	let blockCount = 1; // the seed paragraph buildBlankPackage emitted
	if (fromPath !== undefined) {
		const applied = await applyMarkdownToBody(destination, fromPath);
		if (typeof applied === "number") return applied;
		blockCount = applied.blockCount;
	} else if (textFilePath !== undefined) {
		const applied = await applyLiteralToBody(destination, textFilePath);
		if (typeof applied === "number") return applied;
		blockCount = applied.blockCount;
	}

	await respondAck({
		ok: true,
		operation: "create",
		path: destination,
		bytes: Bun.file(destination).size,
		blocks: blockCount,
	});
	return EXIT.OK;
}

/** Read the markdown source, parse + walk through MarkdownImport, and splice
 * the resulting blocks into the freshly-created doc's body — replacing the
 * empty placeholder paragraph. Returns an exit code on failure, or the
 * applied block count on success so the caller can report it in the ack. */
async function applyMarkdownToBody(
	docxPath: string,
	markdownPath: string,
): Promise<number | { blockCount: number }> {
	let source: string;
	try {
		source =
			markdownPath === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(markdownPath).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --from ${markdownPath}: ${message}`,
		);
	}

	const document = await Document.open(docxPath);
	let blocks: XmlNode[];
	try {
		blocks = await new MarkdownImport(document).blocks(source);
	} catch (error) {
		if (error instanceof MarkdownImportError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	return replacePlaceholderAndSave(document, blocks);
}

/** Read literal text from PATH (or stdin), split it into one paragraph per line
 *  (NO markdown parsing — every character lands verbatim), and splice the result
 *  over the seed paragraph. The `create` counterpart of `insert --text-file`,
 *  for building a doc from prose that GFM would corrupt. */
async function applyLiteralToBody(
	docxPath: string,
	textFilePath: string,
): Promise<number | { blockCount: number }> {
	let source: string;
	try {
		source =
			textFilePath === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(textFilePath).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --text-file ${textFilePath}: ${message}`,
		);
	}

	const document = await Document.open(docxPath);
	return replacePlaceholderAndSave(document, literalParagraphs(source));
}

/** Replace the blank package's single seed `<w:p>` with `blocks` and save.
 *  `buildBlankPackage` always emits exactly one placeholder `<w:p>` plus the
 *  trailing `<w:sectPr>`; we swap the placeholder for the real content and leave
 *  the sectPr at the tail. An empty `blocks` (markdown source that walked to
 *  nothing) keeps the seed paragraph — splicing zero blocks would leave a bare
 *  `<w:sectPr>` as a direct `<w:body>` child, which ECMA-376 §17.2.2 frowns on
 *  and Word for Mac flags on save. */
async function replacePlaceholderAndSave(
	document: Document,
	blocks: XmlNode[],
): Promise<number | { blockCount: number }> {
	if (blocks.length === 0) return { blockCount: 1 };

	// `Document.open()` populated the embedded views; `document.body.body` is
	// the parsed `<w:body>` XmlNode. No need to re-walk the tree.
	const body = document.body.body;
	const placeholderIndex = body.children.findIndex(
		(child) => child.tag === "w:p",
	);
	if (placeholderIndex === -1) {
		return fail(
			"USAGE",
			"Internal: placeholder <w:p> not found in blank-package body",
		);
	}
	body.children.splice(placeholderIndex, 1, ...blocks);
	await document.save();
	return { blockCount: blocks.length };
}
