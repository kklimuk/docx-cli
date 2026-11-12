import {
	applyPageGeometry,
	Document,
	literalParagraphs,
	Marginals,
	MarkdownImport,
	MarkdownImportError,
	type PageGeometry,
	type XmlNode,
} from "@core";
import { buildBlankPackage } from "@core/create";
import { decodeInlineEscapes, parseSectionFlags } from "../parse-helpers";
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

Examples:
  docx create out.docx --from draft.md      # build a whole doc from markdown
  cat draft.md | docx create out.docx --from -
  docx create out.docx --text-file notes.txt   # literal text, no parsing
  docx create out.docx --title "Spec" --author "Claude" --text "First paragraph."
  docx create out.docx --from draft.md --orientation landscape --size a4
  docx create out.docx --text "Memo" --margins 1.25in --page-numbers

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
                     HTML is not converted: block HTML that carries text
                     (an .html page, <div>/<table>/<h1> blocks) is refused
                     and nothing is written — rewrite it as markdown.
                     Footnote/endnote bodies keep bold/italic + hyperlinks;
                     footnote labels renumber to [^fnN] on import. (Under
                     track-changes, note bodies flatten to plain text.)
  --orientation O    Page orientation: portrait (default) | landscape
  --size SIZE        Page size: letter (default) | legal | tabloid | a4 | a3 | a5,
                     or WxH — bare inches or a unit suffix: in, cm, mm, pt, tw
                     (8.5x14in, 210x297mm). A W>H size implies landscape.
  --margins M        Page margins: one value (uniform: bare inches like 1, or
                     1in, 2.5cm) OR four comma-separated top,right,bottom,left
                     (e.g. 1,1,1,1.5). A unit on only the last value applies to
                     all four.
  --header TEXT      Add a page header with this text (all pages)
  --footer TEXT      Add a page footer with this text (all pages)
  --page-numbers     Add centered page numbers in the footer (mutex with --footer).
                     For richer footers (date, "Page X of Y", alignment, fields)
                     use 'docx footers set' after create.
  --force            Overwrite if FILE already exists
  --dry-run          Print what would be created; do not write the file
  -v, --verbose      Print the success ack JSON (default: a one-line confirmation)
  -h, --help         Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  bytes, blocks}. --dry-run prints {operation, dryRun:true, path, ...} and
  writes nothing. Errors print {code, error, hint?} with a nonzero exit.

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
			orientation: { type: "string" },
			size: { type: "string" },
			margins: { type: "string" },
			header: { type: "string" },
			footer: { type: "string" },
			"page-numbers": { type: "boolean" },
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

	const text = decodeInlineEscapes(parsed.values.text as string | undefined);
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

	// Page geometry (--orientation/--size/--margins). parseSectionFlags also reads
	// columns/type, which `create` doesn't expose (unknown flags are rejected by
	// parseArgs), so only the geometry slice is ever populated here.
	const sectionFlags = await parseSectionFlags(parsed.values);
	if (typeof sectionFlags === "number") return sectionFlags;
	const geometry: PageGeometry = {
		pageSize: sectionFlags.pageSize,
		orientation: sectionFlags.orientation,
		margins: sectionFlags.margins,
	};
	const hasGeometry =
		geometry.pageSize !== undefined ||
		geometry.orientation !== undefined ||
		geometry.margins !== undefined;

	// Header/footer (--header/--footer text + --page-numbers footer). Applied via
	// the Marginals lens after content + geometry land (so the two-zone tab, if it
	// were used here, would see the final page size). --footer and --page-numbers
	// both set the footer, so they're mutually exclusive.
	const header = parsed.values.header as string | undefined;
	const footer = parsed.values.footer as string | undefined;
	const pageNumbers = Boolean(parsed.values["page-numbers"]);
	if (footer !== undefined && pageNumbers) {
		return fail(
			"USAGE",
			"--footer and --page-numbers both set the footer — pass one (use 'docx footers set' for a footer with text AND page numbers).",
		);
	}
	const marginals = { header, footer, pageNumbers };
	const hasMarginals =
		header !== undefined || footer !== undefined || pageNumbers;

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
			...(hasGeometry ? { page: geometry } : {}),
			...(hasMarginals ? { marginals } : {}),
		});
		return EXIT.OK;
	}

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const title = (parsed.values.title as string | undefined) ?? "";

	// Build the blank package (seeded with --text, or empty) and run every
	// content/geometry/marginals step against ONE in-memory Document over it,
	// saving once at the end. `Document.fromPackage` gives MarkdownImport the
	// full view set it needs (relationships, content-types, lazy-provisioned
	// views) without a disk round-trip — and a refused source (bad image, block
	// HTML that would drop text) writes NOTHING, so an existing FILE under
	// --force survives a failed create.
	const pkg = buildBlankPackage({ path: destination, title, author, text });
	const document = await Document.fromPackage(pkg);

	let blockCount = 1; // the seed paragraph buildBlankPackage emitted
	if (fromPath !== undefined) {
		const applied = await applyMarkdownToBody(document, fromPath);
		if (typeof applied === "number") return applied;
		blockCount = applied.blockCount;
	} else if (textFilePath !== undefined) {
		const applied = await applyLiteralToBody(document, textFilePath);
		if (typeof applied === "number") return applied;
		blockCount = applied.blockCount;
	}

	// No tracking for geometry: a brand-new document has no prior state to
	// snapshot. Headers/footers follow geometry so the part lands on the final
	// trailing sectPr.
	if (hasGeometry || hasMarginals) {
		const sectPr = document.body.body.children
			.filter((child) => child.tag === "w:sectPr")
			.pop();
		if (!sectPr) {
			return fail(
				"USAGE",
				"Internal: trailing <w:sectPr> not found in blank-package body",
			);
		}
		if (hasGeometry) applyPageGeometry(sectPr, geometry);
		if (hasMarginals) applyMarginals(document, sectPr, marginals);
	}

	await document.save();

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
	document: Document,
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

	let blocks: XmlNode[];
	try {
		blocks = await new MarkdownImport(document).blocks(source);
	} catch (error) {
		if (error instanceof MarkdownImportError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	return replacePlaceholder(document, blocks);
}

/** Read literal text from PATH (or stdin), split it into one paragraph per line
 *  (NO markdown parsing — every character lands verbatim), and splice the result
 *  over the seed paragraph. The `create` counterpart of `insert --text-file`,
 *  for building a doc from prose that GFM would corrupt. */
async function applyLiteralToBody(
	document: Document,
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

	return replacePlaceholder(document, literalParagraphs(source));
}

/** Apply create-time headers/footers to the trailing (document-wide)
 *  `<w:sectPr>` via the `Marginals` lens. */
function applyMarginals(
	document: Document,
	sectPr: XmlNode,
	marginals: { header?: string; footer?: string; pageNumbers: boolean },
): void {
	const lens = new Marginals(document);
	if (marginals.header !== undefined) {
		lens.set([sectPr], "header", "default", { text: marginals.header });
	}
	if (marginals.footer !== undefined) {
		lens.set([sectPr], "footer", "default", { text: marginals.footer });
	}
	if (marginals.pageNumbers) {
		lens.set([sectPr], "footer", "default", {
			field: { type: "page" },
			align: "center",
		});
	}
}

/** Replace the blank package's single seed `<w:p>` with `blocks` (the caller
 *  saves). `buildBlankPackage` always emits exactly one placeholder `<w:p>` plus the
 *  trailing `<w:sectPr>`; we swap the placeholder for the real content and leave
 *  the sectPr at the tail. An empty `blocks` (markdown source that walked to
 *  nothing) keeps the seed paragraph — splicing zero blocks would leave a bare
 *  `<w:sectPr>` as a direct `<w:body>` child, which ECMA-376 §17.2.2 frowns on
 *  and Word for Mac flags on save. */
async function replacePlaceholder(
	document: Document,
	blocks: XmlNode[],
): Promise<number | { blockCount: number }> {
	if (blocks.length === 0) return { blockCount: 1 };

	// `Document.fromPackage()` populated the embedded views; `document.body.body` is
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
	return { blockCount: blocks.length };
}
