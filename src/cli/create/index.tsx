import {
	Document,
	MarkdownImport,
	MarkdownImportError,
	type XmlNode,
} from "@core";
import { buildBlankPackage } from "@core/create";
import {
	EXIT,
	fail,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx create — create a new minimal .docx

Usage:
  docx create FILE [options]

Options:
  --title TEXT     Document title
  --author TEXT    Document author (default: $DOCX_AUTHOR)
  --text TEXT      Seed first paragraph with this text
  --from PATH      Seed the body with parsed markdown from PATH (use "-" for
                   stdin). Mutex with --text. Uses the same markdown dialect
                   as 'docx insert --markdown'.
  --force          Overwrite if file exists
  -v, --verbose    Print the success ack JSON (default: silent on success)
  -h, --help       Show this help

Examples:
  docx create out.docx
  docx create out.docx --title "Spec" --author "Claude" --text "First paragraph."
  docx create out.docx --from draft.md
  cat draft.md | docx create out.docx --from -

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
			from: { type: "string" },
			force: { type: "boolean" },
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
	if (text !== undefined && fromPath !== undefined) {
		return fail(
			"USAGE",
			"Pass either --text or --from, not both",
			"--text seeds a single paragraph; --from parses a markdown file into the body.",
		);
	}

	if ((await Bun.file(path).exists()) && !parsed.values.force) {
		return fail(
			"USAGE",
			`File already exists: ${path}`,
			"Pass --force to overwrite.",
		);
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
	// alive once the package is on disk and re-opened.
	const pkg = buildBlankPackage({ path, title, author, text });
	await pkg.save();

	let blockCount = 1; // the seed paragraph buildBlankPackage emitted
	if (fromPath !== undefined) {
		const applied = await applyMarkdownToBody(path, fromPath);
		if (typeof applied === "number") return applied;
		blockCount = applied.blockCount;
	}

	await respondAck({
		ok: true,
		operation: "create",
		path,
		bytes: Bun.file(path).size,
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

	// Empty markdown source (or one containing only constructs the walker
	// drops, like frontmatter / link-definition orphans) → keep the seed
	// paragraph in place. Splicing zero blocks over it would leave a bare
	// `<w:sectPr>` as a direct `<w:body>` child, which ECMA-376 §17.2.2
	// frowns on and Word for Mac flags on save.
	if (blocks.length === 0) {
		// Nothing to write — the blank package on disk is already valid.
		return { blockCount: 1 };
	}

	// `Document.open()` populated the embedded views; `document.body.body`
	// is the parsed `<w:body>` XmlNode. No need to re-walk the tree.
	const body = document.body.body;
	// `buildBlankPackage` always emits exactly one placeholder `<w:p>` plus
	// the trailing `<w:sectPr>`. Replace the placeholder with the parsed
	// blocks; sectPr stays put at the tail.
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
