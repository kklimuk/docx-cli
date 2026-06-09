import { Images } from "@core/image";
import { resolveView } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { MarkdownLocatorError, renderMarkdown } from "./markdown";

const HELP = `docx read — render document body as Markdown, or print AST as JSON

Usage:
  docx read FILE [options]

Options:
  --ast             Print the typed AST as JSON instead of rendering Markdown.
                    Disables all the Markdown-only flags below.
  --from LOC        Start rendering at top-level block LOC (inclusive)
  --to LOC          End rendering at top-level block LOC (inclusive)
                    A two-ended TOP-LEVEL BLOCK slice — the ends may be
                    different block types. Accepts pN, tN, sN, pN-pM, pN:S-E,
                    pN:S-pM:E, tN:rRcC[:pK[:S-E]]; cell/span/range ends collapse
                    to their enclosing top-level block (the table or paragraph).
                    (This is a block slice, distinct from a character span like
                    pN:S-E. See \`docx info locators\`.)
  --accepted        Default view: render the post-accept document — drop
                    subtractive wrappers (<w:del>, <w:moveFrom>), inline
                    additive wrappers (<w:ins>, <w:moveTo>) as plain text,
                    no markers/refs. Kept as an explicit alias for clarity.
  --baseline        Render the pre-change view: drop additive wrappers
                    (<w:ins>, <w:moveTo>), inline subtractive wrappers
                    (<w:del>, <w:moveFrom>) as plain text, no markers/refs.
  --current         Render the raw concatenation: additive wrappers as
                    {++text++}[^tcN] and subtractive as {--text--}[^tcN]
                    (CriticMarkup); the [^tcN] footnote spells out the kind
                    (insertion / deletion / moveTo / moveFrom). Mutually
                    exclusive with --accepted/--baseline.
  --comments        Append [^cN] after each commented span and emit a
                    footnote definition for each comment at the end of the
                    output (author, date, body).
  -h, --help        Show this help

Output:
  Default: GitHub-flavored Markdown. Each paragraph is trailed by an HTML
  comment with its locator (<!-- p3 -->) so an agent can recover ids from the
  rendered text. --ast: the bare JSON AST (the body object: blocks, comments,
  footnotes, endnotes; no envelope) — see \`docx info schema\`. Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx read input.docx
  docx read input.docx --from p3 --to p20
  docx read input.docx --accepted --comments
  docx read input.docx --baseline
  docx read input.docx --ast | jq '.blocks[] | select(.type == "paragraph")'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			ast: { type: "boolean" },
			from: { type: "string" },
			to: { type: "string" },
			accepted: { type: "boolean" },
			baseline: { type: "boolean" },
			current: { type: "boolean" },
			comments: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const ast = Boolean(parsed.values.ast);
	const from = parsed.values.from as string | undefined;
	const to = parsed.values.to as string | undefined;
	const accepted = Boolean(parsed.values.accepted);
	const baseline = Boolean(parsed.values.baseline);
	const current = Boolean(parsed.values.current);
	const showComments = Boolean(parsed.values.comments);

	if (ast && (from || to || accepted || baseline || current || showComments)) {
		return fail(
			"USAGE",
			"--from, --to, --accepted, --baseline, --current, and --comments are Markdown-only and cannot be combined with --ast",
			HELP,
		);
	}

	const view = resolveView({ accepted, baseline, current });
	if (!view) {
		return fail(
			"USAGE",
			"--accepted, --baseline, and --current are mutually exclusive",
			HELP,
		);
	}

	const docView = await openOrFail(path);
	if (typeof docView === "number") return docView;

	// Hashes are content-addressed image identifiers — read --ast surfaces
	// them on each ImageRun, and read --markdown uses them as the URL in
	// `![alt](<sha256>.<ext>)` so a round-trip through `insert --markdown`
	// can reuse existing media parts instead of re-fetching. Both paths
	// need the hashes populated, so enrich up front for either output mode.
	await new Images(docView).enrichHashes();

	if (ast) {
		await respond(docView.body);
		return EXIT.OK;
	}

	try {
		const rendered = renderMarkdown(docView.body, {
			from,
			to,
			view,
			showComments,
			defaultSizeHalfPoints: docView.styles?.defaultSizeHalfPoints(),
		});
		await writeStdout(rendered);
		return EXIT.OK;
	} catch (err) {
		if (err instanceof MarkdownLocatorError) {
			return fail("INVALID_LOCATOR", err.message);
		}
		throw err;
	}
}
