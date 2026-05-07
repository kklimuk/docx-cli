import { enrichImageHashes } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import { MarkdownLocatorError, renderMarkdown } from "./markdown";

const HELP = `docx read — print AST as JSON, or render document body as Markdown

Usage:
  docx read FILE [options]

Options:
  --markdown        Render the body as GitHub-flavored Markdown
                    (instead of JSON). Locators are emitted as
                    <!-- pN --> HTML comments after each paragraph and
                    inside cell content (invisible in rendered view,
                    but parseable from raw markdown).
  --from LOC        Start markdown rendering at top-level block LOC (inclusive)
  --to LOC          End markdown rendering at top-level block LOC (inclusive)
                    Accepts pN, tN, tN:rRcC[:pK[:S-E]], pN:S-E, pN:S-pM:E.
                    Cell/span/range locators collapse to their enclosing
                    top-level block (the table or paragraph).
  --accepted        Default with --markdown: render the post-accept view —
                    drop subtractive wrappers (<w:del>, <w:moveFrom>), inline
                    additive wrappers (<w:ins>, <w:moveTo>) as plain text,
                    no markers/refs. Kept as an explicit alias for clarity.
  --baseline        With --markdown, render the pre-change view: drop
                    additive wrappers (<w:ins>, <w:moveTo>), inline
                    subtractive wrappers (<w:del>, <w:moveFrom>) as plain
                    text, no markers/refs.
  --current         With --markdown, render the raw concatenation: additive
                    wrappers as {++text++}[^tcN] and subtractive as
                    {--text--}[^tcN] (CriticMarkup); the [^tcN] footnote
                    spells out the kind (insertion / deletion / moveTo /
                    moveFrom). Mutually exclusive with --accepted/--baseline.
  --comments        With --markdown, append [^cN] after each commented span
                    and emit a footnote definition for each comment at the
                    end of the output (author, date, body).
  -h, --help        Show this help

Examples:
  docx read input.docx
  docx read input.docx --markdown
  docx read input.docx --markdown --from p3 --to p20
  docx read input.docx --markdown --accepted --comments
  docx read input.docx --markdown --baseline
  docx read input.docx | jq '.blocks[] | select(.type == "paragraph")'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				markdown: { type: "boolean" },
				from: { type: "string" },
				to: { type: "string" },
				accepted: { type: "boolean" },
				baseline: { type: "boolean" },
				current: { type: "boolean" },
				comments: { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (e) {
		return fail("USAGE", e instanceof Error ? e.message : String(e), HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const markdown = Boolean(parsed.values.markdown);
	const from = parsed.values.from as string | undefined;
	const to = parsed.values.to as string | undefined;
	const accepted = Boolean(parsed.values.accepted);
	const baseline = Boolean(parsed.values.baseline);
	const current = Boolean(parsed.values.current);
	const showComments = Boolean(parsed.values.comments);

	if (
		!markdown &&
		(from || to || accepted || baseline || current || showComments)
	) {
		return fail(
			"USAGE",
			"--from, --to, --accepted, --baseline, --current, and --comments require --markdown",
			HELP,
		);
	}

	const viewFlagCount =
		(accepted ? 1 : 0) + (baseline ? 1 : 0) + (current ? 1 : 0);
	if (viewFlagCount > 1) {
		return fail(
			"USAGE",
			"--accepted, --baseline, and --current are mutually exclusive",
			HELP,
		);
	}

	const view = current ? "current" : baseline ? "baseline" : "accepted";

	const docView = await openOrFail(path);
	if (typeof docView === "number") return docView;

	if (markdown) {
		try {
			const rendered = renderMarkdown(docView.doc, {
				from,
				to,
				view,
				showComments,
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

	await enrichImageHashes(docView);
	await respond(docView.doc);
	return EXIT.OK;
}
