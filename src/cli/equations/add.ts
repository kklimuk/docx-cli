import type { InsertSpec } from "@core";
import { parseParagraphOptions, type RawValues } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import {
	EXIT,
	fail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx equations add — insert a LaTeX equation

Usage:
  docx equations add FILE --after LOCATOR --equation LATEX [--display] [options]
  docx equations add FILE --before LOCATOR --equation LATEX [options]
  docx equations add FILE (--at-start | --at-end) --equation LATEX [options]

Placement (exactly one required):
  --after LOCATOR   Insert after the block at LOCATOR (a pN / tN / cell paragraph)
  --before LOCATOR  Insert before the block at LOCATOR
  --at-start        Insert at the very top of the document (no locator needed)
  --at-end          Insert at the very end, before the trailing section properties

Content (required):
  --equation LATEX  Math equation from LaTeX. Goes through temml (KaTeX/MathJax-
                    compatible LaTeX dialect) → MathML → OMML. Round-trips as
                    $LATEX$ / $$LATEX$$ in markdown.
  --display         Emit a block-mode (display) equation ($$…$$). Omit for an
                    inline equation ($…$).

Paragraph options (apply to the equation paragraph):
  --style NAME       Paragraph style      --alignment ALIGN  left|center|right|justify
  --space-before PT  --space-after PT     --line-spacing N
  --indent-left IN   --indent-right IN    --first-line IN   --hanging IN

Options:
  --track           Record the insert as a tracked change even when the
                    document's track-changes toggle is off.
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: the minted locator)
  -h, --help        Show this help

Examples:
  docx equations add doc.docx --after p4 --equation "x^2 + y^2 = r^2"
  docx equations add doc.docx --after p4 --equation "\\frac{-b}{2a}" --display
  docx equations add doc.docx --at-end --equation "E=mc^2"

Output:
  Prints the inserted block's locator (pN), one per line. --verbose prints
  {ok:true, operation:"insert", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
	equation: { type: "string" },
	display: { type: "boolean" },
	style: { type: "string" },
	alignment: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const placement = await parseTargetPlacement(parsed.values, HELP);
	if (typeof placement === "number") return placement;

	const spec = await resolveEquationSpec(parsed.values);
	if (typeof spec === "number") return spec;

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return placeSpec({
		filePath,
		placement,
		spec,
		paragraphOptions,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	});
}

/** Resolve `--equation LATEX` (required) + `--display` into an `equation` spec.
 *  LaTeX is NOT run through `decodeInlineEscapes` — a `\n` in `\nabla` must
 *  survive as backslash-n, not become a newline. */
async function resolveEquationSpec(
	values: RawValues,
): Promise<Extract<InsertSpec, { kind: "equation" }> | number> {
	const latex = values.equation as string | undefined;
	if (latex === undefined) {
		return fail("USAGE", "Missing --equation LATEX", HELP);
	}
	return { kind: "equation", latex, display: Boolean(values.display) };
}
