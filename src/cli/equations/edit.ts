import {
	EquationNotFoundError,
	EquationParseError,
	EquationStaleError,
	Equations,
} from "@core/equation";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respondAck,
	respondEditDryRun,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx equations edit — replace or retype an existing equation

Usage:
  docx equations edit FILE --at eqN --equation NEW_LATEX [options]
  docx equations edit FILE --at eqN --display [options]
  docx equations edit FILE --at eqN --inline [options]

Examples:
  docx equations edit doc.docx --at eq0 --equation "x = \\\\frac{-b}{2a}"
  docx equations edit doc.docx --at eq0 --display
  docx equations edit doc.docx --at eq0 --equation "y^3" --inline

Locator (required):
  --at eqN          The equation to edit, addressed by document order. Discover
                    ids with \`docx read FILE --ast\` (EquationRun.id).

Content (at least one required):
  --equation LATEX  Replace the equation's content with new LaTeX. Goes through
                    temml → MathML → OMML.
  --display         Switch the equation to display mode (block, $$…$$). Can be
                    combined with --equation, or used alone to toggle mode.
  --inline          Switch to inline mode ($…$). Mutex with --display.

Options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record this edit as a tracked change even when the
                    document's track-changes toggle is off.
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Under track-changes, the edit records a paired tracked delete + insert —
accept keeps the NEW equation, reject restores the OLD.

Output:
  Prints a one-line confirmation on success (exit 0) — an in-place edit shifts
  nothing, so the eqN is unchanged. --verbose prints {ok:true, operation:"edit",
  path, locator}. Errors print {code, error, hint?} with a nonzero exit.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	equation: { type: "string" },
	display: { type: "boolean" },
	inline: { type: "boolean" },
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

	const locator = parsed.values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at eqN", HELP);

	// `--equation` swaps the LaTeX (and optionally display mode); `--display` /
	// `--inline` alone toggle the display mode but keep the existing LaTeX. At
	// least one must change something, else it's a no-op error.
	const latex = parsed.values.equation as string | undefined;
	const displayFlag = parsed.values.display as boolean | undefined;
	const inlineFlag = parsed.values.inline as boolean | undefined;

	if (displayFlag && inlineFlag) {
		return fail("USAGE", "--display and --inline are mutually exclusive", HELP);
	}
	if (latex === undefined && !displayFlag && !inlineFlag) {
		return fail(
			"USAGE",
			"--equation requires --equation NEW_LATEX, --display, or --inline",
			HELP,
		);
	}
	const display: boolean | undefined = displayFlag
		? true
		: inlineFlag
			? false
			: undefined;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const outputPath = parsed.values.output as string | undefined;
	if (parsed.values["dry-run"]) {
		return respondEditDryRun(filePath, locator, outputPath);
	}

	// Tracked when the document toggle is on OR `--track` forces it (Word records
	// equation edits as a paired del/ins); `--author` threads to the change meta.
	try {
		new Equations(document).edit(locator, {
			latex,
			display,
			author: parsed.values.author as string | undefined,
			track: resolveTracked(document, Boolean(parsed.values.track)),
		});
	} catch (error) {
		if (error instanceof EquationNotFoundError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationStaleError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationParseError) {
			return fail(
				"USAGE",
				`Could not parse LaTeX equation: ${error.message}`,
				"Check the LaTeX syntax — temml accepts most KaTeX/MathJax LaTeX.",
			);
		}
		throw error;
	}

	await document.save(outputPath);
	await respondAck({
		ok: true,
		operation: "edit",
		path: outputPath ?? filePath,
		locator,
	});
	return EXIT.OK;
}
