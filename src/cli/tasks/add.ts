import type { InsertSpec } from "@core";
import { parseParagraphOptions, type RawValues } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import {
	decodeInlineEscapes,
	parseRunsArg,
	rejectShellMangledValue,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx tasks add — insert a GFM task-list item (checkbox)

Usage:
  docx tasks add FILE --after LOCATOR --text LABEL [--checked] [options]
  docx tasks add FILE --before LOCATOR --runs JSON [--unchecked] [options]
  docx tasks add FILE (--at-start | --at-end) --text LABEL [options]

Examples:
  docx tasks add doc.docx --after p0 --text "buy groceries"
  docx tasks add doc.docx --after p1 --text "pay rent" --checked
  docx tasks add doc.docx --after p2 --text "nested item" --checked --list-level 1
  docx tasks add doc.docx --at-end --runs '[{"type":"text","text":"ship it"}]'

Placement (exactly one required):
  --after LOCATOR   Insert after the block at LOCATOR (a pN / tN / cell paragraph)
  --before LOCATOR  Insert before the block at LOCATOR
  --at-start        Insert at the very top of the document (no locator needed)
  --at-end          Insert at the very end, before the trailing section properties

Content (exactly one of):
  --text LABEL      The task label (literal prose — a \\n becomes a line break,
                    a \\t a tab; a markdown-looking value lands verbatim, so ** /
                    [](…) go in literally).
  --runs JSON       Custom runs for the label (Run[] JSON).

State:
  --checked         Start the checkbox done (☒).
  --unchecked       Start the checkbox empty (☐). This is the default.

Nesting / paragraph options (apply to the task paragraph):
  --list-level N     Nesting level, integer 0-8 (default 0). If the anchor is
                     already a list, the new task inherits its numId so
                     consecutive \`tasks add\` calls build ONE contiguous list;
                     otherwise a fresh bullet list is allocated.
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

Toggle an existing task with \`docx tasks check\` / \`docx tasks uncheck\`.

Output:
  Prints the inserted block's locator (pN), one per line. --verbose prints
  {ok:true, operation:"insert", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
	text: { type: "string" },
	runs: { type: "string" },
	checked: { type: "boolean" },
	unchecked: { type: "boolean" },
	"list-level": { type: "string" },
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

	const checked = await resolveCheckedState(parsed.values);
	if (typeof checked === "number") return checked;

	const spec = await resolveTaskSpec(parsed.values);
	if (typeof spec === "number") return spec;

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;
	// A task item is a bullet-list paragraph carrying a leading checkbox SDT.
	// `taskState` is the paragraph option the Insert lens keys on to prepend the
	// <w:sdt><w14:checkbox/> and resolve the list numId (inherit from a list
	// anchor, else allocate fresh).
	paragraphOptions.taskState = checked ? "checked" : "unchecked";

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

/** Resolve `--checked` / `--unchecked` (default unchecked). Passing both is a
 *  contradiction, not a silent last-wins. */
async function resolveCheckedState(
	values: RawValues,
): Promise<boolean | number> {
	const checked = Boolean(values.checked);
	const unchecked = Boolean(values.unchecked);
	if (checked && unchecked) {
		return fail(
			"USAGE",
			"Pass at most one of --checked or --unchecked (default: unchecked)",
			HELP,
		);
	}
	return checked;
}

/** Resolve `--text LABEL` (literal prose, decoded like insert's --text) or
 *  `--runs JSON` into the task paragraph's content spec. Exactly one required. */
async function resolveTaskSpec(
	values: RawValues,
): Promise<Extract<InsertSpec, { kind: "text" } | { kind: "runs" }> | number> {
	const text = values.text as string | undefined;
	const runsJson = values.runs as string | undefined;
	if ((text === undefined) === (runsJson === undefined)) {
		return fail("USAGE", "Pass exactly one of --text or --runs", HELP);
	}
	if (text !== undefined) {
		// Task labels decode inline whitespace escapes and land verbatim (a
		// markdown-looking label is kept literally). We still refuse a shell-gutted
		// currency value, exactly like `insert --text`.
		const decoded = decodeInlineEscapes(text);
		const mangled = await rejectShellMangledValue(decoded, "--text");
		if (typeof mangled === "number") return mangled;
		return { kind: "text", text: decoded, format: {} };
	}
	const runs = await parseRunsArg(runsJson as string);
	if (typeof runs === "number") return runs;
	return { kind: "runs", runs };
}
