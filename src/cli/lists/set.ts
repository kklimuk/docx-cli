import {
	type Document,
	FORMAT_TO_NUMFMT,
	type ListFormat,
	ListOperationError,
	Lists,
} from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx lists set — renumber a numbered list (start value, glyph format, restart/continue)

Usage:
  docx lists set FILE --at pN [options]

--at names any item of a NUMBERED (ordered) list. --start/--format change the whole
list (every paragraph sharing its numbering); --restart/--continue act from the
addressed item onward. Bulleted lists aren't numbered, so they're rejected.

Options:
  --start N      Start the list's numbering at N (a positive integer)
  --format FMT   Glyph style: decimal | lower-alpha | upper-alpha | lower-roman | upper-roman
                 (decimal = 1.2.3., lower-alpha = a.b.c., upper-roman = I.II.III., …)
  --restart      Begin a FRESH list at this item: split it off with its own numbering
                 (combine with --start/--format to set the new list's first number/style)
  --continue     Continue the PREVIOUS list's numbering here instead of restarting at 1
  -o, --output PATH / --dry-run / -v, --verbose / -h, --help

--restart and --continue are mutually exclusive; --continue can't combine with
--start/--format (it adopts the previous list's numbering). List numbering edits
are applied directly (untracked) — Word records no revision for them.

Examples:
  docx lists set report.docx --at p12 --start 5
  docx lists set report.docx --at p12 --format upper-roman
  docx lists set report.docx --at p20 --restart --start 1
  docx lists set report.docx --at p20 --continue
`;

type Plan = {
	start?: number;
	format?: ListFormat;
	restart: boolean;
	continue: boolean;
};

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			start: { type: "string" },
			format: { type: "string" },
			restart: { type: "boolean" },
			continue: { type: "boolean" },
			...SAVE_FLAGS,
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
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const at = parsed.values.at as string | undefined;
	if (!at) {
		return fail("USAGE", "Missing --at LOCATOR (a list item, e.g. p5)", HELP);
	}

	const plan = buildPlan(parsed.values);
	if (typeof plan === "string") return fail("USAGE", plan);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const blockRef = await resolveBlockOrFail(document, at);
	if (typeof blockRef === "number") return blockRef;

	const orderedError = validateOrderedList(document, at);
	if (orderedError) return fail("USAGE", orderedError);

	const outputPath = parsed.values.output as string | undefined;
	if (parsed.values["dry-run"]) {
		await respond({
			operation: "lists.set",
			dryRun: true,
			path,
			at,
			applied: appliedList(plan),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const lists = new Lists(document);
	try {
		if (plan.continue) {
			lists.continue(blockRef);
		} else if (plan.restart) {
			lists.restart(blockRef, plan.start ?? 1);
			// cloneListDefinition copies the source glyph; an explicit --format
			// re-styles the freshly split list (blockRef now points at its numId).
			if (plan.format) lists.setFormat(blockRef, plan.format);
		} else {
			if (plan.start !== undefined) lists.setStart(blockRef, plan.start);
			if (plan.format) lists.setFormat(blockRef, plan.format);
		}
	} catch (error) {
		if (error instanceof ListOperationError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		throw error;
	}

	await document.save(outputPath);

	const destination = outputPath ?? path;
	await respondAck({
		ok: true,
		operation: "lists.set",
		path: destination,
		locator: at,
		applied: appliedList(plan),
	});
	return EXIT.OK;
}

function buildPlan(values: Record<string, unknown>): Plan | string {
	const restart = Boolean(values.restart);
	const continueFlag = Boolean(values.continue);
	if (restart && continueFlag) {
		return "Pass only one of --restart / --continue";
	}

	let start: number | undefined;
	if (values.start !== undefined) {
		const parsedStart = Number(values.start);
		if (!Number.isInteger(parsedStart) || parsedStart < 1) {
			return `--start must be a positive integer (got ${values.start})`;
		}
		start = parsedStart;
	}

	let format: ListFormat | undefined;
	if (values.format !== undefined) {
		const candidate = String(values.format);
		if (!Object.hasOwn(FORMAT_TO_NUMFMT, candidate)) {
			return `--format must be one of ${Object.keys(FORMAT_TO_NUMFMT).join(", ")} (got ${candidate})`;
		}
		format = candidate as ListFormat;
	}

	if (continueFlag && (start !== undefined || format !== undefined)) {
		return "--continue adopts the previous list's numbering — it can't combine with --start/--format";
	}
	if (
		!restart &&
		!continueFlag &&
		start === undefined &&
		format === undefined
	) {
		return "Nothing to do — pass --start, --format, --restart, or --continue";
	}

	return { start, format, restart, continue: continueFlag };
}

/** Reject a locator that isn't a numbered list item before any mutation, with a
 * hint that points the agent at the right list kind. */
function validateOrderedList(document: Document, at: string): string | null {
	const block = document.body.findBlockById(at);
	if (!block || block.type !== "paragraph" || !block.list) {
		return `${at} is not a list item — list numbering controls only apply to numbered lists`;
	}
	if (!block.list.ordered) {
		return `${at} is a bulleted list — numbering controls (start/format/restart/continue) apply to numbered (ordered) lists`;
	}
	return null;
}

function appliedList(plan: Plan): string[] {
	const applied: string[] = [];
	if (plan.continue) applied.push("continue");
	if (plan.restart) applied.push("restart");
	if (plan.start !== undefined) applied.push(`start=${plan.start}`);
	if (plan.format) applied.push(`format=${plan.format}`);
	return applied;
}
