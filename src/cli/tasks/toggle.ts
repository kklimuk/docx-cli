import { Edit, EditError } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveTracked,
	respondAck,
	respondEditDryRun,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const OPTION_SPEC = {
	at: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

/** Shared back-end for `docx tasks check` (checked=true) and `docx tasks
 *  uncheck` (checked=false): flip an existing task-list item's checkbox in
 *  place at `--at pN`, via the core `Edit.taskToggle`. Under track-changes (the
 *  doc toggle OR `--track`) it emits Word's canonical toggle shape (an ins/del
 *  glyph pair inside the checkbox SDT + a `w14:checked` flip), surfaced by
 *  `track-changes list` as a `checkboxToggle` revision. Both verbs pass their
 *  own HELP so `--help` prints the right screen. */
export async function toggleTask(
	args: string[],
	checked: boolean,
	help: string,
): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, help);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", help);

	const locator = parsed.values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR (pN)", help);

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const outputPath = parsed.values.output as string | undefined;
	if (parsed.values["dry-run"]) {
		return respondEditDryRun(filePath, locator, outputPath);
	}

	const track = resolveTracked(document, Boolean(parsed.values.track));
	const authorFlag = parsed.values.author as string | undefined;

	const blockRef = await resolveBlockOrFail(document, locator);
	if (typeof blockRef === "number") return blockRef;

	try {
		new Edit(document).taskToggle(blockRef, checked, { authorFlag, track });
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
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
