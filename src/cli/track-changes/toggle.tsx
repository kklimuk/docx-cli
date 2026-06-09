import { TrackChanges } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx track-changes — toggle the document's tracked-changes mode

Usage:
  docx track-changes FILE on|off [options]

Toggle only sets (on) or clears (off) the <w:trackChanges/> flag in
word/settings.xml — nothing else. It does not author any <w:ins>/<w:del>
markers, and existing markers are unaffected by the toggle itself.

When on, the SUBSEQUENT insert/edit/delete/replace commands emit
<w:ins>/<w:del> markers so changes remain reviewable. The --author
attribution (--author NAME, else $DOCX_AUTHOR) is read by those mutating
commands, NOT by toggle — toggle takes no --author.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  mode, previouslyOn}. --dry-run prints the preview {operation, dryRun, path,
  mode, previouslyOn}. Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx track-changes doc.docx on
  docx track-changes doc.docx off
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
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

	const mode = parsed.positionals[1];
	if (mode !== "on" && mode !== "off") {
		return fail(
			"USAGE",
			`Expected "on" or "off", got: ${mode ?? "<nothing>"}`,
			HELP,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const hasTrackChanges = document.settings?.isTrackChangesEnabled() ?? false;
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "track-changes",
			dryRun: true,
			path,
			mode,
			previouslyOn: hasTrackChanges,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	if (mode === "on" || hasTrackChanges) {
		new TrackChanges(document).setEnabled(mode === "on");
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "track-changes",
		path: outputPath ?? path,
		mode,
		previouslyOn: hasTrackChanges,
	});
	return EXIT.OK;
}
