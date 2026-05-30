import { TrackChanges } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx track-changes — toggle the document's tracked-changes mode

Usage:
  docx track-changes FILE on|off [options]

Sets <w:trackChanges/> in word/settings.xml. When on, this CLI's
insert/edit/delete/replace commands also emit <w:ins>/<w:del> markers
(attributed to $DOCX_AUTHOR) so changes remain reviewable. Existing
<w:ins>/<w:del> markers are unaffected by the toggle itself.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Examples:
  docx track-changes doc.docx on
  docx track-changes doc.docx off
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			output: { type: "string", short: "o" },
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
			ok: true,
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
