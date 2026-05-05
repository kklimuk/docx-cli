import { saveDocView } from "@core";
import { w } from "@core/jsx";
import { registerPart } from "@core/package";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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
  -h, --help        Show this help

Examples:
  docx track-changes doc.docx on
  docx track-changes doc.docx off
`;

const SETTINGS_PART = "word/settings.xml";
const SETTINGS_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

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

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const wasMissing = view.settingsTree === undefined;
	if (!view.settingsTree) view.settingsTree = [];
	let settingsRoot = XmlNode.findRoot(view.settingsTree, "w:settings");
	if (!settingsRoot) {
		settingsRoot = (<w.settings {...{ "xmlns:w": NS_W }} />) as XmlNode;
		view.settingsTree.push(settingsRoot);
	}

	const hasTrackChanges = settingsRoot.children.some(
		(child) => child.tag === "w:trackChanges",
	);
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

	if (mode === "on" && !hasTrackChanges) {
		settingsRoot.children.unshift((<w.trackChanges />) as XmlNode);
	} else if (mode === "off" && hasTrackChanges) {
		settingsRoot.children = settingsRoot.children.filter(
			(child) => child.tag !== "w:trackChanges",
		);
	}

	if (wasMissing) {
		registerPart(view.relationshipsTree, view.contentTypesTree, {
			partName: SETTINGS_PART,
			contentType: SETTINGS_CONTENT_TYPE,
			relationshipType: SETTINGS_REL_TYPE,
			target: "settings.xml",
		});
	}

	await saveDocView(view, outputPath);

	await respond({
		ok: true,
		operation: "track-changes",
		path: outputPath ?? path,
		mode,
		previouslyOn: hasTrackChanges,
	});
	return EXIT.OK;
}
