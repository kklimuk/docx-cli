import { openDocView, PkgError } from "@core";
import { w } from "@core/jsx";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx track-changes — toggle the document's tracked-changes mode

Usage:
  docx track-changes FILE on|off [options]

Sets <w:trackChanges/> in word/settings.xml. When on, Word records new
edits as tracked changes. Existing <w:ins>/<w:del> markers are unaffected.

Options:
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

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (openError) {
		if (openError instanceof PkgError) {
			if (openError.code === "FILE_NOT_FOUND") {
				return fail("FILE_NOT_FOUND", openError.message);
			}
			if (openError.code === "NOT_A_ZIP") {
				return fail("NOT_A_ZIP", openError.message);
			}
		}
		throw openError;
	}

	const settingsXml = view.pkg.hasPart(SETTINGS_PART)
		? await view.pkg.readText(SETTINGS_PART)
		: null;
	const settingsTree = settingsXml ? XmlNode.parse(settingsXml) : [];
	let settingsRoot = XmlNode.findRoot(settingsTree, "w:settings");
	if (!settingsRoot) {
		settingsRoot = (<w.settings {...{ "xmlns:w": NS_W }} />) as XmlNode;
		settingsTree.push(settingsRoot);
	}

	const hasTrackChanges = settingsRoot.children.some(
		(child) => child.tag === "w:trackChanges",
	);

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "track-changes",
			dryRun: true,
			path,
			mode,
			previouslyOn: hasTrackChanges,
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

	view.pkg.writeText(
		SETTINGS_PART,
		XmlNode.serializeWithDeclaration(settingsTree),
	);

	if (!settingsXml) {
		registerSettingsPart(view);
	}

	view.pkg.writeText(
		"word/_rels/document.xml.rels",
		XmlNode.serializeWithDeclaration(view.relationshipsTree),
	);
	view.pkg.writeText(
		"[Content_Types].xml",
		XmlNode.serializeWithDeclaration(view.contentTypesTree),
	);
	await view.pkg.save();

	await respond({
		ok: true,
		operation: "track-changes",
		path,
		mode,
		previouslyOn: hasTrackChanges,
	});
	return EXIT.OK;
}

function registerSettingsPart(
	view: Awaited<ReturnType<typeof openDocView>>,
): void {
	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (relationships) {
		const alreadyLinked = relationships.children.some(
			(child) =>
				child.tag === "Relationship" &&
				child.getAttribute("Type") === SETTINGS_REL_TYPE,
		);
		if (!alreadyLinked) {
			relationships.children.push(
				new XmlNode("Relationship", {
					Id: nextRelationshipId(relationships),
					Type: SETTINGS_REL_TYPE,
					Target: "settings.xml",
				}),
			);
		}
	}

	const types = XmlNode.findRoot(view.contentTypesTree, "Types");
	if (types) {
		const exists = types.children.some(
			(child) =>
				child.tag === "Override" &&
				child.getAttribute("PartName") === `/${SETTINGS_PART}`,
		);
		if (!exists) {
			types.children.push(
				new XmlNode("Override", {
					PartName: `/${SETTINGS_PART}`,
					ContentType: SETTINGS_CONTENT_TYPE,
				}),
			);
		}
	}
}

function nextRelationshipId(relationships: XmlNode): string {
	let highest = 0;
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		const id = child.getAttribute("Id");
		if (!id) continue;
		const match = id.match(/^rId(\d+)$/);
		if (!match) continue;
		const numeric = Number(match[1]);
		if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
	}
	return `rId${highest + 1}`;
}
