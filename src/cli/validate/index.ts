import type { Pkg } from "@core/ast/document/package";
import type { XmlNode } from "@core/parser";
import {
	validateWmlXml,
	validationXmlFor,
	WML_ROOT_TAGS,
} from "@core/raw/validate";
import {
	EXIT,
	fail,
	openPkgOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx validate — schema-check a document against ECMA-376

Usage:
  docx validate FILE [--json]

Validates every WordprocessingML part in the package — document, styles,
numbering, settings, footnotes, endnotes, comments, every header/footer,
fontTable, webSettings, … (any part whose root wml.xsd declares) — against
the bundled ECMA-376 5th-edition TRANSITIONAL schemas (the profile Word
itself writes) after markup-compatibility preprocessing (mc:Ignorable
extension markup is skipped, exactly as a consumer would). Exit 0 = clean;
exit 1 = errors, listed per part.

Useful before/after hand-made XML changes (docx raw, or edits made outside
docx-cli) and as a "will Word open this?" first check — note Word TOLERATES
some schema violations, so a finding here is "invalid OOXML," not always
"Word rejects it."
`;

const OPTION_SPEC = {
	json: { type: "boolean" },
	help: { type: "boolean", short: "h" },
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);

	const pkg = await openPkgOrFail(filePath);
	if (typeof pkg === "number") return pkg;

	const parts = await collectWmlParts(pkg);

	if (isStrictProfile(parts)) {
		const note =
			"strict-profile document (ISO/IEC 29500 Strict namespaces) — the bundled schemas cover the TRANSITIONAL profile Word writes, so no verdict is possible";
		if (values.json) {
			await respond({ path: filePath, skipped: "strict-profile", note });
		} else {
			await writeStdout(`skipped — ${note}\n`);
		}
		return EXIT.OK;
	}

	const results = parts.map((part) => ({
		name: part.name,
		issues: validateWmlXml(validationXmlFor(part.root)),
	}));
	const errorCount = results.reduce((sum, part) => sum + part.issues.length, 0);

	if (values.json) {
		await respond({
			path: filePath,
			valid: errorCount === 0,
			errors: errorCount,
			parts: results.map((part) => ({
				part: part.name,
				issues: part.issues,
			})),
		});
		return errorCount === 0 ? EXIT.OK : EXIT.GENERAL_ERROR;
	}

	if (errorCount === 0) {
		await writeStdout(
			`valid — ${results.length} part${results.length === 1 ? "" : "s"} checked\n`,
		);
		return EXIT.OK;
	}
	const lines: string[] = [];
	for (const part of results) {
		if (part.issues.length === 0) continue;
		lines.push(
			`${part.name}: ${part.issues.length} error${part.issues.length === 1 ? "" : "s"}`,
		);
		for (const issue of part.issues) {
			lines.push(
				`  ${issue.message}${issue.line !== undefined ? ` (line ${issue.line})` : ""}`,
			);
		}
	}
	await writeStdout(`${lines.join("\n")}\n`);
	return EXIT.GENERAL_ERROR;
}

/** Every package part whose root is a WML root, parsed once — the parse
 *  feeds both the strict-profile sniff and validation. Sorted with
 *  word/document.xml first (the part a reader looks for), the rest by name. */
async function collectWmlParts(
	pkg: Pkg,
): Promise<{ name: string; root: XmlNode }[]> {
	const names = pkg
		.listParts()
		.filter((name) => name.endsWith(".xml"))
		.sort((a, b) => {
			if (a === "word/document.xml") return -1;
			if (b === "word/document.xml") return 1;
			return a.localeCompare(b);
		});
	const parts: { name: string; root: XmlNode }[] = [];
	for (const name of names) {
		const tree = await pkg.readPart(name);
		const root = tree?.find((node) => !node.isText && node.tag !== "?xml");
		if (!root || !WML_ROOT_TAGS.has(root.tag)) continue;
		parts.push({ name, root });
	}
	return parts;
}

/** ISO Strict re-namespaces every part (`http://purl.oclc.org/ooxml/…`), so the
 *  transitional schemas see unknown content and every finding would be noise. */
function isStrictProfile(parts: { name: string; root: XmlNode }[]): boolean {
	const documentPart = parts.find((part) => part.name === "word/document.xml");
	return (
		documentPart?.root
			.getAttribute("xmlns:w")
			?.startsWith("http://purl.oclc.org/ooxml") ?? false
	);
}
