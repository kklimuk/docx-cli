import {
	findMarginalRef,
	isMarginalLocator,
	isRelationshipLocator,
	marginalConfig,
} from "@core";
import { XmlNode } from "@core/parser";
import { Raw } from "@core/raw";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { failRelationshipNotFound } from "./rels";

const HELP = `docx raw get — print the exact XML of a block, section, header/footer, or relationship

Usage:
  docx raw get FILE --at LOCATOR [--json]

Examples:
  docx raw get doc.docx --at p3          # a paragraph's exact <w:p …>…</w:p>
  docx raw get doc.docx --at s0          # a section's <w:sectPr>…</w:sectPr>
  docx raw get doc.docx --at t0:r0c0:p0  # a paragraph inside a table cell
  docx raw get doc.docx --at ftr0        # a footer's <w:ftr>…</w:ftr> (id from footers list / read)
  docx raw get doc.docx --at rels        # every <Relationship> (ids, types, targets)
  docx raw get doc.docx --at rId5        # one relationship
  # then modify the XML and put it back:
  docx raw replace doc.docx --at p3 --xml '<w:p …>…</w:p>'

Options:
  --at LOCATOR   What to serialize: pN | tN | sN | a cell-chained form
                 (t0:r0c0:p0) | hdrN | ftrN (a page header/footer) | rIdN | rels.
                 Required.
                 (Package parts are not locators — use "docx raw part get".)
  --json         Emit {"locator", "xml"} instead of bare XML.

Output:
  The exact XML, exactly as stored — this is what --find must match byte for
  byte in "docx raw edit". Unknown locators exit 3.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	json: { type: "boolean" },
	help: { type: "boolean", short: "h" },
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	const locator = values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);
	const partGuard = await rejectPartAsLocator(locator);
	if (partGuard !== undefined) return partGuard;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	if (locator === "rels" || isRelationshipLocator(locator)) {
		const xml = new Raw(document).serializeRelationships(
			locator === "rels" ? undefined : locator,
		);
		if (xml === undefined) {
			return failRelationshipNotFound(filePath, locator);
		}
		return emit(locator, xml, Boolean(values.json));
	}

	if (isMarginalLocator(locator)) {
		const ref = findMarginalRef(document, locator);
		if (!ref) {
			return fail(
				"BLOCK_NOT_FOUND",
				`No ${locator} in the document`,
				`docx headers list / footers list ${filePath} print every header/footer with its id; docx read surfaces them as docx:header/docx:footer hints.`,
			);
		}
		const root = XmlNode.findRoot(ref.tree, marginalConfig(ref.kind).rootTag);
		const xml = root ? XmlNode.serialize([root]) : XmlNode.serialize(ref.tree);
		return emit(locator, xml, Boolean(values.json));
	}

	const reference = await resolveBlockOrFail(document, locator);
	if (typeof reference === "number") return reference;
	return emit(
		locator,
		new Raw(document).serializeBlock(reference),
		Boolean(values.json),
	);
}

async function emit(
	locator: string,
	xml: string,
	json: boolean,
): Promise<number> {
	if (json) {
		await respond({ locator, xml });
		return EXIT.OK;
	}
	await writeStdout(`${xml}\n`);
	return EXIT.OK;
}

/** Parts are the package's files, not document content — they never ride the
 *  locator system. Shared by get/replace/edit so a `--at part:…` attempt gets
 *  the recovery path, not a confusing stale-locator error. */
export async function rejectPartAsLocator(
	locator: string,
): Promise<number | undefined> {
	if (locator !== "parts" && !locator.startsWith("part:")) return undefined;
	return fail(
		"USAGE",
		"Package parts aren't locators — locators address document content",
		"Use the part noun instead: docx raw part list|get|add|replace|edit FILE --name word/… .",
	);
}
