import { isRelationshipLocator } from "@core";
import {
	fail,
	openOrFail,
	resolveBlockOrFail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
} from "../respond";
import { runRawReplaceBatch } from "./batch";
import { failTrackRejected, replaceBlockWithXml } from "./commit";
import { rejectPartAsLocator } from "./get";
import { runRelationshipReplace } from "./rels";
import { readXmlSource } from "./xml-source";

const HELP = `docx raw replace — replace one block, section, or relationship with raw OOXML

Usage:
  docx raw replace FILE --at LOCATOR (--xml STR | --xml-file PATH|-)
  docx raw replace FILE --batch FILE.jsonl

Examples:
  # The patch loop — read the exact XML, modify it, put it back:
  docx raw get report.docx --at s0          # exact <w:sectPr>…</w:sectPr>
  docx raw replace report.docx --at s0 --xml '<w:sectPr>…patched…</w:sectPr>'
  # 1-for-N: replace one paragraph with two
  docx raw replace doc.docx --at p3 --xml '<w:p>…</w:p><w:p>…</w:p>'
  # Retarget a relationship (keep its Id — a rename would dangle references)
  docx raw replace doc.docx --at rId5 --xml '<Relationship Id="rId5" Type="…" Target="https://new" TargetMode="External"/>'
  # Many replacements from one read:
  #   swaps.jsonl:  {"at":"p3","xml":"<w:p>…</w:p>"}
  docx raw replace doc.docx --batch swaps.jsonl

Options:
  --at LOCATOR       The target: pN | tN | sN | a cell-chained form
                     (t0:r0c0:p0) | rIdN (a relationship).
                     (Package parts are not locators — "docx raw part replace".)
  --xml STR          The replacement, inline and VERBATIM (no escape decoding —
                     XML has its own escapes: &#10; &lt; &amp;)
  --xml-file PATH    Read the replacement from a file; "-" reads stdin
  --batch FILE.jsonl One {"at", "xml"|"xmlFile"} per line, all locators
                     addressing the document AS READ; "-" = stdin
  --no-validate      Skip the full-document schema gate
  --author NAME      Author for the audit comment when document tracking is on
  -o, --output PATH  Write result to PATH instead of in-place
  --dry-run          Run every gate and preview, write nothing
  -v, --verbose      Full JSON ack instead of bare locators

Rules: a pN/tN takes one or MORE <w:p>/<w:tbl> roots; an sN takes exactly one
<w:sectPr>; an rIdN takes exactly one <Relationship> keeping the same Id.
Raw replaces are never tracked; --track is rejected.

Output:
  Prints the replacement's locator(s), one per line. Every gate must pass or
  NOTHING is written; errors print {code, error, hint?} + nonzero exit.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	xml: { type: "string" },
	"xml-file": { type: "string" },
	batch: { type: "string" },
	"no-validate": { type: "boolean" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	setVerboseAck(Boolean(values.verbose));

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	if (values.track) return failTrackRejected();
	if (values.batch !== undefined) {
		return runRawReplaceBatch(filePath, values.batch as string, values);
	}

	const locator = values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);
	const xml = await readXmlSource(values);
	if (typeof xml === "number") return xml;

	const partGuard = await rejectPartAsLocator(locator);
	if (partGuard !== undefined) return partGuard;
	if (isRelationshipLocator(locator)) {
		return runRelationshipReplace(filePath, locator, xml, values);
	}
	if (locator === "rels") {
		return fail(
			"USAGE",
			"--at rels is a listing, not a replaceable target",
			"Replace one relationship at a time: --at rIdN.",
		);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const reference = await resolveBlockOrFail(document, locator);
	if (typeof reference === "number") return reference;

	return replaceBlockWithXml({
		document,
		filePath,
		reference,
		locator,
		xml,
		values,
		operation: "raw.replace",
	});
}
