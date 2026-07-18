import { isRelationshipLocator } from "@core";
import { Raw } from "@core/raw";
import {
	fail,
	openOrFail,
	resolveBlockOrFail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
} from "../respond";
import {
	failTrackRejected,
	patchOrFail,
	replaceBlockWithXml,
	requireFindWith,
} from "./commit";
import { rejectPartAsLocator } from "./get";
import { failRelationshipNotFound, replaceRelationshipWithXml } from "./rels";

const HELP = `docx raw edit — patch a target's XML in place (find/replace, fully gated)

Usage:
  docx raw edit FILE --at LOCATOR --find STR --with STR

Examples:
  # Add line numbering to a section without shipping the whole sectPr:
  docx raw edit report.docx --at s0 --find '<w:cols' \\
      --with '<w:lnNumType w:countBy="1" w:restart="continuous"/><w:cols'
  # Tweak an attribute on a drop-cap paragraph:
  docx raw edit doc.docx --at p0 --find 'w:lines="2"' --with 'w:lines="3"'
  # Retarget a hyperlink relationship:
  docx raw edit doc.docx --at rId5 --find 'https://old' --with 'https://new'
  # Delete matched text (empty --with):
  docx raw edit doc.docx --at p3 --find '<w:proofErr w:type="spellStart"/>' --with ''

Options:
  --at LOCATOR       What to patch: pN | tN | sN | a cell-chained form
                     (t0:r0c0:p0) | rIdN (a relationship).
                     (Package parts are not locators — "docx raw part edit".)
  --find STR         Literal text to find in the target's exact XML — match
                     what "docx raw get --at LOCATOR" prints, byte for byte.
                     VERBATIM (no escape decoding), not a regex.
  --with STR         The replacement, VERBATIM. May be empty ('') to delete.
  --no-validate      Skip the schema gate on the result
  --author NAME      Author for the audit comment when document tracking is on
  -o, --output PATH  Write result to PATH instead of in-place
  --dry-run          Run every gate and preview, write nothing
  -v, --verbose      Full JSON ack instead of bare locators

This is the get → modify → replace patch loop in ONE call: the target's
current XML is read, every occurrence of --find becomes --with, and the
RESULT runs the same gates as raw replace (well-formedness, roots, child
order, references, schema diff). Raw edits are never tracked; --track is
rejected.

Output:
  Prints the patched target's locator(s) plus a "patched N occurrences" note.
  Zero matches is an error (exit 3) — nothing changed means nothing saved.
  Every gate must pass or NOTHING is written.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	find: { type: "string" },
	with: { type: "string" },
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
	const locator = values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);
	const findWith = await requireFindWith(values, HELP);
	if (typeof findWith === "number") return findWith;
	const { find, replaceWith } = findWith;
	const partGuard = await rejectPartAsLocator(locator);
	if (partGuard !== undefined) return partGuard;
	if (locator === "rels") {
		return fail(
			"USAGE",
			"--at rels is a listing — patch one relationship at a time",
			"Use --at rIdN.",
		);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const readBack = `docx raw get ${filePath} --at ${locator}`;

	if (isRelationshipLocator(locator)) {
		const current = new Raw(document).serializeRelationships(locator);
		if (current === undefined) {
			return failRelationshipNotFound(filePath, locator);
		}
		const patched = await patchOrFail(current, find, replaceWith, readBack);
		if (typeof patched === "number") return patched;
		return replaceRelationshipWithXml({
			document,
			filePath,
			rId: locator,
			xml: patched.xml,
			values,
			operation: "raw.edit",
			extraWarnings: [patched.note],
		});
	}

	const reference = await resolveBlockOrFail(document, locator);
	if (typeof reference === "number") return reference;
	const patched = await patchOrFail(
		new Raw(document).serializeBlock(reference),
		find,
		replaceWith,
		readBack,
	);
	if (typeof patched === "number") return patched;
	return replaceBlockWithXml({
		document,
		filePath,
		reference,
		locator,
		xml: patched.xml,
		values,
		operation: "raw.edit",
		extraWarnings: [patched.note],
	});
}
