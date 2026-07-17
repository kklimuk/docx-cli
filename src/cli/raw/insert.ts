import { isCellScopedLocator } from "@core";
import { isRelationshipFragment, Raw } from "@core/raw";
import { parseTargetPlacement, resolvePlacement } from "../insert/place";
import {
	fail,
	openOrFail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
} from "../respond";
import { runRawInsertBatch } from "./batch";
import {
	commitRawMutation,
	failTrackRejected,
	noteUntrackedRawChange,
	rawOrFail,
	respondDryRun,
} from "./commit";
import { runRelationshipInsert } from "./rels";
import { readXmlSource } from "./xml-source";

const HELP = `docx raw insert — insert raw OOXML block(s) or a relationship

Usage:
  docx raw insert FILE (--xml STR | --xml-file PATH|-)
                       (--after LOCATOR | --before LOCATOR | --at-start | --at-end)
  docx raw insert FILE --xml '<Relationship Type="…" Target="…" TargetMode="External"/>'
  docx raw insert FILE --batch FILE.jsonl

Options:
  --xml STR          The fragment, inline and VERBATIM (no escape decoding —
                     XML has its own escapes: &#10; &lt; &amp;)
  --xml-file PATH    Read the fragment from a file; "-" reads stdin
  --after LOCATOR    Insert after this block (pN, tN, tN:rRcC:pK, …)
  --before LOCATOR   Insert before this block
  --at-start         Insert before the first content block
  --at-end           Insert after the last content block
  --batch FILE.jsonl One {"after"|"before", "xml"|"xmlFile"} per line, all
                     locators addressing the document AS READ; "-" = stdin
  --no-validate      Skip the full-document schema gate (the other gates and
                     the nothing-written-on-failure guarantee still apply)
  --author NAME      Author for the audit comment when document tracking is on
  -o, --output PATH  Write result to PATH instead of in-place
  --dry-run          Run every gate and preview, write nothing
  -v, --verbose      Full JSON ack instead of bare locators

Fragment roots must be <w:p> or <w:tbl> — the elements that get locators.
Anything INSIDE them is unrestricted (drop caps, fields, content controls,
embedded objects, …). Prints each inserted block's locator. Raw inserts are
never tracked; --track is rejected.

A fragment whose roots are <Relationship Id? Type Target [TargetMode]/> goes
to word/_rels/document.xml.rels instead (no placement flags — the rels part
is an unordered set). Omit Id to have a free rIdN minted; the printed id is
what body XML references (r:id/r:embed). An internal Target must name an
existing part; URLs need TargetMode="External".

Example (drop cap):
  docx raw insert report.docx --after p0 --xml '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/></w:pPr><w:r><w:t>D</w:t></w:r></w:p>'
`;

const OPTION_SPEC = {
	xml: { type: "string" },
	"xml-file": { type: "string" },
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
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
		return runRawInsertBatch(filePath, values.batch as string, values);
	}

	const xml = await readXmlSource(values);
	if (typeof xml === "number") return xml;
	if (isRelationshipFragment(xml)) {
		return runRelationshipInsert(filePath, xml, values);
	}
	const placement = await parseTargetPlacement(values, HELP);
	if (typeof placement === "number") return placement;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const raw = new Raw(document);
	const prepared = await rawOrFail(() => raw.prepareFragment(xml, "blocks"));
	if (typeof prepared === "number") return prepared;

	const resolved = await resolvePlacement(document, placement);
	if (typeof resolved === "number") return resolved;
	const { blockRef, mode, locator } = resolved;

	if (values["dry-run"]) {
		return respondDryRun(
			"raw.insert",
			filePath,
			{
				anchor: locator,
				placement: mode,
				roots: prepared.nodes.map((node) => node.tag),
			},
			prepared.warnings,
			values.output as string | undefined,
		);
	}

	const spliced = await rawOrFail(() =>
		raw.spliceBlocks(blockRef, mode, prepared, {
			cellScoped: isCellScopedLocator(locator),
		}),
	);
	if (typeof spliced === "number") return spliced;

	const warnings = [...prepared.warnings, ...spliced];
	noteUntrackedRawChange(
		document,
		prepared.nodes,
		"raw.insert",
		values.author as string | undefined,
		warnings,
	);

	return commitRawMutation({
		document,
		filePath,
		operation: "raw.insert",
		insertedNodes: prepared.nodes,
		validate: !values["no-validate"],
		warnings,
		outputPath: values.output as string | undefined,
	});
}
