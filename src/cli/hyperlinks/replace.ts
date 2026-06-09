import { describeForms } from "@core";
import { countHyperlinkUsages, Hyperlinks } from "@core/hyperlinks";
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

const AT_FORMS = describeForms(["hyperlink"], "                    ");

const HELP = `docx hyperlinks replace — change a hyperlink's URL

Usage:
  docx hyperlinks replace FILE --at linkN --with URL [options]

Required:
  --at linkN        Existing hyperlink to update. Supports:
${AT_FORMS}
                    See \`docx info locators\`.
  --with URL        New target URL

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file (wins over -o)
  -v, --verbose     Print the success ack JSON
  -h, --help        Show this help

Replaces only the targeted hyperlink. If multiple hyperlinks shared the same
underlying relationship, a new relationship is allocated so the others are
unaffected.

When track-changes is on, an audit comment is anchored to the affected
hyperlink span since OOXML has no native tracked-change form for hyperlink
target edits.

Output:
  Prints a one-line confirmation on success (exit 0) — replace mints no new id. --verbose prints
  {ok:true, operation, path, hyperlinkId, from, to}. A --dry-run prints a bare
  preview object. Errors print {code, error, hint?} with a nonzero exit.
  Discover ids with \`docx hyperlinks list FILE\`.

Examples:
  docx hyperlinks replace doc.docx --at link0 --with https://example.com
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			with: { type: "string" },
			author: { type: "string" },
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

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at linkN", HELP);

	const newUrl = parsed.values.with as string | undefined;
	if (!newUrl) return fail("USAGE", "Missing --with URL", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const reference = document.body.hyperlinkById.get(targetId);
	if (!reference) {
		return fail("HYPERLINK_NOT_FOUND", `Hyperlink not found: ${targetId}`);
	}

	const existingId = reference.relationshipId;
	const oldUrl = existingId
		? document.relationships.hyperlinksByRelationshipId.get(existingId)?.url
		: undefined;
	const sharedCount = existingId
		? countHyperlinkUsages(document.documentTree, existingId)
		: 0;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "hyperlinks.replace",
			dryRun: true,
			path,
			hyperlinkId: targetId,
			from: oldUrl,
			to: newUrl,
			sharedRelationship: sharedCount > 1,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const { from } = new Hyperlinks(document).replace(targetId, newUrl, {
		author: parsed.values.author as string | undefined,
	});

	await document.save(outputPath);

	// replace repoints an existing hyperlink and mints no new addressable
	// handle, so it stays silent on success unless --verbose.
	await respondAck({
		ok: true,
		operation: "hyperlinks.replace",
		path: outputPath ?? path,
		hyperlinkId: targetId,
		from,
		to: newUrl,
	});
	return EXIT.OK;
}
