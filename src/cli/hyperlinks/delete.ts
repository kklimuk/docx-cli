import { HyperlinkStaleError, Hyperlinks } from "@core/hyperlinks";
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

const HELP = `docx hyperlinks delete — unwrap a hyperlink (keep the text)

Usage:
  docx hyperlinks delete FILE --at LINK_ID [options]

Required:
  --at LINK_ID      Existing hyperlink to remove (e.g., link0)

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

The display text stays in place; only the <w:hyperlink> wrapper is removed.
If the underlying relationship is no longer referenced, it is pruned from the
rels file too.

When track-changes is on, an audit comment is anchored to the surviving text
since OOXML has no native tracked-change form for hyperlink removal.

Examples:
  docx hyperlinks delete doc.docx --at link0
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			author: { type: "string" },
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

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at LINK_ID", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const reference = document.body.hyperlinkById.get(targetId);
	if (!reference) {
		return fail("HYPERLINK_NOT_FOUND", `Hyperlink not found: ${targetId}`);
	}

	const oldUrl = reference.relationshipId
		? document.relationships.hyperlinksByRelationshipId.get(
				reference.relationshipId,
			)?.url
		: undefined;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "hyperlinks.delete",
			dryRun: true,
			path,
			hyperlinkId: targetId,
			from: oldUrl,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	try {
		new Hyperlinks(document).delete(targetId, {
			author: parsed.values.author as string | undefined,
		});
	} catch (error) {
		if (error instanceof HyperlinkStaleError) {
			return fail("HYPERLINK_NOT_FOUND", error.message);
		}
		throw error;
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "hyperlinks.delete",
		path: outputPath ?? path,
		hyperlinkId: targetId,
		from: oldUrl,
	});
	return EXIT.OK;
}
