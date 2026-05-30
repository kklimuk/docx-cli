import {
	type Locator,
	LocatorParseError,
	locatorToBlockTarget,
	parseLocator,
} from "@core";
import { Hyperlinks, HyperlinkWrapError } from "@core/hyperlinks";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx hyperlinks add — wrap an existing span in a hyperlink

Usage:
  docx hyperlinks add FILE --at LOCATOR --url URL [options]

Required:
  --at LOCATOR      Where to wrap. Supports:
                      pN:S-E          chars S..E of pN
                      tT:rRcC:pK:S-E  chars S..E of a cell paragraph
  --url URL         Target URL

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

The span must lie inside a single paragraph and must not overlap an existing
hyperlink, a tracked-change wrapper (<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>),
or any other run-bearing wrapper that we model. Resolve or accept the
wrapper first, then add the hyperlink.

When track-changes is on, an audit comment is anchored to the wrapped span
since OOXML has no native tracked-change form for hyperlink edits.

Examples:
  docx hyperlinks add doc.docx --at p3:5-20 --url https://example.com
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			url: { type: "string" },
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

	const atInput = parsed.values.at as string | undefined;
	const url = parsed.values.url as string | undefined;
	if (!atInput) return fail("USAGE", "Missing --at LOCATOR", HELP);
	if (!url) return fail("USAGE", "Missing --url URL", HELP);

	let locator: Locator;
	try {
		locator = parseLocator(atInput);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	const target = locatorToBlockTarget(locator);
	if (!target?.span) {
		return fail(
			"INVALID_LOCATOR",
			"hyperlinks add requires a span locator like pN:S-E or tT:rRcC:pK:S-E",
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const paragraphRef = await resolveBlockOrFail(document, target.blockId);
	if (typeof paragraphRef === "number") return paragraphRef;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "hyperlinks.add",
			dryRun: true,
			path,
			at: atInput,
			url,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	try {
		new Hyperlinks(document).add(paragraphRef.node, target.span, url, {
			author: parsed.values.author as string | undefined,
		});
	} catch (error) {
		if (error instanceof HyperlinkWrapError) {
			return fail("USAGE", error.message);
		}
		throw error;
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "hyperlinks.add",
		path: outputPath ?? path,
		at: atInput,
		url,
	});
	return EXIT.OK;
}
