import {
	Document,
	describeForms,
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
	respondMinted,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(["span", "cellSpan"], "                    ");

const HELP = `docx hyperlinks add — wrap an existing span in a hyperlink

Usage:
  docx hyperlinks add FILE --at LOCATOR --url URL [options]

Required:
  --at LOCATOR      Span to wrap (within a single paragraph or cell). Supports:
${AT_FORMS}
                    Tip: don't hand-count offsets — run
                    \`docx find FILE "phrase"\` to get the exact span locator.
                    See \`docx info locators\`.
  --url URL         Target URL

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file (wins over -o)
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

The span must lie inside a single paragraph and must not overlap an existing
hyperlink, a tracked-change wrapper (<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>),
or any other run-bearing wrapper that we model. Resolve or accept the
wrapper first, then add the hyperlink.

When track-changes is on, an audit comment is anchored to the wrapped span
since OOXML has no native tracked-change form for hyperlink edits.

Output:
  Prints the new hyperlink id (e.g. linkN) on success — address it later with
  \`--at linkN\`. --verbose prints the full ack {ok:true, operation, path,
  hyperlinkId, text, at, url} — \`text\` echoes the span that was actually
  wrapped, so an off-by-one offset is obvious. A --dry-run prints a bare preview
  object (no id minted). Errors print {code, error, hint?} with a nonzero exit.
  Notation: uppercase letters are numeric indices; offsets are 0-based,
  end-exclusive.

Examples:
  docx hyperlinks add doc.docx --at p3:5-20 --url https://example.com
  docx hyperlinks add doc.docx --at t0:r1c2:p0:0-8 --url https://example.com
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

	// `Hyperlinks.add` mutates the XML tree in place but doesn't re-key
	// `hyperlinkById` (ids are positional, assigned by the reader's walk order).
	// Re-read the saved file and locate the hyperlink now covering the wrapped
	// span — that positional id is the addressable handle the agent can't
	// reconstruct, so print it by default; --verbose upgrades to the full ack.
	const savedPath = outputPath ?? path;
	const minted = await findMintedHyperlink(
		savedPath,
		target.blockId,
		target.span.start,
	);
	await respondMinted(minted ? [minted.id] : [], {
		ok: true,
		operation: "hyperlinks.add",
		path: savedPath,
		...(minted ? { hyperlinkId: minted.id, text: minted.text } : {}),
		at: atInput,
		url,
	});
	return EXIT.OK;
}

/** Re-read the saved document and find the linkN now covering `spanStart` in
 *  the target block, plus the text it actually wraps. `add` rejects spans
 *  overlapping an existing hyperlink, so the only hyperlink covering this offset
 *  after the add is the new one; the contiguous text runs carrying that id are
 *  the wrapped span — echoing them lets the caller catch an off-by-one offset.
 *  The text-offset model mirrors `paragraphText` (only `text` runs count).
 *  Returns `undefined` only if the re-read can't relocate the link (defensive —
 *  the span was just wrapped). */
async function findMintedHyperlink(
	savedPath: string,
	blockId: string,
	spanStart: number,
): Promise<{ id: string; text: string } | undefined> {
	const reread = await Document.open(savedPath);
	const block = reread.body.findBlockById(blockId);
	if (!block || block.type !== "paragraph") return undefined;

	let offset = 0;
	let foundId: string | undefined;
	let text = "";
	for (const run of block.runs) {
		if (run.type !== "text") continue;
		const runEnd = offset + run.text.length;
		// The first text run reaching past `spanStart` carries the new hyperlink.
		if (foundId === undefined && run.hyperlink && spanStart < runEnd) {
			foundId = run.hyperlink.id;
		}
		// Collect the contiguous run text carrying that id — the wrapped span.
		if (foundId !== undefined && run.hyperlink?.id === foundId) {
			text += run.text;
		}
		offset = runEnd;
	}
	return foundId === undefined ? undefined : { id: foundId, text };
}
