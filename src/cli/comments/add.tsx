import {
	type Locator,
	LocatorParseError,
	locatorToBlockTarget,
	parseLocator,
	saveDocView,
} from "@core";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	writeStdout,
} from "../respond";
import {
	addCommentMarkersToParagraph,
	addCommentRangeMarkers,
	authorInitials,
	CommentBody,
	type CommentSpan,
	ensureCommentsPart,
	generateParaId,
	nextCommentId,
	SpanOutOfRangeError,
} from "./helpers";

const HELP = `docx comments add — anchor a new comment to a locator

Usage:
  docx comments add FILE [options]

Required:
  --range LOCATOR   Where to anchor. Supports:
                      pN              whole paragraph
                      pN:S-E          chars S..E of pN
                      pN:S-pM:E       chars S of pN through char E of pM (cross-paragraph)
                      tT:rRcC:pK      whole cell paragraph
                      tT:rRcC:pK:S-E  chars S..E of cell paragraph
  --text TEXT       Comment body

Optional:
  --author NAME     Author name (default: $DOCX_AUTHOR)
  --dry-run         Print what would be added; do not write the file
  -h, --help        Show this help

Examples:
  docx comments add doc.docx --range p3 --text "Reconsider this paragraph"
  docx comments add doc.docx --range p3:5-20 --text "Sharper wording?" --author "Jane"
  docx comments add doc.docx --range p3:5-p5:10 --text "Whole section?" --author "Reviewer"
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				range: { type: "string" },
				text: { type: "string" },
				author: { type: "string" },
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

	const rangeInput = parsed.values.range as string | undefined;
	const text = parsed.values.text as string | undefined;
	if (!rangeInput) return fail("USAGE", "Missing --range LOCATOR", HELP);
	if (!text) return fail("USAGE", "Missing --text TEXT", HELP);

	let locator: Locator;
	try {
		locator = parseLocator(rangeInput);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	if (locator.kind === "range") {
		return runCrossBlock(parsed, path, rangeInput, locator, text);
	}

	const target = locatorToBlockTarget(locator);
	if (!target) {
		return fail(
			"INVALID_LOCATOR",
			"comments add supports paragraph locators only (pN, pN:start-end, pN:S-pM:E, or tT:rRcC:pK[:start-end])",
			"Comments and images are not valid anchors.",
		);
	}
	const blockId = target.blockId;

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const paragraphRef = await resolveBlockOrFail(view, blockId);
	if (typeof paragraphRef === "number") return paragraphRef;

	const span: CommentSpan | undefined = target.span;

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const date = new Date().toISOString();
	const numericId = nextCommentId(view);
	const paraId = generateParaId();

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.add",
			dryRun: true,
			path,
			commentId: `c${numericId}`,
			locator: rangeInput,
		});
		return EXIT.OK;
	}

	try {
		addCommentMarkersToParagraph(paragraphRef.node, numericId, span);
	} catch (error) {
		if (error instanceof SpanOutOfRangeError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	const commentsRoot = ensureCommentsPart(view);
	commentsRoot.children.push(
		<CommentBody
			options={{
				id: numericId,
				author,
				date,
				initials: authorInitials(author),
				paraId,
				text,
			}}
		/>,
	);

	await saveDocView(view);

	await respond({
		ok: true,
		operation: "comments.add",
		path,
		commentId: `c${numericId}`,
		locator: rangeInput,
	});
	return EXIT.OK;
}

async function runCrossBlock(
	parsed: ReturnType<typeof parseArgs>,
	path: string,
	rangeInput: string,
	locator: Extract<Locator, { kind: "range" }>,
	text: string,
): Promise<number> {
	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const startRef = await resolveBlockOrFail(view, locator.start.blockId);
	if (typeof startRef === "number") return startRef;
	const endRef = await resolveBlockOrFail(view, locator.end.blockId);
	if (typeof endRef === "number") return endRef;

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const date = new Date().toISOString();
	const numericId = nextCommentId(view);
	const paraId = generateParaId();

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.add",
			dryRun: true,
			path,
			commentId: `c${numericId}`,
			locator: rangeInput,
		});
		return EXIT.OK;
	}

	try {
		addCommentRangeMarkers(
			startRef.node,
			locator.start.offset,
			endRef.node,
			locator.end.offset,
			numericId,
		);
	} catch (error) {
		if (error instanceof SpanOutOfRangeError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	const commentsRoot = ensureCommentsPart(view);
	commentsRoot.children.push(
		<CommentBody
			options={{
				id: numericId,
				author,
				date,
				initials: authorInitials(author),
				paraId,
				text,
			}}
		/>,
	);

	await saveDocView(view);

	await respond({
		ok: true,
		operation: "comments.add",
		path,
		commentId: `c${numericId}`,
		locator: rangeInput,
	});
	return EXIT.OK;
}
