import {
	type Locator,
	LocatorParseError,
	LocatorResolveError,
	openDocView,
	PkgError,
	parseLocator,
	resolveBlock,
	saveDocView,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";
import {
	addCommentMarkersToParagraph,
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
  --range LOCATOR   Where to anchor (e.g., p3 for the whole paragraph,
                    p3:5-20 to comment on chars 5..20 of p3)
  --text TEXT       Comment body

Optional:
  --author NAME     Author name (default: $DOCX_AUTHOR)
  --dry-run         Print what would be added; do not write the file
  -h, --help        Show this help

Examples:
  docx comments add doc.docx --range p3 --text "Reconsider this paragraph"
  docx comments add doc.docx --range p3:5-20 --text "Sharper wording?" --author "Jane"
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

	const blockId =
		locator.kind === "block"
			? locator.blockId
			: locator.kind === "blockSpan"
				? locator.blockId
				: null;
	if (!blockId?.startsWith("p")) {
		return fail(
			"INVALID_LOCATOR",
			"comments add supports paragraph locators only (pN or pN:start-end)",
			"Cross-block ranges and table-cell anchors are not yet supported.",
		);
	}

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (openError) {
		if (openError instanceof PkgError) {
			if (openError.code === "FILE_NOT_FOUND") {
				return fail("FILE_NOT_FOUND", openError.message);
			}
			if (openError.code === "NOT_A_ZIP") {
				return fail("NOT_A_ZIP", openError.message);
			}
		}
		throw openError;
	}

	let paragraphRef: ReturnType<typeof resolveBlock>;
	try {
		paragraphRef = resolveBlock(view, blockId);
	} catch (error) {
		if (error instanceof LocatorResolveError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		throw error;
	}

	const span: CommentSpan | undefined =
		locator.kind === "blockSpan"
			? { start: locator.start, end: locator.end }
			: undefined;

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
