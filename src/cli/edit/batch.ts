import {
	type BlockReference,
	CLEARABLE_ATTRS,
	type Document,
	Edit,
	EditError,
	LocatorParseError,
	LocatorResolveError,
	locatorToBlockTarget,
	MarkdownImport,
	MarkdownImportError,
	parseLocator,
	type Run,
	type RunFormat,
	resolveClearTags,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import type { XmlNode } from "@core/parser";
import {
	firstInvalidRunFormat,
	type RunFormatEnums,
} from "@core/run-formatting";
import {
	parseSpacingIndentFlags,
	parseTaskFlag,
	readJsonlObjects,
} from "../parse-helpers";
import {
	type ErrorCode,
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
} from "../respond";
import { normalizeHexColor } from "./index";
import { parseTabsValue, resolveTabsDirective } from "./tabs";

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** `docx edit --batch FILE.jsonl`: apply many edits from one read. Each JSONL
 *  line is `{ at, <one content field> }` — content is `text`, `clear`,
 *  `markdown`, `runs`, or `task`, OR a content-free run-formatting entry (set
 *  `bold`/`italic`/`underline`/`strike`/`color`/`highlight`/`shade`/`font`/`size`/
 *  `caps`/`smallcaps`/`superscript`/`subscript` on the existing text — the inverse
 *  of `clear`), plus per-entry `style`/`alignment`/`author`. All entries resolve
 *  against the document
 *  AS READ — we hold live node refs, so an edit to one paragraph never
 *  invalidates another's locator. Same-paragraph spans apply in descending
 *  offset order (so earlier offsets stay valid); a paragraph may take one
 *  whole-paragraph edit OR several non-overlapping spans, not both. Range
 *  (pN-pM), section (sN), and equation (eqN) edits are done one at a time. */
export async function runEditBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const conflicting = SINGLE_SHOT_FLAGS.find(
		(flag) => values[flag] !== undefined && values[flag] !== false,
	);
	if (conflicting) {
		return fail(
			"USAGE",
			`--batch reads each edit from the JSONL file; don't also pass --${conflicting} on the CLI`,
			"Put per-entry fields (at, text, clear, markdown, runs, code, task, style, …) on each JSONL line.",
		);
	}

	const authorFlag = values.author as string | undefined;
	const trackFlag = Boolean(values.track);
	const outputPath = values.output as string | undefined;
	const dryRun = Boolean(values["dry-run"]);
	const noFormatting = Boolean(values["no-formatting"]);

	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const track = resolveTracked(document, trackFlag);

	let resolved: ResolvedEntry[];
	try {
		resolved = [];
		for (let index = 0; index < rawEntries.length; index++) {
			const raw = rawEntries[index];
			if (raw === undefined) continue;
			resolved.push(
				await resolveEntry(document, raw, index, {
					authorFlag,
					noFormatting,
					track,
				}),
			);
		}
	} catch (error) {
		if (error instanceof EntryError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	let ordered: ResolvedEntry[];
	try {
		ordered = orderEntries(resolved);
	} catch (error) {
		if (error instanceof EntryError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	if (dryRun) {
		await respond({
			operation: "edit",
			dryRun: true,
			path: filePath,
			batch: resolved.map((entry) => ({ locator: entry.locatorString })),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	for (const entry of ordered) {
		try {
			entry.apply();
		} catch (error) {
			if (error instanceof EditError) {
				return fail(
					error.code,
					`entry ${entry.index} (${entry.locatorString}): ${error.message}`,
					error.hint,
				);
			}
			throw error;
		}
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "edit",
		path: outputPath ?? filePath,
		count: resolved.length,
		batch: resolved.map((entry) => ({ locator: entry.locatorString })),
	});
	return EXIT.OK;
}

/** Single-shot flags that have no meaning under `--batch` (each entry carries
 *  its own). Booleans default to `false` from parseArgs, so the caller checks
 *  `!== false` too. */
const SINGLE_SHOT_FLAGS = [
	"at",
	"text",
	"runs",
	"code",
	"code-file",
	"markdown",
	"markdown-file",
	"clear",
	"task",
	"equation",
	"display",
	"inline",
	"columns",
	"type",
	"style",
	"alignment",
	"tabs",
	"color",
	"bold",
	"italic",
	"underline",
	"strike",
	"caps",
	"smallcaps",
	"superscript",
	"subscript",
	"font",
	"size",
	"highlight",
	"shade",
	"space-before",
	"space-after",
	"line-spacing",
	"indent-left",
	"indent-right",
	"first-line",
	"hanging",
] as const;

// `clear` is NOT in here — it's a modifier that can stand alone OR ride along
// with one content key (fill text AND strip formatting in one entry, e.g.
// {at, text, clear:"highlight"} — the canonical form-fill + un-highlight move).
const CONTENT_KEYS = ["text", "markdown", "runs", "code", "task"] as const;

// Run-formatting keys (the inverse of `clear`): like `clear`, they can stand
// alone (set formatting on existing text) OR ride along with one content key
// (fill text AND format it). `color`/`bold`/`italic` are shared with the
// whole-paragraph `text` build; the rest are set-only.
const SET_KEYS = [
	"bold",
	"italic",
	"underline",
	"strike",
	"caps",
	"smallcaps",
	"superscript",
	"subscript",
	"color",
	"font",
	"size",
	"highlight",
	"shade",
] as const;

// Batch-entry keys that carry paragraph properties (no content of their own) —
// used to detect a props-only entry. Spacing/indent keys mirror the CLI flags.
const PARAGRAPH_PROP_KEYS = [
	"style",
	"alignment",
	"tabs",
	"space-before",
	"space-after",
	"line-spacing",
	"indent-left",
	"indent-right",
	"first-line",
	"hanging",
] as const;

type EntryOptions = {
	authorFlag?: string;
	noFormatting: boolean;
	track: boolean;
};

type ResolvedEntry = {
	index: number;
	locatorString: string;
	/** The target paragraph node — entries are grouped by it to detect
	 *  conflicting edits on one paragraph and to order same-paragraph spans. */
	node: XmlNode;
	span: { start: number; end: number } | null;
	/** Performs the mutation via the `Edit` lens. Throws `EditError`. */
	apply: () => void;
};

/** Validate one JSONL entry and produce its `apply` closure. Markdown content
 *  is pre-built here (async) so `apply` stays synchronous. Throws `EntryError`
 *  with the entry index for any malformed input. */
async function resolveEntry(
	document: Document,
	raw: Record<string, unknown>,
	index: number,
	opts: EntryOptions,
): Promise<ResolvedEntry> {
	const present = CONTENT_KEYS.filter((key) => raw[key] !== undefined);
	const hasClear = raw.clear !== undefined;
	const hasSet = SET_KEYS.some((key) => raw[key] !== undefined);
	const hasProps = PARAGRAPH_PROP_KEYS.some((key) => raw[key] !== undefined);
	if (present.length === 0 && !hasClear && !hasSet && !hasProps) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: no content — provide one of ${CONTENT_KEYS.join(", ")}, "clear", run-formatting (bold/color/font/…), or "style"/"alignment"/"tabs" to adjust in place`,
		);
	}
	if (present.length > 1) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: provide exactly one content field, got ${present.join(", ")}`,
		);
	}
	// A content-free set-formatting entry can't also carry clear/style/alignment —
	// those are separate operations (do them in separate entries).
	if (present.length === 0 && hasSet && (hasClear || hasProps)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: set-formatting (bold/color/font/…) can't combine with "clear"/"style"/"alignment" in one entry — use separate entries`,
		);
	}
	// A content key, "set" alone (format existing text), "clear" alone, a content
	// key + "set"/"clear" (fill then format/strip), or "style"/"alignment" alone.
	const kind = (present[0] ??
		(hasSet ? "set" : hasClear ? "clear" : "props")) as
		| (typeof CONTENT_KEYS)[number]
		| "set"
		| "clear"
		| "props";

	const at = raw.at;
	if (typeof at !== "string" || at.length === 0) {
		throw new EntryError("USAGE", `entry ${index}: "at" is required`);
	}
	if (/^s\d+$/.test(at)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: section edits (${at}) aren't supported in --batch`,
			"Edit sections individually with `docx edit --at sN`.",
		);
	}
	if (/^eq\d+$/.test(at)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: equation edits (${at}) aren't supported in --batch`,
			"Edit equations individually with `docx edit --at eqN`.",
		);
	}
	try {
		parseLocator(at);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			throw new EntryError(
				"INVALID_LOCATOR",
				`entry ${index}: ${error.message}`,
			);
		}
		throw error;
	}
	const target = locatorToBlockTarget(parseLocator(at));
	if (!target) {
		throw new EntryError(
			"INVALID_LOCATOR",
			`entry ${index}: "${at}" is not a paragraph, span, or cell-paragraph locator`,
			"Batch edits address pN, pN:S-E, or tN:rRcC:pK[:S-E]. Edit ranges (pN-pM), sections (sN), and equations (eqN) one at a time.",
		);
	}

	let blockRef: BlockReference;
	try {
		blockRef = document.body.resolveBlock(target.blockId);
	} catch (error) {
		if (error instanceof LocatorResolveError) {
			throw new EntryError(
				"BLOCK_NOT_FOUND",
				`entry ${index}: ${error.message}`,
			);
		}
		throw error;
	}

	const span = target.span ?? null;
	const apply = await buildApply(
		document,
		raw,
		index,
		kind,
		blockRef,
		span,
		opts,
	);
	return { index, locatorString: at, node: blockRef.node, span, apply };
}

/** Build the mutation closure for an entry once its locator is resolved. Spans
 *  accept only `text`/`clear`; whole paragraphs accept every content kind. */
async function buildApply(
	document: Document,
	raw: Record<string, unknown>,
	index: number,
	kind: (typeof CONTENT_KEYS)[number] | "set" | "clear" | "props",
	blockRef: BlockReference,
	span: { start: number; end: number } | null,
	opts: EntryOptions,
): Promise<() => void> {
	const author = typeof raw.author === "string" ? raw.author : opts.authorFlag;
	const clearTags =
		raw.clear !== undefined ? resolveClearOrThrow(raw.clear, index) : null;

	// Set-formatting entry (no content key): SET the named formatting on the
	// existing text — a span or whole paragraph (the inverse of "clear").
	if (kind === "set") {
		const format = readRunFormat(raw, index);
		if (!format) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: no run-formatting fields to apply`,
			);
		}
		return () => new Edit(document).setFormatting(blockRef, span, format);
	}

	// Style-only entry (no content key): restyle the paragraph in place, keeping
	// its runs. Mirrors single-shot `edit --at pN --style X`.
	if (kind === "props") {
		if (span) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: a character span (${raw.at}) can't take style/alignment/spacing/indent/tabs — restyle the whole paragraph (pN).`,
			);
		}
		const paragraphOptions = readParagraphOptions(document, raw, index);
		return () =>
			void new Edit(document).paragraphProperties(blockRef, paragraphOptions, {
				authorFlag: author,
				track: opts.track,
			});
	}

	// Clear-only entry (no content key): strip formatting in place.
	if (kind === "clear") {
		// clearTags is non-null here (resolveEntry only sets kind="clear" when
		// raw.clear is present), but assert for the type checker.
		const tags = clearTags ?? new Set<string>();
		return () => new Edit(document).clearFormatting(blockRef, span, tags);
	}

	// A content kind, optionally combined with `clear` (apply content, then strip
	// — the canonical "fill this placeholder AND remove its highlight" move).
	if (span) {
		if (kind === "text") {
			const text = requireString(raw.text, index, "text");
			rejectSpanParagraphFlags(raw, index);
			// On a span, `edit.span` only replaces text (it ignores run format), so
			// EVERY formatting key — including color/bold/italic — rides along.
			const ride = readRunFormatRideAlong(raw, index, false);
			if (clearTags || ride) {
				// `find` emits span locators, so {at:span, text, clear/bold/…} is the
				// natural one-shot: replace the span, then clear/set formatting on the
				// just-written range (offsets shift to the new length).
				return () => {
					const edit = new Edit(document);
					edit.span(blockRef, span, text, {
						authorFlag: author,
						track: opts.track,
					});
					const replaced = { start: span.start, end: span.start + text.length };
					if (clearTags) {
						edit.clearFormattingNode(blockRef.node, replaced, clearTags);
					}
					if (ride) edit.setFormattingNode(blockRef.node, replaced, ride);
				};
			}
			return () =>
				new Edit(document).span(blockRef, span, text, {
					authorFlag: author,
					track: opts.track,
				});
		}
		throw new EntryError(
			"USAGE",
			`entry ${index}: a character span (${raw.at}) supports "text", "clear", or run-formatting (bold/color/font/…)`,
			"Use a whole-paragraph locator (pN) for markdown/runs/code/task.",
		);
	}

	// Whole-paragraph content → a closure returning the resulting paragraph node
	// (so a trailing clear/set-format targets the post-edit node, not the
	// replaced one). For `text`, color/bold/italic land on the new runs via the
	// content build, so they're excluded from the ride-along (the rest still ride).
	const contentNode = await buildWholeParagraphContent(
		document,
		raw,
		index,
		kind,
		blockRef,
		author,
		opts,
	);
	const ride = readRunFormatRideAlong(raw, index, kind === "text");
	if (!clearTags && !ride) return () => void contentNode();
	return () => {
		const node = contentNode();
		const edit = new Edit(document);
		if (clearTags) edit.clearFormattingNode(node, null, clearTags);
		if (ride) edit.setFormattingNode(node, null, ride);
	};
}

/** Build a closure that applies one whole-paragraph content edit and returns the
 *  resulting paragraph node. Split out so the combined content+clear path can
 *  clear that node afterward. */
async function buildWholeParagraphContent(
	document: Document,
	raw: Record<string, unknown>,
	index: number,
	kind: (typeof CONTENT_KEYS)[number],
	blockRef: BlockReference,
	author: string | undefined,
	opts: EntryOptions,
): Promise<() => XmlNode> {
	const paragraphOptions = readParagraphOptions(document, raw, index);
	if (kind === "text") {
		const text = requireString(raw.text, index, "text");
		// Empty text leaves a blank space-consuming paragraph, not a removed line —
		// redirect to the delete batch (the honest verb for removals).
		if (text === "") {
			throw new EntryError(
				"USAGE",
				`entry ${index}: empty "text" leaves a blank paragraph in place, it doesn't remove the line (${raw.at}).`,
				'Remove lines with `docx delete --batch` ({ "at": "pN" } per line). To keep an empty spacer, use "runs": [] instead.',
			);
		}
		const format = readTextFormat(raw, index);
		return () =>
			new Edit(document).paragraph(
				blockRef,
				{ kind: "text", text, format, paragraphOptions },
				{
					authorFlag: author,
					noFormatting: opts.noFormatting,
					track: opts.track,
				},
			);
	}
	if (kind === "task") {
		const taskRaw = requireString(raw.task, index, "task");
		const checked = parseTaskFlag(taskRaw);
		if (checked === null) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "task" must be "checked" or "unchecked", got "${taskRaw}"`,
			);
		}
		return () => {
			new Edit(document).taskToggle(blockRef, checked, {
				authorFlag: author,
				track: opts.track,
			});
			return blockRef.node; // toggled in place
		};
	}
	if (kind === "runs") {
		const runs = readRuns(raw.runs, index);
		return () =>
			new Edit(document).paragraph(
				blockRef,
				{ kind: "runs", runs, paragraphOptions },
				{ authorFlag: author, track: opts.track },
			);
	}
	if (kind === "code") {
		const content = requireString(raw.code, index, "code");
		const language =
			typeof raw.language === "string" ? raw.language : undefined;
		return () =>
			new Edit(document).paragraph(
				blockRef,
				{
					kind: "code",
					content,
					...(language ? { language } : {}),
					paragraphOptions,
				},
				{ authorFlag: author, track: opts.track },
			);
	}
	// kind === "markdown" — pre-build blocks now so apply stays synchronous.
	const source = requireString(raw.markdown, index, "markdown");
	let blocks: XmlNode[];
	try {
		blocks = await new MarkdownImport(document).blocks(source);
	} catch (error) {
		if (error instanceof MarkdownImportError) {
			throw new EntryError(
				error.code,
				`entry ${index}: ${error.message}`,
				error.hint,
			);
		}
		throw error;
	}
	return () =>
		new Edit(document).paragraph(
			blockRef,
			{ kind: "markdown-blocks", blocks, paragraphOptions },
			{ authorFlag: author, track: opts.track },
		);
}

/** Group entries by their target paragraph and resolve same-paragraph
 *  conflicts: a paragraph takes one whole-paragraph edit OR several
 *  non-overlapping spans applied in descending offset order. Cross-paragraph
 *  ordering is irrelevant — we hold live node refs, so editing one paragraph
 *  never shifts another's. Throws `EntryError` on a conflict. */
function orderEntries(resolved: ResolvedEntry[]): ResolvedEntry[] {
	const groups = new Map<XmlNode, ResolvedEntry[]>();
	const order: XmlNode[] = [];
	for (const entry of resolved) {
		let group = groups.get(entry.node);
		if (!group) {
			group = [];
			groups.set(entry.node, group);
			order.push(entry.node);
		}
		group.push(entry);
	}

	const out: ResolvedEntry[] = [];
	for (const node of order) {
		const group = groups.get(node);
		if (!group) continue;
		if (group.length === 1) {
			const only = group[0];
			if (only) out.push(only);
			continue;
		}
		const indices = group.map((entry) => entry.index).join(", ");
		const whole = group.find((entry) => entry.span === null);
		if (whole) {
			throw new EntryError(
				"USAGE",
				`entries ${indices} target the same paragraph (${whole.locatorString})`,
				'One whole-paragraph edit per entry. To fill text AND strip formatting on the same paragraph, combine them in ONE entry: {"at":"…","text":"…","clear":"highlight"}.',
			);
		}
		// Descending start so an earlier-offset edit doesn't shift a later one.
		const sorted = [...group].sort(
			(left, right) => (right.span?.start ?? 0) - (left.span?.start ?? 0),
		);
		for (let position = 0; position < sorted.length - 1; position++) {
			const current = sorted[position]?.span;
			const next = sorted[position + 1]?.span;
			if (current && next && next.end > current.start) {
				throw new EntryError(
					"USAGE",
					`entries ${indices} have overlapping spans on the same paragraph`,
					"Coalesce them into one span edit.",
				);
			}
		}
		out.push(...sorted);
	}
	return out;
}

function requireString(value: unknown, index: number, field: string): string {
	if (typeof value !== "string") {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "${field}" must be a string`,
		);
	}
	return value;
}

function rejectSpanParagraphFlags(
	raw: Record<string, unknown>,
	index: number,
): void {
	if (PARAGRAPH_PROP_KEYS.some((key) => raw[key] !== undefined)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: style/alignment/spacing/indent/tabs apply to a whole paragraph, not a character span`,
		);
	}
}

/** Build a `RunFormat` from an entry's set-formatting keys (the inverse of
 *  `clear`). Throws `EntryError` on a bad value. Returns null if none are set.
 *  `underline` accepts a style string (e.g. "double") or a boolean (→ "single"). */
function readRunFormat(
	raw: Record<string, unknown>,
	index: number,
): RunFormat | null {
	const format: RunFormat = {};
	// Boolean toggles only turn a property ON (matching single-shot parseRunFormat);
	// a falsy value (false/0/null) is ignored — turning a property OFF is `clear`'s
	// job, not set's.
	if (raw.bold) format.bold = true;
	if (raw.italic) format.italic = true;
	if (raw.strike) format.strike = true;
	if (raw.caps) format.allCaps = true;
	if (raw.smallcaps) format.smallCaps = true;
	// `underline` accepts a style string (e.g. "double") or `true` (→ "single");
	// a falsy non-string is ignored (so `underline:false` doesn't turn it ON).
	if (typeof raw.underline === "string") format.underline = raw.underline;
	else if (raw.underline === true) format.underline = "single";
	if (raw.underlineColor !== undefined) {
		format.underlineColor = normalizeHexColor(
			requireString(raw.underlineColor, index, "underlineColor"),
		);
	}
	if (raw.color !== undefined)
		format.color = normalizeHexColor(requireString(raw.color, index, "color"));
	if (raw.font !== undefined)
		format.font = requireString(raw.font, index, "font");
	if (raw.highlight !== undefined) {
		format.highlight = requireString(raw.highlight, index, "highlight");
	}
	if (raw.shade !== undefined)
		format.shade = normalizeHexColor(requireString(raw.shade, index, "shade"));
	if (raw.superscript && raw.subscript) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "superscript" and "subscript" are mutually exclusive`,
		);
	}
	if (raw.superscript) format.vertAlign = "superscript";
	if (raw.subscript) format.vertAlign = "subscript";
	if (raw.size !== undefined) {
		const points =
			typeof raw.size === "number"
				? raw.size
				: Number.parseFloat(String(raw.size));
		if (!Number.isFinite(points) || points <= 0) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "size" must be a positive point size (e.g. 12 or 11.5)`,
			);
		}
		format.sizeHalfPoints = Math.round(points * 2);
	}
	const invalid = firstInvalidRunFormat(format);
	if (invalid) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: invalid ${invalid.field} "${invalid.value}"`,
			`Use ${invalid.valid}.`,
		);
	}
	if (Object.keys(format).length === 0) return null;
	return format;
}

/** The set-formatting that rides along with a content edit. `excludeBasic` drops
 *  color/bold/italic (the whole-paragraph `text` build already applies them to
 *  the new runs); on a span every key rides since `edit.span` ignores format. */
function readRunFormatRideAlong(
	raw: Record<string, unknown>,
	index: number,
	excludeBasic: boolean,
): RunFormat | null {
	const format = readRunFormat(raw, index);
	if (!format) return null;
	if (excludeBasic) {
		format.color = undefined;
		format.bold = undefined;
		format.italic = undefined;
	}
	if (!Object.values(format).some((value) => value !== undefined)) return null;
	return format;
}

function readTextFormat(
	raw: Record<string, unknown>,
	index: number,
): { color?: string; bold?: boolean; italic?: boolean } {
	const out: { color?: string; bold?: boolean; italic?: boolean } = {};
	if (raw.color !== undefined) {
		if (typeof raw.color !== "string") {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "color" must be a hex string`,
			);
		}
		out.color = raw.color;
	}
	if (raw.bold !== undefined) out.bold = Boolean(raw.bold);
	if (raw.italic !== undefined) out.italic = Boolean(raw.italic);
	return out;
}

function readParagraphOptions(
	document: Document,
	raw: Record<string, unknown>,
	index: number,
): ParagraphOptions {
	const out: ParagraphOptions = {};
	if (raw.style !== undefined) {
		if (typeof raw.style !== "string") {
			throw new EntryError("USAGE", `entry ${index}: "style" must be a string`);
		}
		out.style = raw.style;
	}
	if (raw.alignment !== undefined) {
		const alignment = raw.alignment;
		if (
			alignment !== "left" &&
			alignment !== "center" &&
			alignment !== "right" &&
			alignment !== "justify"
		) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: invalid "alignment" — use left, center, right, or justify`,
			);
		}
		out.alignment = alignment;
	}
	if (raw.tabs !== undefined) {
		if (typeof raw.tabs !== "string") {
			throw new EntryError("USAGE", `entry ${index}: "tabs" must be a string`);
		}
		const parsed = parseTabsValue(raw.tabs);
		if ("error" in parsed) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: ${parsed.error}`,
				parsed.hint,
			);
		}
		out.tabs = resolveTabsDirective(parsed, document);
	}
	// Spacing/indent keys mirror the CLI flags (space-before, line-spacing,
	// indent-left, …); the shared parser handles numbers or strings.
	const spacingIndent = parseSpacingIndentFlags(raw);
	if ("error" in spacingIndent) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: ${spacingIndent.error}`,
			spacingIndent.hint,
		);
	}
	if (spacingIndent.spacing) out.spacing = spacingIndent.spacing;
	if (spacingIndent.indent) out.indent = spacingIndent.indent;
	return out;
}

function readRuns(value: unknown, index: number): Run[] {
	if (!Array.isArray(value)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "runs" must be a JSON array of Run objects`,
		);
	}
	for (const run of value) {
		if (
			run !== null &&
			typeof run === "object" &&
			(run as { type?: unknown }).type === "text"
		) {
			const invalid = firstInvalidRunFormat(run as RunFormatEnums);
			if (invalid) {
				throw new EntryError(
					"USAGE",
					`entry ${index}: invalid ${invalid.field} "${invalid.value}" in a run`,
					`Use ${invalid.valid}.`,
				);
			}
		}
	}
	return value as Run[];
}

function resolveClearOrThrow(value: unknown, index: number): Set<string> {
	// Accept "highlight", "highlight,underline", or ["highlight","underline"].
	const raw =
		typeof value === "string"
			? [value]
			: Array.isArray(value) && value.every((v) => typeof v === "string")
				? (value as string[])
				: null;
	if (!raw) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "clear" must be an attribute name, a comma list, or an array of names (or "all")`,
		);
	}
	const names = raw
		.flatMap((entry) => entry.split(","))
		.map((name) => name.trim().toLowerCase())
		.filter(Boolean);
	if (names.length === 0) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "clear" needs an attribute name, or "all"`,
		);
	}
	const tags = resolveClearTags(names);
	if (!tags) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: unknown attribute in "${raw.join(",")}". Valid: ${CLEARABLE_ATTRS.join(", ")}, all`,
		);
	}
	return tags;
}

/** Per-entry validation failure. `code` is a CLI `ErrorCode` so the caller can
 *  `fail(err.code, …)` directly, matching the `comments add --batch` pattern. */
class EntryError extends Error {
	constructor(
		public code: ErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "EntryError";
	}
}
