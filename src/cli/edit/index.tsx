import {
	type BlockRangeReference,
	type BlockReference,
	CLEARABLE_ATTRS,
	type Document,
	describeForms,
	Edit,
	EditError,
	type Locator,
	locatorToBlockTarget,
	MarkdownImport,
	MarkdownImportError,
	type ParagraphContentSpec,
	parseLocator,
	type Run,
	resolveClearTags,
	type SectionType,
	type XmlNode,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	EquationNotFoundError,
	EquationParseError,
	EquationStaleError,
	Equations,
} from "@core/equation";
import type { parseArgs } from "util";
import {
	parseRunsArg,
	parseSectionFlags,
	parseTaskFlag,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveBlockRangeOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runEditBatch } from "./batch";

const AT_FORMS = describeForms(
	[
		"paragraph",
		"span",
		"blockRange",
		"cellParagraph",
		"cellSpan",
		"section",
		"equation",
	],
	"                      ",
);

const HELP = `docx edit — replace a paragraph (or paragraph range), a section, or an equation

Usage:
  docx edit FILE --at LOCATOR <content> [options]
  docx edit FILE --batch FILE.jsonl [options]   # many edits, one read
  docx edit FILE --batch -          [options]   # read JSONL from stdin

Locator (required):
  --at LOCATOR      What to edit. One of:
${AT_FORMS}
                    pN / pN-pM take paragraph content (below); sN takes
                    --columns/--type; eqN takes --equation/--display/--inline.
                    A character span (pN:S-E, or a cell paragraph
                    tN:rRcC:pK:S-E) replaces just those characters with --text,
                    inheriting the existing run's formatting — paste a locator
                    straight from \`docx find\`. See \`docx info locators\`.

Paragraph content (one required for paragraph / range locators):
  --text TEXT       Replace with a single-run paragraph
  --runs JSON       Replace with custom runs (Run[] JSON)
  --code TEXT       Replace with a code block — newlines split into one
                    CodeBlock-styled paragraph per source line
  --code-file PATH  Same as --code, but read content from PATH (use "-" for stdin)
  --markdown TEXT   Replace with parsed GFM markdown. Same dialect as
                    'docx insert --markdown' (headings, lists, tables, code,
                    blockquotes, links, math, footnotes, CriticMarkup, …).
                    Multi-block sources expand naturally — a paragraph
                    locator gets replaced by however many blocks the source
                    parses to.
  --markdown-file PATH  Same as --markdown, but read content from PATH
                    (use "-" for stdin).
  --clear ATTRS     Strip run formatting in place, keeping the text. ATTRS is a
                    comma list of: bold, italic, strike, underline, highlight,
                    shade, color, font, size, vertalign, caps, smallcaps, style
                    — or "all". Repeatable: \`--clear highlight --clear underline\`
                    accumulates (same as \`--clear highlight,underline\`).
                    Works on a whole paragraph (pN) or a span
                    (pN:S-E). Pairs with \`docx find --highlight\`. (Not tracked.)
                    Can RIDE ALONG with content (whole paragraph OR span):
                    \`--text "Delaware" --clear highlight\` fills then strips the
                    highlight in ONE call — the form-fill + un-highlight move,
                    and the natural \`find --highlight | edit\` one-shot. Prefer
                    this over \`--clear all\`, which also drops the font size.
  --task STATE      Flip a task-list item's checkbox state in place ("checked" or
                    "unchecked"). Requires a single paragraph locator that is
                    already a GFM task list item (has a leading <w14:checkbox/>
                    SDT). Under track-changes, emits Word's canonical toggle
                    shape (ins/del pair inside sdtContent + w14:checked flip);
                    "track-changes list" surfaces it as a checkboxToggle revision.

Equation editing (requires --at eqN):
  --equation LATEX  Replace the equation's content with new LaTeX. Goes through
                    temml → MathML → OMML.
  --display         Switch the equation to display mode (block, $$…$$). Can be
                    combined with --equation, or used alone to toggle mode.
  --inline          Switch to inline mode ($…$). Mutex with --display.

Paragraph options:
  --style NAME       Paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic

Code options (only with --code / --code-file):
  --language LANG   Syntax-highlight via lowlight (37 common languages bundled).
                    Survives round-trip via a CodeBlock-LANG pStyle suffix.

Section options (for section locators sN):
  --columns N        Number of columns for the targeted section
  --type T           continuous | nextPage | evenPage | oddPage | nextColumn

Formatting (single-paragraph --text only):
  By default --text preserves run-level formatting (bold/italic/color/etc.)
  on words shared between the old and new text via a word-level diff. New
  words inherit formatting from the nearest unchanged neighbor. Pass
  --no-formatting to fall back to a single fresh run with no formatting.
  Passing --color/--bold/--italic also bypasses preservation — those flags
  apply uniformly to the new paragraph. Range edits (pN-pM) always rewrite
  the span wholesale; per-word formatting preservation is not applied.

Batch (--batch PATH | -):
  Apply many edits from one read — no need to re-read between changes. Each
  JSONL line is one edit: { "at": LOCATOR, <one content field> }. Content is
  "text", "markdown", "runs", "code", or "task"; whole-paragraph entries may
  also carry "style"/"alignment"/"color"/"bold"/"italic"/"author". An entry may
  ALSO carry "clear" alongside its content — whole paragraph OR span — to fill
  then strip formatting in one entry ({ "at": p, "text": "...", "clear": "highlight" })
  — the form-fill + un-highlight move; "clear" alone (no content) is also valid.
  All locators address the document AS READ. A paragraph takes one whole-
  paragraph edit OR several non-overlapping spans (applied right-to-left so
  offsets stay valid). Range (pN-pM), section (sN), and equation (eqN) edits
  run one at a time, not in a batch. Don't mix --batch with --at/--text/etc.
  Tip: a value reaches docx-cli verbatim through the JSONL file — no shell in the
  way — so prefer --batch for money ($1,250.00), regex, or other shell-special
  text. (A bare --text value may also start with "-"; it no longer needs "=".)

General options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record this edit as a tracked change even when the
                    document's track-changes toggle is off (it is OFF by
                    default — check with \`docx track-changes list FILE\`).
  --no-formatting   Replace with a single fresh run; do not preserve rPr
                    on unchanged words
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0) — the edited locator is unchanged, so there's
  nothing to mint. --verbose prints {ok:true, operation, path, locator}.
  Errors print {code, error, hint?} with a nonzero exit. Discover ids with
  \`docx read FILE --ast\` (equation ids appear on EquationRun nodes).

Examples:
  docx find doc.docx "fill in state"            # → p4:25-38
  docx edit doc.docx --at p4:25-38 --text "Delaware"   # replace just that span
  docx edit doc.docx --at p4:25-38 --clear highlight   # un-highlight that span
  docx edit doc.docx --at p4 --text "Delaware" --clear highlight  # fill + un-highlight
  docx edit doc.docx --at p2 --clear all               # strip all run formatting
  docx edit doc.docx --at p3 --text "Replaced." --style Heading2
  docx edit doc.docx --at p0 --runs '[{"type":"text","text":"X","bold":true}]'
  docx edit doc.docx --at p2-p5 --text "Rewrite this section as one paragraph."
  docx edit doc.docx --at p3-p7 --code-file new-snippet.go --language go
  docx edit doc.docx --at s0 --columns 2 --type continuous
  docx edit doc.docx --at eq0 --equation "x = \\\\frac{-b}{2a}" --display
  docx edit doc.docx --batch fills.jsonl            # fill many spans at once

Batch JSONL example (one edit per line):
  {"at": "p4:25-38", "text": "Delaware"}
  {"at": "t0:r1c2:p0", "text": "$4.2M"}
  {"at": "p9", "clear": "highlight"}
  {"at": "t1:r1c1:p0", "text": "June 8, 2026", "clear": "highlight"}
  {"at": "p1", "markdown": "## Revised heading"}
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const batchInput = parsed.values.batch as string | undefined;
	if (batchInput !== undefined) {
		return runEditBatch(filePath, batchInput, parsed.values);
	}

	const opts = await validateSingleShotOptions(filePath, parsed.values);
	if (typeof opts === "number") return opts;

	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	// Range locator (`pN-pM`): replaces a span of paragraphs as a unit. Section
	// edits don't make sense here (sN has its own grammar).
	if (isBlockRangeLocator(opts.locator)) {
		if (opts.spec.kind === "section") {
			return fail(
				"USAGE",
				"Range locators (pN-pM) don't accept --columns/--type — use sN for section edits",
				HELP,
			);
		}
		return commitRangeEdit(document, opts);
	}

	// Equation locator (`eqN`) targets an `<m:oMath>` inside a paragraph, not
	// the paragraph itself — resolves via the Equations lens, not the block
	// resolver.
	if (opts.spec.kind === "equation") {
		return commitEquationEdit(document, opts.spec, opts);
	}

	// Character-span locator (`pN:S-E` or a cell paragraph `tN:rRcC:pK:S-E`):
	// edit just those characters in place, inheriting the run's formatting.
	const spanTarget = spanLocatorTarget(opts.locator);
	if (spanTarget) return commitSpanEdit(document, spanTarget, opts);

	const blockRef = await resolveBlockOrFail(document, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	return commitBlockEdit(document, blockRef, opts);
}

function isBlockRangeLocator(locator: string): boolean {
	return /^p\d+-p\d+$/.test(locator);
}

/** A `--at` value that addresses a character span within one paragraph
 *  (`pN:S-E` or a cell paragraph `tN:rRcC:pK:S-E`) → `{blockId, span}`.
 *  Returns null for any other locator (whole block, range, entity), which
 *  falls through to the block-edit path. */
function spanLocatorTarget(
	locator: string,
): { blockId: string; span: { start: number; end: number } } | null {
	let parsed: Locator;
	try {
		parsed = parseLocator(locator);
	} catch {
		return null;
	}
	const target = locatorToBlockTarget(parsed);
	if (!target?.span) return null;
	return { blockId: target.blockId, span: target.span };
}

/** Span edit: replace the addressed characters in place via the Edit lens.
 *  v1 accepts only `--text` (the replacement inherits the existing run's
 *  rPr); formatting/paragraph flags belong to whole-paragraph edits. */
async function commitSpanEdit(
	document: Document,
	spanTarget: { blockId: string; span: { start: number; end: number } },
	opts: ValidatedOptions,
): Promise<number> {
	const spec = opts.spec;
	if (spec.kind !== "text" && spec.kind !== "clear") {
		return fail(
			"USAGE",
			"A character-span locator (pN:S-E) supports --text or --clear. Use a whole-paragraph locator (pN) for --markdown/--runs/--code.",
			HELP,
		);
	}
	if (spec.kind === "text") {
		if (spec.format.color || spec.format.bold || spec.format.italic) {
			return fail(
				"USAGE",
				"--color/--bold/--italic aren't supported on a character span — the replacement inherits the existing run's formatting. Edit the whole paragraph (pN) to set uniform run formatting.",
				HELP,
			);
		}
		if (spec.paragraphOptions.style || spec.paragraphOptions.alignment) {
			return fail(
				"USAGE",
				"--style/--alignment apply to a whole paragraph, not a character span (pN:S-E).",
				HELP,
			);
		}
	}

	const blockRef = await resolveBlockOrFail(document, spanTarget.blockId);
	if (typeof blockRef === "number") return blockRef;

	if (opts.dryRun) return respondDryRun(opts);

	try {
		if (spec.kind === "clear") {
			new Edit(document).clearFormatting(blockRef, spanTarget.span, spec.tags);
		} else {
			const edit = new Edit(document);
			edit.span(blockRef, spanTarget.span, spec.text, {
				authorFlag: opts.authorFlag,
				track: resolveTracked(document, opts.trackFlag),
			});
			// Combined `--text … --clear highlight` on a span: strip the named
			// formatting from the JUST-REPLACED range (offsets shift to the new
			// text length). This is the `find --highlight | edit` one-shot.
			if (opts.clearTags) {
				edit.clearFormattingNode(
					blockRef.node,
					{
						start: spanTarget.span.start,
						end: spanTarget.span.start + spec.text.length,
					},
					opts.clearTags,
				);
			}
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Single-block edit: section / task / paragraph dispatch through the Edit
 * lens. The lens handles tracked-vs-untracked, style ensures, and the
 * formatting-preservation decision. */
async function commitBlockEdit(
	document: Document,
	blockRef: BlockReference,
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.dryRun) return respondDryRun(opts);

	const track = resolveTracked(document, opts.trackFlag);
	try {
		const edit = new Edit(document);
		// The paragraph node the content edit produced — a combined `--clear`
		// strips formatting from THIS node afterward (it may be a freshly spliced
		// node, not the original blockRef).
		let resultNode: XmlNode | null = null;
		if (opts.spec.kind === "section") {
			edit.section(blockRef, opts.spec, { authorFlag: opts.authorFlag, track });
		} else if (opts.spec.kind === "task") {
			edit.taskToggle(blockRef, opts.spec.checked, {
				authorFlag: opts.authorFlag,
				track,
			});
			resultNode = blockRef.node;
		} else if (
			opts.spec.kind === "text" ||
			opts.spec.kind === "runs" ||
			opts.spec.kind === "code"
		) {
			resultNode = edit.paragraph(blockRef, opts.spec, {
				authorFlag: opts.authorFlag,
				noFormatting: opts.noFormatting,
				track,
			});
		} else if (opts.spec.kind === "markdown") {
			const resolved = await resolveMarkdownBlocks(document, opts.spec);
			if (typeof resolved === "number") return resolved;
			resultNode = edit.paragraph(blockRef, resolved, {
				authorFlag: opts.authorFlag,
				track,
			});
		} else if (opts.spec.kind === "clear") {
			edit.clearFormatting(blockRef, null, opts.spec.tags);
		} else {
			return fail("USAGE", "Unsupported edit spec for single-block locator");
		}
		// Combined content + `--clear`: strip the named formatting from the
		// post-edit paragraph (e.g. `--text "June 8, 2026" --clear highlight`).
		if (opts.clearTags && resultNode) {
			edit.clearFormattingNode(resultNode, null, opts.clearTags);
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Range replace path (`pN-pM`): resolve to a block range, hand to the Edit
 *  lens. The lens rejects tracked ranges that span a non-paragraph block. */
async function commitRangeEdit(
	document: Document,
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.spec.kind === "section") {
		// Type narrowing — caller already rejected this above.
		return fail("USAGE", "Section edits don't support range locators", HELP);
	}
	if (opts.spec.kind === "task") {
		return fail(
			"USAGE",
			"--task takes a single paragraph locator (pN), not a range",
			HELP,
		);
	}
	if (opts.spec.kind === "equation") {
		return fail(
			"USAGE",
			"--equation takes a single equation locator (eqN), not a paragraph range",
			HELP,
		);
	}

	const rangeRef: BlockRangeReference | number = await resolveBlockRangeOrFail(
		document,
		opts.locator,
	);
	if (typeof rangeRef === "number") return rangeRef;

	if (opts.dryRun) return respondDryRun(opts);

	let spec: ParagraphContentSpec;
	if (opts.spec.kind === "markdown") {
		const resolved = await resolveMarkdownBlocks(document, opts.spec);
		if (typeof resolved === "number") return resolved;
		spec = resolved;
	} else if (
		opts.spec.kind === "text" ||
		opts.spec.kind === "runs" ||
		opts.spec.kind === "code"
	) {
		spec = opts.spec;
	} else {
		return fail("USAGE", "Unsupported edit spec for range locator");
	}

	try {
		new Edit(document).range(rangeRef, spec, {
			authorFlag: opts.authorFlag,
			track: resolveTracked(document, opts.trackFlag),
		});
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Resolve an `eqN` locator, splice in a new OMML subtree, save. The spec
 *  carries optional `latex` (content swap) and `display` (mode toggle).
 *  At least one must change something, else it's a no-op error.
 *
 *  Tracking: when `<w:trackChanges/>` is on, the splice is replaced by a
 *  paired `<w:del>OLD</w:del><w:ins>NEW</w:ins>`. Word's accept-all picks
 *  the right equation but leaves an empty container skeleton next to it —
 *  a Word normalization quirk that's cosmetic. */
async function commitEquationEdit(
	document: Document,
	spec: Extract<EditSpec, { kind: "equation" }>,
	opts: ValidatedOptions,
): Promise<number> {
	if (spec.latex === undefined && spec.display === undefined) {
		return fail(
			"USAGE",
			"--equation requires --equation NEW_LATEX, --display, or --inline",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	try {
		new Equations(document).edit(spec.locator, {
			latex: spec.latex,
			display: spec.display,
			author: opts.authorFlag,
		});
	} catch (error) {
		if (error instanceof EquationNotFoundError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationStaleError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationParseError) {
			return fail(
				"USAGE",
				`Could not parse LaTeX equation: ${error.message}`,
				"Check the LaTeX syntax — temml accepts most KaTeX/MathJax LaTeX.",
			);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

async function validateSingleShotOptions(
	filePath: string,
	values: RawValues,
): Promise<ValidatedOptions | number> {
	const locator = values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);

	const paragraphOptions = await parseParagraphOptions(values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	const isSectionLocator = /^s\d+$/.test(locator);
	const spec = isSectionLocator
		? await validateSectionEdit(values)
		: await validateParagraphEdit(values, paragraphOptions);
	if (typeof spec === "number") return spec;

	// `--clear` combined with content (spec is a content kind, not clear-alone):
	// apply the content edit, then strip these tags. The dispatch reads
	// opts.clearTags after committing the content.
	let clearTags: Set<string> | undefined;
	if (values.clear !== undefined && spec.kind !== "clear") {
		const parsed = await parseClearTagsOrFail(values.clear as string[]);
		if (typeof parsed === "number") return parsed;
		clearTags = parsed;
	}

	return {
		filePath,
		locator,
		spec,
		authorFlag: values.author as string | undefined,
		trackFlag: Boolean(values.track),
		outputPath: values.output as string | undefined,
		dryRun: Boolean(values["dry-run"]),
		noFormatting: Boolean(values["no-formatting"]),
		...(clearTags ? { clearTags } : {}),
	};
}

/** Paragraph-level flags that are meaningless under `--markdown` /
 *  `--markdown-file` (the markdown source already encodes block styling).
 *  See `chooseContentSpec` in `cli/insert/index.tsx` for the symmetric
 *  rejection on the insert side. */
const MARKDOWN_INCOMPATIBLE_FLAGS = ["style", "alignment"] as const;

const OPTION_SPEC = {
	at: { type: "string" },
	batch: { type: "string" },
	text: { type: "string" },
	runs: { type: "string" },
	code: { type: "string" },
	"code-file": { type: "string" },
	markdown: { type: "string" },
	"markdown-file": { type: "string" },
	// Repeatable: `--clear highlight --clear underline` accumulates, and each
	// value may itself be a comma list (`--clear highlight,underline`).
	clear: { type: "string", multiple: true },
	language: { type: "string" },
	task: { type: "string" },
	equation: { type: "string" },
	display: { type: "boolean" },
	inline: { type: "boolean" },
	columns: { type: "string" },
	type: { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	color: { type: "string" },
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	author: { type: "string" },
	track: { type: "boolean" },
	"no-formatting": { type: "boolean" },
	...SAVE_FLAGS,
} as const;

type ValidatedOptions = {
	filePath: string;
	locator: string;
	spec: EditSpec;
	authorFlag?: string;
	/** `--track`: force tracked emission for this command even if the document's
	 *  global track-changes toggle is off. */
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
	/** Opt-out of word-level formatting preservation. When true, --text
	 *  produces a single fresh `<w:r>` with no rPr (today's behavior). */
	noFormatting: boolean;
	/** `--clear` riding along with a content flag (e.g. `--text X --clear
	 *  highlight`): apply the content edit, THEN strip these rPr tags. Whole-
	 *  paragraph locators only. Undefined when --clear isn't combined with
	 *  content (clear-alone is the `{kind:"clear"}` spec instead). */
	clearTags?: Set<string>;
};

type EditSpec =
	| { kind: "section"; columns?: number; sectionType?: SectionType }
	| {
			kind: "text";
			text: string;
			format: TextFormatting;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "runs"; runs: Run[]; paragraphOptions: ParagraphOptions }
	| {
			kind: "code";
			content: string;
			language?: string;
			paragraphOptions: ParagraphOptions;
	  }
	| {
			kind: "markdown";
			source: string;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "task"; checked: boolean }
	| { kind: "clear"; tags: Set<string> }
	| {
			kind: "equation";
			locator: string;
			/** `undefined` means "keep the existing LaTeX" (a pure display-mode
			 *  toggle); a string means replace the content. */
			latex: string | undefined;
			/** `undefined` means "keep the existing display flag"; a boolean
			 *  switches to that mode. */
			display: boolean | undefined;
	  };

type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
};

type RawValues = ReturnType<typeof parseArgs>["values"];

async function validateSectionEdit(
	values: RawValues,
): Promise<EditSpec | number> {
	if (values.text !== undefined || values.runs !== undefined) {
		return fail(
			"USAGE",
			"Section locators (sN) take --columns and --type, not --text/--runs",
			HELP,
		);
	}
	if (values.columns === undefined && values.type === undefined) {
		return fail("USAGE", "Section edit requires --columns and/or --type", HELP);
	}
	const sectionFlags = await parseSectionFlags(values);
	if (typeof sectionFlags === "number") return sectionFlags;
	return { kind: "section", ...sectionFlags };
}

/** Parse a `--clear` value (comma list of attrs, or "all") into the rPr tag set,
 *  or return a `fail()` exit code. Shared by the clear-alone spec and the
 *  combined content+clear path. */
async function parseClearTagsOrFail(
	clearFlag: string | string[],
): Promise<Set<string> | number> {
	// Accept a single value, a comma list, repeated --clear flags, or any mix.
	const raw = Array.isArray(clearFlag) ? clearFlag : [clearFlag];
	const names = raw
		.flatMap((entry) => entry.split(","))
		.map((name) => name.trim().toLowerCase())
		.filter(Boolean);
	if (names.length === 0) {
		return fail("USAGE", "--clear needs an attribute name, or 'all'", HELP);
	}
	const tags = resolveClearTags(names);
	if (!tags) {
		return fail(
			"USAGE",
			`--clear: unknown attribute in "${raw.join(",")}". Valid: ${CLEARABLE_ATTRS.join(", ")}, all`,
			HELP,
		);
	}
	return tags;
}

async function validateParagraphEdit(
	values: RawValues,
	paragraphOptions: ParagraphOptions,
): Promise<EditSpec | number> {
	if (values.columns !== undefined || values.type !== undefined) {
		return fail(
			"USAGE",
			"--columns and --type require a section locator (sN)",
			HELP,
		);
	}

	// `--clear` strips run formatting. It may stand ALONE (its own content kind)
	// or RIDE ALONG with a content flag — `--text X --clear highlight` fills the
	// paragraph then strips the highlight in one call (the canonical form-fill +
	// un-highlight move; targeted, so it won't nuke font size the way `all` does).
	// Combined-with-content is handled by the caller (it sets opts.clearTags after
	// the content edit); here we only return the clear-alone spec.
	const clearFlag = values.clear as string[] | undefined;
	if (clearFlag !== undefined) {
		if (
			values.equation !== undefined ||
			values.display === true ||
			values.inline === true
		) {
			return fail(
				"USAGE",
				"--clear can't be combined with --equation/--display/--inline",
				HELP,
			);
		}
		const hasContent = [
			"text",
			"runs",
			"code",
			"code-file",
			"markdown",
			"markdown-file",
			"task",
		].some((flag) => values[flag] !== undefined);
		if (!hasContent) {
			const tags = await parseClearTagsOrFail(clearFlag);
			if (typeof tags === "number") return tags;
			return { kind: "clear", tags };
		}
		// Combined with content — fall through to parse the content spec; the
		// assembler re-reads --clear into opts.clearTags and applies it after.
	}
	const text = values.text as string | undefined;
	const runsJson = values.runs as string | undefined;
	const codeInline = values.code as string | undefined;
	const codeFile = values["code-file"] as string | undefined;
	const language = values.language as string | undefined;
	const taskFlag = values.task as string | undefined;
	const equationFlag = values.equation as string | undefined;
	const displayFlag = values.display as boolean | undefined;
	const inlineFlag = values.inline as boolean | undefined;

	// `--equation` swaps the LaTeX (and optionally display mode) of an
	// `eqN`-addressed equation. `--display` / `--inline` alone (without
	// `--equation`) toggle the display mode but keep the existing LaTeX.
	const equationLike =
		equationFlag !== undefined || displayFlag === true || inlineFlag === true;
	if (equationLike) {
		const otherContent = [
			text !== undefined,
			runsJson !== undefined,
			codeInline !== undefined,
			codeFile !== undefined,
			taskFlag !== undefined,
		].filter(Boolean).length;
		if (otherContent > 0) {
			return fail(
				"USAGE",
				"--equation / --display / --inline cannot be combined with --text/--runs/--code/--code-file/--task",
				HELP,
			);
		}
		if (displayFlag && inlineFlag) {
			return fail(
				"USAGE",
				"--display and --inline are mutually exclusive",
				HELP,
			);
		}
		const displayMode: boolean | undefined = displayFlag
			? true
			: inlineFlag
				? false
				: undefined;
		return {
			kind: "equation",
			locator: values.at as string,
			latex: equationFlag,
			display: displayMode,
		};
	}

	// `--task` is its own content kind — it flips an existing task's state in
	// place rather than replacing the paragraph.
	if (taskFlag !== undefined) {
		const otherFlags = [
			text !== undefined,
			runsJson !== undefined,
			codeInline !== undefined,
			codeFile !== undefined,
		].filter(Boolean).length;
		if (otherFlags > 0) {
			return fail(
				"USAGE",
				"--task cannot be combined with --text, --runs, --code, or --code-file",
				HELP,
			);
		}
		const checked = parseTaskFlag(taskFlag);
		if (checked === null) {
			return fail(
				"USAGE",
				`--task must be "checked" or "unchecked", got "${taskFlag}"`,
				HELP,
			);
		}
		return { kind: "task", checked };
	}

	const markdownInline = values.markdown as string | undefined;
	const markdownFile = values["markdown-file"] as string | undefined;
	const contentFlags = [
		text !== undefined,
		runsJson !== undefined,
		codeInline !== undefined,
		codeFile !== undefined,
		markdownInline !== undefined,
		markdownFile !== undefined,
	].filter(Boolean).length;
	if (contentFlags === 0) {
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --code, --code-file, --markdown, --markdown-file, --task, or --equation",
			HELP,
		);
	}
	if (contentFlags > 1) {
		return fail(
			"USAGE",
			"Pass only one of --text, --runs, --code, --code-file, --markdown, --markdown-file",
			HELP,
		);
	}
	if (
		language !== undefined &&
		codeInline === undefined &&
		codeFile === undefined
	) {
		return fail("USAGE", "--language requires --code or --code-file", HELP);
	}

	if (text !== undefined) {
		return {
			kind: "text",
			text,
			format: {
				color: values.color as string | undefined,
				bold: values.bold as boolean | undefined,
				italic: values.italic as boolean | undefined,
			},
			paragraphOptions,
		};
	}

	if (codeInline !== undefined || codeFile !== undefined) {
		const content =
			codeInline !== undefined
				? codeInline
				: await loadCodeFile(codeFile as string);
		if (typeof content === "number") return content;
		return {
			kind: "code",
			content,
			...(language ? { language } : {}),
			paragraphOptions,
		};
	}

	if (markdownInline !== undefined || markdownFile !== undefined) {
		// Markdown encodes its own block styling — paragraph-level flags
		// would be silently dropped. Reject up front for the same reason
		// `--language` is gated to `--code` / `--code-file`.
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with --markdown / --markdown-file (the markdown source controls block-level styling)`,
				HELP,
			);
		}
		const source =
			markdownInline !== undefined
				? markdownInline
				: await loadMarkdownFile(markdownFile as string);
		if (typeof source === "number") return source;
		return { kind: "markdown", source, paragraphOptions };
	}

	const runs = await parseRunsArg(runsJson as string);
	if (typeof runs === "number") return runs;
	return { kind: "runs", runs, paragraphOptions };
}

/** Read content for `--markdown-file PATH`. `-` means stdin. */
async function loadMarkdownFile(path: string): Promise<string | number> {
	try {
		return path === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(path).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --markdown-file ${path}: ${message}`,
		);
	}
}

/** Resolve a `markdown` spec into a `markdown-blocks` spec by parsing the
 * source against the now-open document. The lens registers footnote bodies,
 * mints image rels, and provisions any styles the source references — all on
 * `document` — before returning the splice-ready blocks. */
async function resolveMarkdownBlocks(
	document: Document,
	spec: Extract<EditSpec, { kind: "markdown" }>,
): Promise<
	Extract<ParagraphContentSpec, { kind: "markdown-blocks" }> | number
> {
	try {
		const blocks = await new MarkdownImport(document).blocks(spec.source);
		return {
			kind: "markdown-blocks",
			blocks,
			paragraphOptions: spec.paragraphOptions,
		};
	} catch (error) {
		if (error instanceof MarkdownImportError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}
}

/** Read content for `--code-file PATH`. `-` means stdin (handled the same as
 *  `insert --code-file -`). Mirrors `cli/insert/index.tsx::resolveCodeSpec`. */
async function loadCodeFile(path: string): Promise<string | number> {
	try {
		return path === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(path).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --code-file ${path}: ${message}`,
		);
	}
}

async function parseParagraphOptions(
	values: RawValues,
): Promise<ParagraphOptions | number> {
	const out: ParagraphOptions = {};

	const styleValue = values.style as string | undefined;
	if (styleValue) out.style = styleValue;

	const alignmentValue = values.alignment as string | undefined;
	if (alignmentValue) {
		if (
			alignmentValue !== "left" &&
			alignmentValue !== "center" &&
			alignmentValue !== "right" &&
			alignmentValue !== "justify"
		) {
			return fail(
				"USAGE",
				`Invalid --alignment: ${alignmentValue}`,
				"Valid values: left, center, right, justify",
			);
		}
		out.alignment = alignmentValue;
	}

	return out;
}

async function respondDryRun(opts: ValidatedOptions): Promise<number> {
	await respond({
		operation: "edit",
		dryRun: true,
		path: opts.filePath,
		locator: opts.locator,
		...(opts.outputPath ? { output: opts.outputPath } : {}),
	});
	return EXIT.OK;
}

async function emitEditAck(opts: ValidatedOptions): Promise<number> {
	await respondAck({
		ok: true,
		operation: "edit",
		path: opts.outputPath ?? opts.filePath,
		locator: opts.locator,
	});
	return EXIT.OK;
}
