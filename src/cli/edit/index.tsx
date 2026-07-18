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
	type RunFormat,
	resolveClearTags,
	type XmlNode,
} from "@core";
import type { ParagraphOptions, TabStop } from "@core/blocks";
import { removeParagraphLine } from "@core/track-changes/replace";
import type { parseArgs } from "util";
import {
	batchExampleIntro,
	decodeInlineEscapes,
	hasRunFormatFlags,
	parseRunFormat,
	parseRunsArg,
	parseSpacingIndentFlags,
	pickContextualHelp,
	rejectShellMangledValue,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveBlockRangeOrFail,
	resolveTracked,
	respondAck,
	respondEditDryRun,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runEditBatch } from "./batch";
import {
	parseTabsValue,
	resolveTabsDirective,
	type TabsDirective,
} from "./tabs";

const AT_FORMS = describeForms(
	["paragraph", "span", "blockRange", "cellParagraph", "cellSpan"],
	"                      ",
);

const EDIT_HELP = `docx edit — replace content at a locator

Usage:
  docx edit FILE --at LOCATOR <content> [options]
  docx edit FILE --batch FILE.jsonl [options]   # many edits, one read
  docx edit FILE --batch -          [options]   # read JSONL from stdin

Examples:
${batchExampleIntro("Edit several blocks")}
  #   edits.jsonl:
  #     {"at":"p2","markdown":"New **bold** text"}
  #     {"at":"p4","text":"Delaware"}
  #     {"at":"p7:0-4","text":"ACME","bold":true}
  #     {"at":"p9","style":"Heading1"}
  #     {"at":"p5","text":""}
  #     {"at":"p6","text":"Acme Corp","clear":"highlight"}
  docx edit doc.docx --batch edits.jsonl
  # …or one at a time:
  docx find doc.docx "fill in state"                   # → p4:25-38
  docx edit doc.docx --at p4:25-38 --text "Delaware"   # replace just that span
  docx edit doc.docx --at p3 --markdown "## Revised heading"
  docx edit doc.docx --at p4 --text "Title" --bold     # fill + format in one call
  docx edit doc.docx --at p9-p38 --tabs right          # fix every wrapping tab line

Locator (required):
  --at LOCATOR      What to edit. One of:
${AT_FORMS}
                    A character span (pN:S-E, or tN:rRcC:pK:S-E in a cell)
                    replaces just those characters. More: \`docx info locators\`.

Content (required, UNLESS you pass only the formatting options below):
  --markdown TEXT   Replace with parsed GFM markdown (headings, lists, tables,
                    code, links, math, footnotes, CriticMarkup, …). Same dialect
                    as \`docx insert --markdown\`. A multi-block source expands —
                    the paragraph is replaced by however many blocks it parses to.
  --markdown-file PATH  Same as --markdown, but read content from PATH ("-" = stdin).
  --text TEXT       Replace with a single LITERAL run — every character verbatim.
                    Empty "" REMOVES the line (a table cell's last paragraph is
                    blanked, not deleted). To FORMAT it, add run-formatting flags
                    or use --markdown — see \`docx edit --text --help\`.
  --runs JSON       Replace with custom runs (Run[] JSON). Field list + the full
                    run-formatting flag list: \`docx edit --runs --help\`.
  --clear ATTRS     Strip run formatting in place, keeping the text. ATTRS is a
                    comma list (bold, italic, underline, highlight, shade, color,
                    font, size, …) or "all". Rides along with content on a
                    paragraph or span: \`--text "Delaware" --clear highlight\` fills
                    then un-highlights in one call. (Not tracked.)

Formatting options (except --tabs, not combinable with --markdown /
--markdown-file; pass them ALONE to reformat a block in place, keeping its
text — or add them to --text):
  --style NAME       Paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify
  --space-before PT / --space-after PT   Space above / below, in points
  --line-spacing N   A multiple (1, 1.5, 2), a name (single, double), or 15pt
  --indent-left IN / --indent-right IN   Indent, in inches (negative outdents)
  --first-line IN / --hanging IN         First-line / hanging indent, in inches
  --tabs SPEC        Replace the paragraph's tab stops. SPEC is one of:
                       right — a single RIGHT tab at the text margin
                       clear — remove all tab stops
                       left@1in,right@7.5in — explicit stops (left|right|center)
                     A RANGE locator fixes every wrapping line at once:
                     \`--at pN-pM --tabs right\`.

Batch (--batch PATH | -):
  Apply many edits from one read — the preferred way to update several blocks
  (locators do not shift between entries). Each JSONL line is one edit whose
  keys mirror the flags: {"at": LOCATOR, ...} with at most one content field
  ("text"/"markdown"/"runs"). "style"/"alignment"/"clear"/"bold"/"color"/…
  can ride along with content, or stand alone to format existing text.
  \`"text": ""\` or \`"delete": true\` removes the line. One entry per paragraph —
  to fill AND format the same paragraph, put both fields in that one entry.
  Ranges (pN-pM) aren't batchable; sections (sN) and equations (eqN) aren't
  edit's at all — use \`docx sections\` / \`docx equations edit\`.
  Don't pass --at/--text/… alongside --batch.

General options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record this edit as a tracked change even when the document's
                    toggle is off (check with \`docx track-changes list FILE\`).
  --no-formatting   Replace with a single fresh run; don't preserve rPr on
                    unchanged words
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0) — an in-place edit shifts
  nothing, so the edited locator is unchanged and there's nothing to mint.
  --verbose prints {ok:true, operation, path, locator}. Errors print {code, error,
  hint?} with a nonzero exit. Heads up: a locator you hold from BEFORE a structural
  edit (an insert/delete elsewhere renumbers ids) is stale — re-read after any
  insert/delete, or apply the whole set from one read with --batch.
`;

const EDIT_TEXT_HELP = `docx edit --text — replace a line's text and format it

Usage:
  docx edit FILE --at pN --text "New text" [formatting]
  docx edit FILE --at pN:S-E --text "New text" [formatting]   # just a span

Examples:
  docx edit doc.docx --at p4 --text "Delaware"
  docx edit doc.docx --at p4 --text "Title" --bold --color C00000
  docx edit doc.docx --at p4:4-13 --text "flawless" --italic
  docx edit doc.docx --at p4 --markdown "New **bold** text"

--text writes LITERAL characters — every character lands verbatim in a single
run. \`--text "**bold**"\` puts the literal ** characters in (use --markdown to
parse them instead). To get FORMATTED text there are two paths:

  1. Run-formatting flags that ride along with --text (they format the whole new
     run; on a span pN:S-E too):
       --bold            Bold                 --italic       Italic
       --underline       Underline (single)   --strike       Strikethrough
       --color HEX       Text color (C00000)  --highlight NAME  Highlighter
       --shade HEX       Background fill hex   --font NAME    Font family
       --size PT         Font size (points)    --caps         All caps
       --smallcaps       Small caps            --superscript / --subscript
     e.g. \`--at pN --text "Title" --bold --color C00000\`.

  2. Use --markdown instead of --text for inline syntax (**bold**, \`code\`,
     [links](url), ~~strike~~) and MIXED / multi-run formatting in one line:
       docx edit FILE --at pN --markdown "New **bold** and *italic* text"

Other --text behavior:
  --text ""         Empty removes the line (a table cell's last paragraph is
                    blanked, not deleted). To keep an empty spacer, use --runs '[]'.
  --clear ATTRS     Strip formatting while filling: \`--text "Delaware" --clear
                    highlight\` fills then un-highlights in one call.
  --no-formatting   Replace with a single fresh run, dropping the per-word rPr
                    preservation --text applies by default.

Run-formatting flags apply DIRECTLY — never recorded as tracked changes. To
SET formatting WITHOUT changing the text, see \`docx edit --runs --help\`.
`;

const EDIT_RUNS_HELP = `docx edit --runs — build runs from JSON, or set formatting on existing text

Two advanced surfaces. Prefer --text + flags or --markdown unless you need one.

Examples:
  docx edit doc.docx --at p2 --font "Times New Roman" --size 12
  docx edit doc.docx --at p4:4-13 --bold --color C00000
  docx edit doc.docx --at p0-p9 --italic           # format every paragraph in range
  docx edit doc.docx --at p2 --clear all           # strip all run formatting
  docx edit doc.docx --at p0 --runs '[{"type":"text","text":"X","bold":true}]'

--runs JSON — replace a paragraph with an explicit array of runs (Run[] JSON).
Each run object may carry:
  { "type": "text", "text": "…",
    "bold": true, "italic": true, "underline": true, "strike": true,
    "color": "C00000",         // hex, no '#'
    "highlight": "yellow",     // named highlighter
    "shade": "EEEEEE",         // background fill, hex
    "font": "Times New Roman", "size": 12,
    "caps": true, "smallcaps": true,
    "vertAlign": "superscript" | "subscript" }
  e.g. --runs '[{"type":"text","text":"X","bold":true},{"type":"text","text":" y"}]'

SET formatting on EXISTING text (no content flag — the text is kept). Target a
span (pN:S-E), a whole paragraph (pN), or a range (pN-pM) — paste a span from
\`docx find\`:
  --bold            Bold                 --italic       Italic
  --underline       Underline (single)   --strike       Strikethrough
  --color HEX       Text color (hex)     --highlight NAME  Highlighter
  --shade HEX       Background fill hex   --font NAME    Font family
  --size PT         Font size (points)    --caps         All caps
  --smallcaps       Small caps            --superscript / --subscript
  --clear ATTRS     Turn formatting OFF (comma list, or "all")

These apply DIRECTLY — never recorded as tracked changes, regardless of
--track or the document's track-changes toggle.
`;

export async function run(args: string[]): Promise<number> {
	const help = pickContextualHelp(args, {
		default: EDIT_HELP,
		text: EDIT_TEXT_HELP,
		runs: EDIT_RUNS_HELP,
	});
	const parsed = await tryParseArgs(args, OPTION_SPEC, help);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", EDIT_HELP);

	const batchInput = parsed.values.batch as string | undefined;
	if (batchInput !== undefined) {
		return runEditBatch(filePath, batchInput, parsed.values);
	}

	const opts = await validateSingleShotOptions(filePath, parsed.values);
	if (typeof opts === "number") return opts;

	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	// `--tabs` resolves now that the document (and its content width) is loaded:
	// `right` → a single right tab flush at the text margin, the cure for the
	// fragile LEFT tab `read` warns about.
	if (opts.tabsDirective) {
		injectTabsIntoSpec(
			opts.spec,
			resolveTabsDirective(opts.tabsDirective, document),
		);
	}

	// Range locator (`pN-pM`): replaces a span of paragraphs as a unit.
	if (isBlockRangeLocator(opts.locator)) {
		return commitRangeEdit(document, opts);
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

/** Span edit: address the characters in place via the Edit lens. Accepts
 *  `--text` (replace the characters, inheriting the run's rPr), `--clear` (strip
 *  formatting from the span), and the run-formatting flags (`--bold`/`--color`/…,
 *  which SET formatting on the span). `--text` may ride along with `--clear`
 *  and/or formatting flags to fill-then-format the just-written range in one
 *  call. Paragraph properties (`--style`/`--alignment`) stay whole-paragraph. */
async function commitSpanEdit(
	document: Document,
	spanTarget: { blockId: string; span: { start: number; end: number } },
	opts: ValidatedOptions,
): Promise<number> {
	const spec = opts.spec;
	if (
		spec.kind !== "text" &&
		spec.kind !== "clear" &&
		spec.kind !== "setFormat"
	) {
		return fail(
			"USAGE",
			"A character-span locator (pN:S-E) supports --text, --clear, or run-formatting flags (--bold/--color/--font/…). Use a whole-paragraph locator (pN) for --markdown/--runs.",
			EDIT_HELP,
		);
	}
	if (spec.kind === "text") {
		const paragraphOptions = spec.paragraphOptions;
		if (
			paragraphOptions.style ||
			paragraphOptions.alignment ||
			paragraphOptions.spacing ||
			paragraphOptions.indent ||
			paragraphOptions.tabs
		) {
			return fail(
				"USAGE",
				"--style/--alignment/--space-*/--line-spacing/--indent-*/--first-line/--hanging/--tabs apply to a whole paragraph, not a character span (pN:S-E). Use a whole-paragraph locator (pN).",
				EDIT_HELP,
			);
		}
	}

	const blockRef = await resolveBlockOrFail(document, spanTarget.blockId);
	if (typeof blockRef === "number") return blockRef;

	if (opts.dryRun) return respondDryRun(opts);

	try {
		const edit = new Edit(document);
		if (spec.kind === "clear") {
			edit.clearFormatting(blockRef, spanTarget.span, spec.tags);
		} else if (spec.kind === "setFormat") {
			edit.setFormatting(blockRef, spanTarget.span, spec.format);
		} else {
			edit.span(blockRef, spanTarget.span, spec.text, {
				authorFlag: opts.authorFlag,
				track: resolveTracked(document, opts.trackFlag),
			});
			// Combined `--text … --clear/--bold` on a span: clear then set the named
			// formatting on the JUST-REPLACED range (offsets shift to the new text
			// length). This is the `find … | edit` fill-and-format one-shot.
			const replaced = {
				start: spanTarget.span.start,
				end: spanTarget.span.start + spec.text.length,
			};
			if (opts.clearTags) {
				edit.clearFormattingNode(blockRef.node, replaced, opts.clearTags);
			}
			if (opts.setFormat) {
				edit.setFormattingNode(blockRef.node, replaced, opts.setFormat);
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

/** Single-block edit: paragraph content / props dispatch through the Edit
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
		if (opts.spec.kind === "text" || opts.spec.kind === "runs") {
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
				noFormatting: opts.noFormatting,
				track,
			});
		} else if (opts.spec.kind === "removeLine") {
			// Empty `--text` on a whole paragraph removes the line (cell-safe — a
			// table cell's last paragraph is blanked, not deleted). Same path as
			// `docx delete --at pN`.
			removeParagraphLine(document, blockRef, {
				track,
				author: opts.authorFlag,
			});
		} else if (opts.spec.kind === "clear") {
			edit.clearFormatting(blockRef, null, opts.spec.tags);
		} else if (opts.spec.kind === "setFormat") {
			edit.setFormatting(blockRef, null, opts.spec.format);
		} else if (opts.spec.kind === "paragraphProps") {
			resultNode = edit.paragraphProperties(
				blockRef,
				opts.spec.paragraphOptions,
				{
					authorFlag: opts.authorFlag,
					track,
				},
			);
		} else {
			return fail("USAGE", "Unsupported edit spec for single-block locator");
		}
		// Combined content + `--clear`/run-formatting: strip then set the named
		// formatting on the post-edit paragraph (e.g. `--text "June 8, 2026"
		// --clear highlight` or `--text "Title" --bold`). Clear runs first so an
		// explicit set wins on any shared property.
		if (opts.clearTags && resultNode) {
			edit.clearFormattingNode(resultNode, null, opts.clearTags);
		}
		if (opts.setFormat && resultNode) {
			edit.setFormattingNode(resultNode, null, opts.setFormat);
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
	if (opts.spec.kind === "removeLine") {
		return fail(
			"USAGE",
			"Empty --text removes a single paragraph, not a range — use `docx delete --at pN-pM` to remove a span of paragraphs.",
			EDIT_HELP,
		);
	}
	if (opts.spec.kind === "paragraphProps") {
		return commitRangeProps(document, opts, opts.spec.paragraphOptions);
	}
	if (opts.spec.kind === "setFormat") {
		return commitRangeSetFormat(document, opts, opts.spec.format);
	}
	// Run-formatting/clear flags riding along with a range content replace would
	// have to re-find every spliced-in paragraph; do it in a second, explicit call
	// instead. Reject loudly rather than silently dropping the ride-along.
	if (opts.setFormat) {
		return fail(
			"USAGE",
			"Set run formatting on a replaced range in a separate call: `edit --at pN-pM --bold` (or --color/--font/…) after the content edit.",
			EDIT_HELP,
		);
	}
	if (opts.clearTags) {
		return fail(
			"USAGE",
			"Strip run formatting on a replaced range in a separate call: `edit --at pN-pM --clear bold` (or --clear highlight/…) after the content edit.",
			EDIT_HELP,
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
	} else if (opts.spec.kind === "text" || opts.spec.kind === "runs") {
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

/** Properties-only RANGE edit (`--at pN-pM --tabs right`, etc.): apply the
 *  paragraph properties to each paragraph in the range, in place, without
 *  rewriting any content. This is the one-call tab-stop cure — `read` flags N
 *  wrapping lines, and `edit --at pN-pM --tabs right` fixes them all at once
 *  instead of N separate calls. `--tabs` only touches paragraphs that ALREADY
 *  have tab stops AND aren't list items (a bullet's tab is structural — see
 *  `scopeRangeProps`); `--style`/`--alignment` apply to every paragraph. */
async function commitRangeProps(
	document: Document,
	opts: ValidatedOptions,
	options: ParagraphOptions,
): Promise<number> {
	const rangeRef = await resolveBlockRangeOrFail(document, opts.locator);
	if (typeof rangeRef === "number") return rangeRef;
	if (opts.dryRun) return respondDryRun(opts);

	const track = resolveTracked(document, opts.trackFlag);
	let applied = 0;
	try {
		const edit = new Edit(document);
		for (let index = rangeRef.startIndex; index <= rangeRef.endIndex; index++) {
			const node = rangeRef.parent[index];
			if (!node || node.tag !== "w:p") continue;
			const perParagraph = scopeRangeProps(node, options);
			if (!perParagraph) continue;
			edit.paragraphProperties(
				{ node, parent: rangeRef.parent },
				perParagraph,
				{
					authorFlag: opts.authorFlag,
					track,
				},
			);
			applied++;
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	if (applied === 0) {
		return fail(
			"BLOCK_NOT_FOUND",
			options.tabs !== undefined
				? `No non-list paragraphs with tab stops in ${opts.locator} — --tabs only adjusts tab-using lines (the ones \`read\` flags with docx:layout), and skips bullets.`
				: `No paragraphs in ${opts.locator} to restyle.`,
		);
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Run-formatting RANGE edit (`--at pN-pM --bold`, etc.): set the formatting on
 *  every paragraph in the range, in place, keeping all text. Tables and other
 *  non-paragraph blocks in the range are skipped. */
async function commitRangeSetFormat(
	document: Document,
	opts: ValidatedOptions,
	format: RunFormat,
): Promise<number> {
	const rangeRef = await resolveBlockRangeOrFail(document, opts.locator);
	if (typeof rangeRef === "number") return rangeRef;
	if (opts.dryRun) return respondDryRun(opts);

	let applied = 0;
	try {
		const edit = new Edit(document);
		for (let index = rangeRef.startIndex; index <= rangeRef.endIndex; index++) {
			const node = rangeRef.parent[index];
			if (!node || node.tag !== "w:p") continue;
			edit.setFormatting({ node, parent: rangeRef.parent }, null, format);
			applied++;
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	if (applied === 0) {
		return fail(
			"BLOCK_NOT_FOUND",
			`No paragraphs in ${opts.locator} to format.`,
		);
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** The subset of `options` to apply to ONE paragraph in a range props edit:
 *  `--tabs` only rides along when the paragraph carries tab stops AND is not a
 *  list/numbered paragraph; style and alignment always apply. Returns null when
 *  nothing applies to this paragraph.
 *
 *  The list-paragraph exclusion is load-bearing: a bullet's `<w:pPr><w:tabs>` is
 *  the STRUCTURAL bullet-to-text tab, not a content alignment tab, and replacing
 *  it with the right-margin cure jumps the bullet text to the far margin (the
 *  résumé "Built…" → stray "B" corruption the fix-all hint caused). `read` never
 *  flags bullets (they have no `<w:tab/>` RUN), so the consolidated `--at pN-pM
 *  --tabs right` cure spans them only by min..max — and must skip them here so
 *  the one-call cure is safe to paste verbatim. */
function scopeRangeProps(
	node: XmlNode,
	options: ParagraphOptions,
): ParagraphOptions | null {
	const out: ParagraphOptions = {};
	if (options.style !== undefined) out.style = options.style;
	if (options.alignment !== undefined) out.alignment = options.alignment;
	// Spacing/indent apply to every paragraph in the range (like style/alignment),
	// unlike --tabs which is gated to existing tab-using, non-list lines.
	if (options.spacing !== undefined) out.spacing = options.spacing;
	if (options.indent !== undefined) out.indent = options.indent;
	if (
		options.tabs !== undefined &&
		paragraphHasTabStops(node) &&
		!isListParagraph(node)
	) {
		out.tabs = options.tabs;
	}
	if (
		out.style === undefined &&
		out.alignment === undefined &&
		out.spacing === undefined &&
		out.indent === undefined &&
		out.tabs === undefined
	)
		return null;
	return out;
}

/** True when a paragraph carries a `<w:pPr><w:tabs>` — i.e. it's a tab-using line
 *  (the kind `read`'s docx:layout warning targets). */
function paragraphHasTabStops(node: XmlNode): boolean {
	return node.findChild("w:pPr")?.findChild("w:tabs") !== undefined;
}

/** True when a paragraph is a list/numbered item (`<w:pPr><w:numPr>`). Its tab
 *  stops position the bullet text and must NOT be replaced by the range `--tabs`
 *  cure. */
function isListParagraph(node: XmlNode): boolean {
	return node.findChild("w:pPr")?.findChild("w:numPr") !== undefined;
}

async function validateSingleShotOptions(
	filePath: string,
	values: RawValues,
): Promise<ValidatedOptions | number> {
	const locator = values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", EDIT_HELP);

	// Equation editing moved to its own noun-verb command. Fire on the eqN
	// locator OR the equation flags (still declared in OPTION_SPEC so they parse
	// and hit this explicit redirect rather than a generic "unknown option").
	if (
		/^eq\d+$/.test(locator) ||
		values.equation !== undefined ||
		values.display === true ||
		values.inline === true
	) {
		return fail(
			"USAGE",
			"edit no longer edits equations — use `docx equations edit`",
			'e.g. `docx equations edit FILE --at eqN --equation "x^2"` (or --display/--inline). See `docx equations edit --help`.',
		);
	}

	// Task-checkbox toggling moved to its own noun-verb command. Fire on the
	// `--task` flag (still declared in OPTION_SPEC so it parses and hits this
	// explicit redirect rather than a generic "unknown option").
	if (values.task !== undefined) {
		return fail(
			"USAGE",
			"edit no longer toggles task checkboxes — use `docx tasks check` / `docx tasks uncheck`",
			"e.g. `docx tasks check FILE --at pN` (or `docx tasks uncheck FILE --at pN`). See `docx tasks --help`.",
		);
	}

	// Code blocks moved to their own noun-verb command. Hoisted here (like
	// equations/tasks above) so it fires for ANY locator — a section locator +
	// `--code` gets the redirect too, not a confusing "section needs --columns".
	if (
		values.code !== undefined ||
		values["code-file"] !== undefined ||
		values.language !== undefined
	) {
		return fail(
			"USAGE",
			"edit no longer builds code blocks — insert with `docx code add`, replace a block with `docx code edit --at pN`",
			"e.g. `docx code edit FILE --at pN --code-file snippet.py --language python`. See `docx code edit --help`.",
		);
	}

	// Section editing (columns / type / page geometry) moved to `docx sections`.
	// Fire on the sN locator OR the section flags (kept in OPTION_SPEC so they
	// parse and hit this redirect rather than a generic "unknown option").
	if (
		/^s\d+$/.test(locator) ||
		values.columns !== undefined ||
		values.type !== undefined
	) {
		return fail(
			"USAGE",
			"edit no longer edits sections — use `docx sections`",
			"e.g. `docx sections FILE --at sN --columns 2` (or --type/--margins/--orientation/--size). See `docx sections --help`.",
		);
	}

	const paragraphOptions = await parseParagraphOptions(values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	let tabsDirective: TabsDirective | undefined;
	if (values.tabs !== undefined) {
		const parsed = parseTabsValue(values.tabs as string);
		if ("error" in parsed) return fail("USAGE", parsed.error, parsed.hint);
		tabsDirective = parsed;
	}

	const spec = await validateParagraphEdit(values, paragraphOptions);
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

	// Run-formatting flags riding along with a content edit: apply the content,
	// then SET this formatting on the result (the set-side twin of clearTags). For
	// a WHOLE-PARAGRAPH `--text`, color/bold/italic land on the freshly built runs
	// via the text spec, so we drop them from the ride-along. On a SPAN, `edit.span`
	// only replaces text (it ignores the text spec's format), so every flag rides.
	let setFormat: RunFormat | undefined;
	if (
		spec.kind === "text" ||
		spec.kind === "runs" ||
		spec.kind === "markdown"
	) {
		const parsed = parseRunFormat(values);
		if (parsed && "error" in parsed) {
			return fail("USAGE", parsed.error, parsed.hint);
		}
		if (parsed) {
			const ride: RunFormat = { ...parsed };
			if (spec.kind === "text" && !spanLocatorTarget(locator)) {
				ride.color = undefined;
				ride.bold = undefined;
				ride.italic = undefined;
			}
			if (Object.values(ride).some((value) => value !== undefined)) {
				setFormat = ride;
			}
		}
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
		...(setFormat ? { setFormat } : {}),
		...(tabsDirective ? { tabsDirective } : {}),
	};
}

/** Paragraph-level flags that are meaningless under `--markdown` /
 *  `--markdown-file` (the markdown source already encodes block styling).
 *  See `chooseContentSpec` in `cli/insert/index.tsx` for the symmetric
 *  rejection on the insert side. */
// `--tabs` is deliberately ABSENT: `setTabsOnSpec`/`injectTabsIntoSpec` applies it
// to a markdown spec on purpose (a tab-stop fix is orthogonal to block styling).
// The spacing/indent flags ARE incompatible — the markdown source owns block-level
// layout, so they'd be silently dropped; reject them up front like style/alignment.
const MARKDOWN_INCOMPATIBLE_FLAGS = [
	"style",
	"alignment",
	"space-before",
	"space-after",
	"line-spacing",
	"indent-left",
	"indent-right",
	"first-line",
	"hanging",
] as const;

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
	// Task toggling moved to `docx tasks check`/`uncheck`, equations to `docx
	// equations edit`; these stay declared so `edit --at pN --task …` /
	// `--equation …` parse and hit the explicit redirects in
	// validateSingleShotOptions rather than a generic "unknown option" error.
	task: { type: "string" },
	equation: { type: "string" },
	display: { type: "boolean" },
	inline: { type: "boolean" },
	columns: { type: "string" },
	type: { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	tabs: { type: "string" },
	color: { type: "string" },
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	underline: { type: "boolean" },
	strike: { type: "boolean" },
	caps: { type: "boolean" },
	smallcaps: { type: "boolean" },
	superscript: { type: "boolean" },
	subscript: { type: "boolean" },
	font: { type: "string" },
	size: { type: "string" },
	highlight: { type: "string" },
	shade: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
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
	/** Run-formatting flags (`--bold`/`--color`/`--font`/…) riding along with a
	 *  content flag: apply the content edit, THEN set this formatting on the
	 *  resulting run(s). The set-side twin of `clearTags`. Undefined when no
	 *  format flag rides content (a content-free format edit is the `{kind:
	 *  "setFormat"}` spec instead). */
	setFormat?: RunFormat;
	/** `--tabs`: replace/clear the paragraph's tab stops. Parsed pre-open (the
	 *  document isn't loaded yet), resolved to concrete twips in `run()` once the
	 *  section's content width is available. */
	tabsDirective?: TabsDirective;
};

type EditSpec =
	| {
			kind: "text";
			text: string;
			format: TextFormatting;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "runs"; runs: Run[]; paragraphOptions: ParagraphOptions }
	| {
			kind: "markdown";
			source: string;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "removeLine" }
	| { kind: "clear"; tags: Set<string> }
	| { kind: "setFormat"; format: RunFormat }
	| { kind: "paragraphProps"; paragraphOptions: ParagraphOptions };

type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
};

type RawValues = ReturnType<typeof parseArgs>["values"];

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
		return fail(
			"USAGE",
			"--clear needs an attribute name, or 'all'",
			EDIT_HELP,
		);
	}
	const tags = resolveClearTags(names);
	if (!tags) {
		return fail(
			"USAGE",
			`--clear: unknown attribute in "${raw.join(",")}". Valid: ${CLEARABLE_ATTRS.join(", ")}, all`,
			EDIT_HELP,
		);
	}
	return tags;
}

async function validateParagraphEdit(
	values: RawValues,
	paragraphOptions: ParagraphOptions,
): Promise<EditSpec | number> {
	// `--clear` strips run formatting. It may stand ALONE (its own content kind)
	// or RIDE ALONG with a content flag — `--text X --clear highlight` fills the
	// paragraph then strips the highlight in one call (the canonical form-fill +
	// un-highlight move; targeted, so it won't nuke font size the way `all` does).
	// Combined-with-content is handled by the caller (it sets opts.clearTags after
	// the content edit); here we only return the clear-alone spec.
	const clearFlag = values.clear as string[] | undefined;
	if (clearFlag !== undefined) {
		const hasContent = ["text", "runs", "markdown", "markdown-file"].some(
			(flag) => values[flag] !== undefined,
		);
		if (!hasContent) {
			if (hasRunFormatFlags(values)) {
				return fail(
					"USAGE",
					"Strip formatting (--clear) and set formatting (--bold/--color/--font/…) can't combine in one call — do them in separate calls.",
					EDIT_HELP,
				);
			}
			const tags = await parseClearTagsOrFail(clearFlag);
			if (typeof tags === "number") return tags;
			return { kind: "clear", tags };
		}
		// Combined with content — fall through to parse the content spec; the
		// assembler re-reads --clear into opts.clearTags and applies it after.
	}
	const text = decodeInlineEscapes(values.text as string | undefined);
	const runsJson = values.runs as string | undefined;

	const markdownInline = decodeInlineEscapes(
		values.markdown as string | undefined,
	);
	const markdownFile = values["markdown-file"] as string | undefined;
	const contentFlags = [
		text !== undefined,
		runsJson !== undefined,
		markdownInline !== undefined,
		markdownFile !== undefined,
	].filter(Boolean).length;
	if (contentFlags === 0) {
		const hasProps = Boolean(
			paragraphOptions.style ||
				paragraphOptions.alignment ||
				paragraphOptions.spacing ||
				paragraphOptions.indent ||
				values.tabs !== undefined,
		);
		// Run-formatting edit: `--bold`/`--color`/`--font`/… with no content SET
		// the formatting on the EXISTING runs in place (a span, whole paragraph, or
		// range), keeping the text — the inverse of `--clear`.
		if (hasRunFormatFlags(values)) {
			if (hasProps) {
				return fail(
					"USAGE",
					"Set run formatting (--bold/--color/--font/…) and paragraph properties (--style/--alignment/--tabs) in separate calls, or use --text to set both on new content.",
					EDIT_HELP,
				);
			}
			const parsed = parseRunFormat(values);
			if (parsed === null) {
				return fail("USAGE", "No run-formatting flags to apply", EDIT_HELP);
			}
			if ("error" in parsed) return fail("USAGE", parsed.error, parsed.hint);
			return { kind: "setFormat", format: parsed };
		}
		// Properties-only edit: `--style`/`--alignment`/`--tabs` with no content
		// keeps the paragraph's existing runs and just re-applies the paragraph
		// properties. (Re-styling shouldn't force a dummy --text that would otherwise
		// replace the content and could drop direct run formatting.) `--tabs` alone
		// is the tab-stop cure for the `read` LEFT-tab wrapping warning.
		if (hasProps) {
			return { kind: "paragraphProps", paragraphOptions };
		}
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --markdown, or --markdown-file — or run-formatting flags (--bold/--color/--font/--size/--underline/…) to format the EXISTING text, or --style/--alignment/--tabs to adjust the paragraph in place",
			EDIT_HELP,
		);
	}
	if (contentFlags > 1) {
		return fail(
			"USAGE",
			"Pass only one of --text, --runs, --markdown, --markdown-file",
			EDIT_HELP,
		);
	}
	if (text !== undefined) {
		// Whole-paragraph empty --text REMOVES the line (the move weak agents
		// reach for when clearing a placeholder). It routes to the same cell-safe
		// removal as `docx delete --at pN` — splice the paragraph, or, when it's a
		// table cell's last paragraph, blank it in place so the `<w:tc>` keeps a
		// paragraph. A SPAN locator (pN:S-E) is EXEMPT: there `--text ""` deletes
		// just those characters in place. To keep an empty spacer paragraph
		// instead of removing the line, pass `--runs '[]'`.
		const at = values.at as string | undefined;
		if (text === "" && !(at && spanLocatorTarget(at))) {
			// Removal is exclusive — a co-passed --clear / run-format / paragraph
			// flag would silently no-op (the removed line has no node to format), and
			// the --batch path already rejects the same combo, so reject here too.
			if (
				values.clear !== undefined ||
				hasRunFormatFlags(values) ||
				paragraphOptions.style !== undefined ||
				paragraphOptions.alignment !== undefined ||
				paragraphOptions.spacing !== undefined ||
				paragraphOptions.indent !== undefined ||
				values.tabs !== undefined
			) {
				return fail(
					"USAGE",
					"Empty --text removes the line and can't combine with --clear / run-formatting / --style/--alignment/--tabs — drop those flags, or use --runs '[]' to keep an empty (formatted) spacer paragraph.",
					EDIT_HELP,
				);
			}
			return { kind: "removeLine" };
		}
		const mangled = await rejectShellMangledValue(text, "--text");
		if (typeof mangled === "number") return mangled;
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

	if (markdownInline !== undefined || markdownFile !== undefined) {
		// Markdown encodes its own block styling — paragraph-level flags
		// would be silently dropped. Reject up front.
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with --markdown / --markdown-file (the markdown source controls block-level styling)`,
				EDIT_HELP,
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

	const spacingIndent = parseSpacingIndentFlags(values);
	if ("error" in spacingIndent) {
		return fail("USAGE", spacingIndent.error, spacingIndent.hint);
	}
	if (spacingIndent.spacing) out.spacing = spacingIndent.spacing;
	if (spacingIndent.indent) out.indent = spacingIndent.indent;

	return out;
}

/** Set `tabs` on the spec's `paragraphOptions` (every paragraph-content kind that
 *  carries them). Section/clear specs have no paragraph options, so
 *  they're left untouched — `--tabs` with those locators is a no-op by design. */
function injectTabsIntoSpec(spec: EditSpec, tabs: TabStop[]): void {
	if (
		spec.kind === "text" ||
		spec.kind === "runs" ||
		spec.kind === "markdown" ||
		spec.kind === "paragraphProps"
	) {
		spec.paragraphOptions.tabs = tabs;
	}
}

function respondDryRun(opts: ValidatedOptions): Promise<number> {
	return respondEditDryRun(opts.filePath, opts.locator, opts.outputPath);
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
