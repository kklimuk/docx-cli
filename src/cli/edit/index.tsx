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
	type SectionType,
	type XmlNode,
} from "@core";
import type { ParagraphOptions, TabStop } from "@core/blocks";
import {
	EquationNotFoundError,
	EquationParseError,
	EquationStaleError,
	Equations,
} from "@core/equation";
import { firstInvalidRunFormat } from "@core/run-formatting";
import type { parseArgs } from "util";
import {
	parseRunsArg,
	parseSectionFlags,
	parseTaskFlag,
	rejectMarkdownInText,
	rejectShellMangledValue,
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
import {
	parseTabsValue,
	resolveTabsDirective,
	type TabsDirective,
} from "./tabs";

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

Paragraph content (one required for paragraph / range locators — UNLESS you pass
only --style/--alignment below, which restyle the paragraph in place):
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
  --tabs SPEC        Replace the paragraph's tab stops. SPEC is:
                       right   — a single RIGHT tab flush at the text margin.
                                 This is the CURE for the \`docx:layout … warn=…\`
                                 hint \`read\` prints on a line whose trailing
                                 content sits on a fixed LEFT tab: a long value
                                 (e.g. "San Francisco, CA") overflows the margin
                                 and WRAPS to a second line. A right tab at the
                                 margin right-aligns it instead, so it never wraps.
                       clear   — remove the paragraph's tab stops.
                       <list>  — explicit stops, e.g. \`right@7.5in\` or
                                 \`left@1in,right@7.5in\` (ALIGN@POSin, comma list).
  Pass any of --style/--alignment/--tabs ALONE (no content flag) to adjust the
  paragraph in place, keeping its text/runs: \`docx edit doc.docx --at p4 --style
  Heading1\`, \`docx edit doc.docx --at p9 --tabs right\`. --tabs also rides along
  with --text (fill the line AND fix its tab in one call), and works per-entry in
  --batch ({"at":"p9","text":"Harvard University\\tCambridge, MA","tabs":"right"}).
  ONE-CALL cure: a RANGE locator fixes every wrapping tab line at once —
  \`docx edit doc.docx --at p9-p38 --tabs right\` (the exact "fix-all" command
  \`read\` prints at the top when lines wrap; it skips paragraphs with no tab stops).

Run formatting (the inverse of --clear — SET formatting on EXISTING text):
  Pass any of these ALONE (no content flag) to format the text already there,
  keeping it. Target a span (pN:S-E), a whole paragraph (pN), or a range (pN-pM)
  — paste a span straight from \`docx find\`. They also ride along with --text to
  fill AND format in one call (e.g. \`--text "Title" --bold\`).
  --bold            Bold
  --italic          Italic
  --underline       Underline (single). Other styles via --runs JSON.
  --strike          Strikethrough
  --color HEX       Text color, hex (e.g., 800080 for purple; no '#')
  --highlight NAME  Highlighter color (yellow, green, cyan, magenta, red, …)
  --shade HEX       Background fill, arbitrary hex (no '#')
  --font NAME       Font family (e.g., "Times New Roman")
  --size PT         Font size in points (e.g., 12 or 11.5)
  --caps            ALL CAPS (display only — the text stays as typed)
  --smallcaps       Small caps
  --superscript     Superscript    --subscript    Subscript
  To turn a property OFF, use --clear (e.g. \`--clear bold\`). Like --clear and
  --style, formatting changes apply DIRECTLY — they are NOT recorded as tracked
  changes (Word's <w:rPrChange> isn't modeled), regardless of --track / the
  document's track-changes toggle.

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
  docx edit doc.docx --at p4:4-13 --bold --color C00000 # bold + red that span
  docx edit doc.docx --at p2 --font "Times New Roman" --size 12  # restyle a paragraph's runs
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
  {"at": "p3:0-5", "bold": true, "color": "C00000"}
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

	// `--tabs` resolves now that the document (and its content width) is loaded:
	// `right` → a single right tab flush at the text margin, the cure for the
	// fragile LEFT tab `read` warns about.
	if (opts.tabsDirective) {
		injectTabsIntoSpec(
			opts.spec,
			resolveTabsDirective(opts.tabsDirective, document),
		);
	}

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
			"A character-span locator (pN:S-E) supports --text, --clear, or run-formatting flags (--bold/--color/--font/…). Use a whole-paragraph locator (pN) for --markdown/--runs/--code.",
			HELP,
		);
	}
	if (spec.kind === "text") {
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
		} else if (opts.spec.kind === "setFormat") {
			edit.setFormatting(blockRef, null, opts.spec.format);
		} else if (opts.spec.kind === "paragraphProps") {
			resultNode = edit.paragraphProperties(
				blockRef,
				opts.spec.paragraphOptions,
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
			HELP,
		);
	}
	if (opts.clearTags) {
		return fail(
			"USAGE",
			"Strip run formatting on a replaced range in a separate call: `edit --at pN-pM --clear bold` (or --clear highlight/…) after the content edit.",
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

	let applied = 0;
	try {
		const edit = new Edit(document);
		for (let index = rangeRef.startIndex; index <= rangeRef.endIndex; index++) {
			const node = rangeRef.parent[index];
			if (!node || node.tag !== "w:p") continue;
			const perParagraph = scopeRangeProps(node, options);
			if (!perParagraph) continue;
			edit.paragraphProperties({ node, parent: rangeRef.parent }, perParagraph);
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

	let tabsDirective: TabsDirective | undefined;
	if (values.tabs !== undefined) {
		const parsed = parseTabsValue(values.tabs as string);
		if ("error" in parsed) return fail("USAGE", parsed.error, parsed.hint);
		tabsDirective = parsed;
	}

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

	// Run-formatting flags riding along with a content edit: apply the content,
	// then SET this formatting on the result (the set-side twin of clearTags). For
	// a WHOLE-PARAGRAPH `--text`, color/bold/italic land on the freshly built runs
	// via the text spec, so we drop them from the ride-along. On a SPAN, `edit.span`
	// only replaces text (it ignores the text spec's format), so every flag rides.
	let setFormat: RunFormat | undefined;
	if (
		spec.kind === "text" ||
		spec.kind === "runs" ||
		spec.kind === "code" ||
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
	| { kind: "setFormat"; format: RunFormat }
	| { kind: "paragraphProps"; paragraphOptions: ParagraphOptions }
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
	// A section break has no runs — run-formatting/clear flags would silently do
	// nothing, so reject them with a targeted message (mirrors the --task/--equation
	// guards) instead of letting them fall through to the columns/type check.
	if (hasRunFormatFlags(values) || values.clear !== undefined) {
		return fail(
			"USAGE",
			"Section locators (sN) take --columns and --type — run-formatting flags (--bold/--color/--font/…) and --clear apply to a paragraph's runs, which a section break has none of.",
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

/** The flags that SET run formatting on existing text (the inverse of `--clear`).
 *  `color`/`bold`/`italic` are shared with the `--text` content ride-along; the
 *  rest are set-only. Used to detect a content-free format edit and to reject
 *  nonsensical combinations (e.g. with `--equation`/`--task`). */
const RUN_FORMAT_FLAGS = [
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

/** True when the invocation carries any run-formatting flag. Booleans default to
 *  `false` from parseArgs (not undefined), so we test against both. */
function hasRunFormatFlags(values: RawValues): boolean {
	return RUN_FORMAT_FLAGS.some(
		(flag) => values[flag] !== undefined && values[flag] !== false,
	);
}

/** Normalize a hex color for `<w:rPr>` (`w:color`/`w:shd@w:fill`): strip a single
 *  leading `#` so `--color "#FF0000"` becomes the schema-valid `FF0000` the help
 *  promises (ST_HexColor has no `#`). Other values pass through unchanged (Word
 *  degrades unknown colors gracefully, matching the `--runs`/markdown ingress). */
export function normalizeHexColor(value: string): string {
	return value.startsWith("#") ? value.slice(1) : value;
}

/** Build a `RunFormat` from the formatting flags, or `null` if none are set.
 *  Returns a `{ error }` shape on a bad value (size, mutually-exclusive
 *  super/subscript, or an out-of-range highlight) so the caller can `fail()`. */
function parseRunFormat(
	values: RawValues,
): RunFormat | null | { error: string; hint?: string } {
	const format: RunFormat = {};
	if (values.bold) format.bold = true;
	if (values.italic) format.italic = true;
	if (values.strike) format.strike = true;
	if (values.caps) format.allCaps = true;
	if (values.smallcaps) format.smallCaps = true;
	if (values.underline) format.underline = "single";
	if (values.color !== undefined)
		format.color = normalizeHexColor(values.color as string);
	if (values.font !== undefined) format.font = values.font as string;
	if (values.highlight !== undefined)
		format.highlight = values.highlight as string;
	if (values.shade !== undefined)
		format.shade = normalizeHexColor(values.shade as string);

	if (values.superscript && values.subscript) {
		return { error: "--superscript and --subscript are mutually exclusive" };
	}
	if (values.superscript) format.vertAlign = "superscript";
	if (values.subscript) format.vertAlign = "subscript";

	if (values.size !== undefined) {
		const points = Number.parseFloat(values.size as string);
		if (!Number.isFinite(points) || points <= 0) {
			return {
				error: `Invalid --size: ${values.size}`,
				hint: "Pass a positive point size, e.g. --size 12 or --size 11.5.",
			};
		}
		format.sizeHalfPoints = Math.round(points * 2);
	}

	const invalid = firstInvalidRunFormat(format);
	if (invalid) {
		return {
			error: `Invalid --${invalid.field}: ${invalid.value}`,
			hint: `Use ${invalid.valid}.`,
		};
	}

	if (Object.keys(format).length === 0) return null;
	return format;
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
			if (hasRunFormatFlags(values)) {
				return fail(
					"USAGE",
					"Strip formatting (--clear) and set formatting (--bold/--color/--font/…) can't combine in one call — do them in separate calls.",
					HELP,
				);
			}
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
		if (hasRunFormatFlags(values)) {
			return fail(
				"USAGE",
				"--equation/--display/--inline can't combine with run-formatting flags (--bold/--color/--font/…)",
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
		if (hasRunFormatFlags(values)) {
			return fail(
				"USAGE",
				"--task can't combine with run-formatting flags (--bold/--color/--font/…)",
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
		const hasProps = Boolean(
			paragraphOptions.style ||
				paragraphOptions.alignment ||
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
					HELP,
				);
			}
			const parsed = parseRunFormat(values);
			if (parsed === null) {
				return fail("USAGE", "No run-formatting flags to apply", HELP);
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
			"Missing content: pass --text, --runs, --code, --code-file, --markdown, --markdown-file, --task, or --equation — or run-formatting flags (--bold/--color/--font/--size/--underline/…) to format the EXISTING text, or --style/--alignment/--tabs to adjust the paragraph in place",
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
		// Whole-paragraph empty --text leaves an invisible, space-consuming blank
		// paragraph rather than removing the line — a weak-agent trap (the instinct
		// is "set the inapplicable section to empty → it's gone"). Redirect to the
		// honest verb. A SPAN locator (pN:S-E) is EXEMPT: there `--text ""` deletes
		// just those characters in place (the paragraph keeps its other content) — a
		// legitimate, common move (e.g. strip an inline `[Note: …]`).
		const at = values.at as string | undefined;
		if (text === "" && !(at && spanLocatorTarget(at))) {
			return fail(
				"USAGE",
				`Empty --text leaves a blank paragraph in place, it doesn't remove the line.`,
				`To DELETE the paragraph: \`docx delete --at ${at ?? "pN"}\` (or \`docx delete --batch\` for many). To delete just SOME characters, use a span locator (\`--at pN:S-E --text ""\`). To blank the paragraph but keep an empty spacer, pass --runs '[]'. Help:\n${HELP}`,
			);
		}
		const rejected = await rejectMarkdownInText(text, HELP);
		if (typeof rejected === "number") return rejected;
		const mangled = await rejectShellMangledValue(text, HELP, "--text");
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

/** Set `tabs` on the spec's `paragraphOptions` (every paragraph-content kind that
 *  carries them). Section/task/clear/equation specs have no paragraph options, so
 *  they're left untouched — `--tabs` with those locators is a no-op by design. */
function injectTabsIntoSpec(spec: EditSpec, tabs: TabStop[]): void {
	if (
		spec.kind === "text" ||
		spec.kind === "runs" ||
		spec.kind === "code" ||
		spec.kind === "markdown" ||
		spec.kind === "paragraphProps"
	) {
		spec.paragraphOptions.tabs = tabs;
	}
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
