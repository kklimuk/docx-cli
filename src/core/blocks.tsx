import type {
	ParagraphIndent,
	ParagraphSpacing,
	Run,
	TextRun,
} from "./ast/types";
import { w } from "./jsx";
import { type NullableXmlNode, XmlNode } from "./parser";
import { TaskCheckbox } from "./task-list";
import type { TrackedMeta } from "./track-changes";

/** A `<w:p>`. Pass `runs` for full control, or `text` (+ optional run-level
 * formatting) for a single-run paragraph. `runs` wins if both are given; with
 * neither, an empty paragraph is emitted.
 *
 * When `taskState` is set, prepends a `<w:sdt><w14:checkbox/></w:sdt>` content
 * control plus a space run before the user content — the shape Word, Pandoc,
 * and LibreOffice all emit for GFM task list items. The detection mirror in
 * `core/ast/read.ts::detectLeadingCheckbox` strips this prefix back out on
 * read, so a paragraph round-trips losslessly. */
export function Paragraph({
	style,
	alignment,
	list,
	taskState,
	tabs,
	spacing,
	indent,
	text,
	runs,
	...formatting
}: ParagraphOptions &
	Partial<EmittableTextFormatting> & {
		text?: string;
		runs?: Run[];
	}): XmlNode {
	const resolvedRuns: Run[] = runs ?? textToRuns(text ?? "", formatting);
	return (
		<w.p>
			<ParagraphProperties
				options={{ style, alignment, list, tabs, spacing, indent }}
			/>
			{taskState && <TaskCheckbox checked={taskState === "checked"} />}
			{taskState && (
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}> </w.t>
				</w.r>
			)}
			{resolvedRuns.map((run) => (
				<RunElement run={run} />
			))}
		</w.p>
	);
}

/** Turn a `text` prop into runs, mapping an embedded newline to a `<w:br/>`
 *  line break and a tab to `<w:tab/>` (run-level formatting rides on each text
 *  segment). This keeps line-per-line `--text` — verse, addresses, signature
 *  blocks — line-per-line in the document instead of collapsing to one wrapped
 *  line (Word swallows a literal `\n` inside `<w:t>`). Mirrors `textWithBreaks`
 *  on the markdown-import side; `read` emits a newline/tab back so it round-trips.
 *  Plain single-line text returns one run, byte-identical to before. */
function textToRuns(
	text: string,
	formatting: Partial<EmittableTextFormatting>,
): Run[] {
	if (!text.includes("\n") && !text.includes("\t")) {
		return [{ type: "text", text, ...formatting }];
	}
	const runs: Run[] = [];
	for (const segment of text.split(/(\n|\t)/)) {
		if (segment === "") continue;
		if (segment === "\n") runs.push({ type: "break", kind: "line" });
		else if (segment === "\t") runs.push({ type: "tab" });
		else runs.push({ type: "text", text: segment, ...formatting });
	}
	return runs;
}

export type ParagraphOptions = {
	style?: string;
	alignment?: "left" | "center" | "right" | "justify";
	list?: { level: number; numId: number };
	taskState?: "checked" | "unchecked";
	/** Paragraph tab stops (`<w:tabs>`), in twips. A non-empty array REPLACES the
	 *  paragraph's existing tab stops; an EMPTY array CLEARS them. Used by
	 *  `edit --tabs` to cure the fragile right-edge LEFT tab (which left-aligns
	 *  trailing content from a fixed point so a long value overflows the margin and
	 *  wraps) by swapping it for a RIGHT tab at the margin. `undefined` = leave the
	 *  paragraph's tabs untouched. */
	tabs?: TabStop[];
	/** `<w:spacing>` — paragraph spacing (before/after) + line spacing. Each set
	 *  attribute is applied/merged onto an existing `<w:spacing>`; `undefined` =
	 *  leave untouched. Units: twips for before/after, 240ths for line (lineRule
	 *  "auto"). */
	spacing?: ParagraphSpacing;
	/** `<w:ind>` — paragraph indentation (left/right/firstLine/hanging), twips.
	 *  Each set attribute merges onto an existing `<w:ind>`; setting `firstLine`
	 *  clears `hanging` and vice versa (same slot). `undefined` = leave untouched. */
	indent?: ParagraphIndent;
};

/** One `<w:tab>` entry: `align` is the `w:val` (left/right/center/…), `pos` the
 *  `w:pos` in twips (absolute from the left text margin). */
export type TabStop = { align: string; pos: number };

/** Normalize a tab stop's `w:val` to the canonical alignment vocabulary so the
 *  AST reader, the `docx:layout` wrap detector, and the page-setup tab-reflow cure
 *  all classify a stop the same way (they were diverging on the LTR-aware aliases
 *  `start`/`end` and a missing/empty val). `start`→`left`, `end`→`right`, missing
 *  or empty→`left`; everything else (center/right/decimal/bar/clear/num) passes
 *  through. Keep the three consumers reading through THIS function. */
export function normalizeTabAlign(raw: string | undefined): string {
	if (raw === undefined || raw === "" || raw === "start") return "left";
	if (raw === "end") return "right";
	return raw;
}

type EmittableTextFormatting = Pick<
	TextRun,
	| "color"
	| "colorTheme"
	| "colorThemeTint"
	| "colorThemeShade"
	| "highlight"
	| "shade"
	| "bold"
	| "italic"
	| "underline"
	| "underlineColor"
	| "strike"
	| "font"
	| "sizeHalfPoints"
>;

/** A paragraph that belongs to a numbered or bulleted list. Sets pStyle to
 * "ListParagraph" (the canonical Word style for list items — caller is
 * responsible for ensuring it via `view.ensureStyles().ensureStyle("ListParagraph")`) and
 * emits the `<w:numPr>` reference. `numId` must come from
 * `view.ensureNumbering().allocate(kind)` in [src/core/numbering.tsx]. */
export function ListParagraph({
	numId,
	level,
	alignment,
	taskState,
	text,
	runs,
}: {
	numId: number;
	level: number;
	alignment?: "left" | "center" | "right" | "justify";
	taskState?: "checked" | "unchecked";
	text?: string;
	runs?: Run[];
}): XmlNode {
	return (
		<Paragraph
			style="ListParagraph"
			alignment={alignment}
			list={{ level, numId }}
			taskState={taskState}
			text={text}
			runs={runs}
		/>
	);
}

/** A GFM task list item — a `<ListParagraph>` whose body is prefixed with a
 * Word checkbox content control. Sugar over `ListParagraph` with `taskState`
 * set; the SDT prefix is materialized inside `Paragraph` itself.
 *
 * @public Staged for the S8 markdown walker's `- [ ]` / `- [x]` task list
 * items; CLI authoring today goes through `insert --task` / `edit --task`
 * which build the paragraph via `Paragraph({ list, taskState })` directly. */
export function TaskListItem({
	numId,
	level,
	checked,
	text,
	runs,
}: {
	numId: number;
	level: number;
	checked: boolean;
	text?: string;
	runs?: Run[];
}): XmlNode {
	return (
		<ListParagraph
			numId={numId}
			level={level}
			taskState={checked ? "checked" : "unchecked"}
			text={text}
			runs={runs}
		/>
	);
}

/** Emits a run as fresh OOXML. The fresh-emission path supports text/break/tab
 * — the run types an agent can plausibly hand-author via `--runs '[...]'`.
 * Round-tripped runs that we surface but can't re-emit (image, equation,
 * footnoteRef, chart) are silently dropped: their underlying XML lives in the
 * source document and is never re-serialized through this path, so a
 * `read | jq | edit` workflow degrades gracefully instead of crashing. */
export function RunElement({ run }: { run: Run }): NullableXmlNode {
	if (run.type === "text") return <TextRunElement run={run} />;
	if (run.type === "break") {
		return (
			<w.r>
				<w.br w-type={run.kind === "line" ? undefined : run.kind} />
			</w.r>
		);
	}
	if (run.type === "tab") {
		return (
			<w.r>
				<w.tab />
			</w.r>
		);
	}
	return null;
}

/** Mutate a paragraph's children list in place to apply `--style` /
 *  `--alignment` to its `<w:pPr>`. Creates a pPr if absent. Used by the
 *  tracked-edit helpers in `core/track-changes/` and any caller that wants to
 *  swap paragraph properties without rebuilding the whole paragraph. The
 *  `rebuilt` list is the paragraph's `children` array — usually already
 *  filtered to non-run nodes by the caller. */
export function applyParagraphOptionsInPlace(
	rebuilt: XmlNode[],
	options: ParagraphOptions,
): void {
	if (
		!options.style &&
		!options.alignment &&
		!options.tabs &&
		!options.spacing &&
		!options.indent
	)
		return;
	let pPr = rebuilt.find((child) => child.tag === "w:pPr");
	if (!pPr) {
		pPr = new XmlNode("w:pPr");
		rebuilt.unshift(pPr);
	}
	if (options.tabs) {
		// A non-empty array replaces the paragraph's tab stops; an empty array
		// clears them. Drop the existing <w:tabs> either way, then re-add when there
		// are stops to set — keeping CT_PPr order via insertPprChildInOrder.
		pPr.children = pPr.children.filter((child) => child.tag !== "w:tabs");
		if (options.tabs.length > 0) {
			const tabsNode = new XmlNode("w:tabs");
			for (const stop of options.tabs) {
				tabsNode.children.push(
					new XmlNode("w:tab", {
						"w:val": stop.align,
						"w:pos": String(Math.round(stop.pos)),
					}),
				);
			}
			insertPprChildInOrder(pPr, tabsNode);
		}
	}
	if (options.style) {
		const existingStyle = pPr.findChild("w:pStyle");
		if (existingStyle) {
			existingStyle.setAttribute("w:val", options.style);
		} else {
			insertPprChildInOrder(
				pPr,
				new XmlNode("w:pStyle", { "w:val": options.style }),
			);
		}
	}
	if (options.alignment) {
		const existingJc = pPr.findChild("w:jc");
		if (existingJc) {
			existingJc.setAttribute("w:val", options.alignment);
		} else {
			insertPprChildInOrder(
				pPr,
				new XmlNode("w:jc", { "w:val": options.alignment }),
			);
		}
	}
	if (options.spacing) {
		const node = findOrCreatePprChild(pPr, "w:spacing");
		const { before, after, line, lineRule } = options.spacing;
		if (before !== undefined) node.setAttribute("w:before", String(before));
		if (after !== undefined) node.setAttribute("w:after", String(after));
		if (line !== undefined) {
			node.setAttribute("w:line", String(line));
			node.setAttribute("w:lineRule", lineRule ?? "auto");
		}
	}
	if (options.indent) {
		const node = findOrCreatePprChild(pPr, "w:ind");
		const { left, right, firstLine, hanging } = options.indent;
		if (left !== undefined) node.setAttribute("w:left", String(left));
		if (right !== undefined) node.setAttribute("w:right", String(right));
		// firstLine and hanging share a slot — setting one clears the other.
		if (firstLine !== undefined) {
			node.setAttribute("w:firstLine", String(firstLine));
			delete node.attributes["w:hanging"];
		}
		if (hanging !== undefined) {
			node.setAttribute("w:hanging", String(hanging));
			delete node.attributes["w:firstLine"];
		}
	}
}

/** Find a `<w:pPr>` child by tag, or create + splice it at its CT_PPr slot.
 *  Returns the live node so the caller can merge attributes onto it (preserving
 *  any the caller didn't set — e.g. a quote's existing `<w:ind w:left>`). */
function findOrCreatePprChild(pPr: XmlNode, tag: string): XmlNode {
	const existing = pPr.findChild(tag);
	if (existing) return existing;
	const created = new XmlNode(tag);
	insertPprChildInOrder(pPr, created);
	return created;
}

/** Find the paragraph's `<w:pPr>`, creating an empty one (as the first child) if
 *  absent. The home for direct paragraph properties. */
export function ensureParagraphProperties(paragraph: XmlNode): XmlNode {
	const existing = paragraph.findChild("w:pPr");
	if (existing) return existing;
	const created = new XmlNode("w:pPr");
	paragraph.children.unshift(created);
	return created;
}

/** Snapshot the prior `<w:pPr>` into a `<w:pPrChange>` marker (its last child),
 *  so a subsequent property mutation is a tracked revision Word can accept/reject.
 *  The paragraph analog of `wrapSectPrChange`: clone every current pPr child
 *  EXCEPT a pre-existing `<w:pPrChange>` (and an inline `<w:sectPr>`) into
 *  `<w:pPrChange><w:pPr>…</w:pPr>`, replacing any prior marker. MUST be called
 *  BEFORE the live mutation so the snapshot captures the prior state. Empirically
 *  matches Word for Mac's shape
 *  (`<w:pPrChange w:id w:author w:date><w:pPr>…prior…</w:pPr></w:pPrChange>`). */
export function wrapPprChange(pPr: XmlNode, meta: TrackedMeta): void {
	// The snapshot's inner `<w:pPr>` is CT_PPrChange's pPr — type CT_PPrBase,
	// which does NOT permit `<w:sectPr>` (a section break is tracked via a
	// separate `<w:sectPrChange>`, never inside a pPrChange). Cloning an inline
	// sectPr here would make Word reject the file ("unreadable content") AND
	// duplicate the section break. Filter it out of the snapshot; the live sectPr
	// stays put. Keep the paragraph-mark `<w:rPr>` — Word includes it.
	const prior = pPr.children.filter(
		(child) => child.tag !== "w:pPrChange" && child.tag !== "w:sectPr",
	);
	injectPprChange(pPr, prior, meta);
}

/** Inject a `<w:pPrChange>` carrying an EXPLICIT prior-pPr snapshot into `pPr`,
 *  replacing any existing marker. Unlike `wrapPprChange` (which snapshots the
 *  node's CURRENT children), this takes the prior children directly — used when
 *  the live pPr already holds the NEW properties (a freshly-built paragraph for a
 *  content+props ride-along edit), so the prior state must be supplied separately.
 *  The caller is responsible for excluding a pre-existing pPrChange and an inline
 *  `<w:sectPr>` from `priorChildren` (CT_PPrBase forbids sectPr in the snapshot). */
export function injectPprChange(
	pPr: XmlNode,
	priorChildren: XmlNode[],
	meta: TrackedMeta,
): void {
	pPr.children = pPr.children.filter((child) => child.tag !== "w:pPrChange");
	const change = new XmlNode(
		"w:pPrChange",
		{
			"w:id": String(meta.revisionId),
			"w:author": meta.author,
			"w:date": meta.date,
		},
		[
			new XmlNode(
				"w:pPr",
				{},
				priorChildren.map((child) => child.clone()),
			),
		],
	);
	insertPprChildInOrder(pPr, change);
}

/** True when `options` carries any direct paragraph property (a `<w:pPr>` child):
 *  style/alignment/tabs/spacing/indent. The gate for "is there a pPr change to
 *  apply — and, under tracking, to snapshot into a `<w:pPrChange>`?" Mirrors the
 *  early-return in `applyParagraphOptionsInPlace`. */
export function hasParagraphProperties(options: ParagraphOptions): boolean {
	return Boolean(
		options.style ||
			options.alignment ||
			options.tabs ||
			options.spacing ||
			options.indent,
	);
}

/** CT_PPr child sequence (ECMA-376 §17.3.1.26), the subset we ever emit or
 *  inherit. Word REJECTS a `<w:pPr>` whose children are out of this order —
 *  most commonly `<w:jc>` after the trailing paragraph-mark `<w:rPr>` — with the
 *  "unreadable content / repair" prompt. Any code that splices a child into an
 *  existing pPr must go through `insertPprChildInOrder`, never `push`. */
const PPR_CHILD_ORDER = [
	"w:pStyle",
	"w:keepNext",
	"w:keepLines",
	"w:pageBreakBefore",
	"w:framePr",
	"w:widowControl",
	"w:numPr",
	"w:suppressLineNumbers",
	"w:pBdr",
	"w:shd",
	"w:tabs",
	"w:suppressAutoHyphens",
	"w:bidi",
	"w:adjustRightInd",
	"w:snapToGrid",
	"w:spacing",
	"w:ind",
	"w:contextualSpacing",
	"w:mirrorIndents",
	"w:suppressOverlap",
	"w:jc",
	"w:textDirection",
	"w:textAlignment",
	"w:textboxTightWrap",
	"w:outlineLvl",
	"w:divId",
	"w:cnfStyle",
	"w:rPr",
	"w:sectPr",
	"w:pPrChange",
] as const;

/** Rank a pPr child by its position in CT_PPr. Unknown tags rank just before
 *  `<w:rPr>` so they still land ahead of the paragraph-mark run props (and
 *  sectPr/pPrChange), never after — the only ordering that matters for validity. */
function pprChildRank(tag: string): number {
	const index = PPR_CHILD_ORDER.indexOf(
		tag as (typeof PPR_CHILD_ORDER)[number],
	);
	if (index >= 0) return index;
	return PPR_CHILD_ORDER.indexOf("w:rPr") - 0.5;
}

/** Splice `child` into `pPr.children` at its canonical CT_PPr position: before
 *  the first existing child that ranks after it. Use this instead of `push`
 *  whenever you add to an already-built pPr (which usually ends in the
 *  paragraph-mark `<w:rPr>`); see `PPR_CHILD_ORDER`. */
export function insertPprChildInOrder(pPr: XmlNode, child: XmlNode): void {
	const rank = pprChildRank(child.tag);
	const at = pPr.children.findIndex(
		(existing) => pprChildRank(existing.tag) > rank,
	);
	if (at < 0) pPr.children.push(child);
	else pPr.children.splice(at, 0, child);
}

/** CT_RPr child order (ECMA-376 §17.3.2.28). The same validity rule as pPr:
 *  splice a child into an already-built `<w:rPr>` at its canonical slot or Word
 *  rejects the file. The classic break is adding `<w:sz>` after `<w:rFonts>` when
 *  `<w:b>`/`<w:color>`/`<w:kern>` sit between them (CT_RPr puts sz well after
 *  those) — `insertAfter(rFonts)` would land sz too early. */
const RPR_CHILD_ORDER = [
	"w:rStyle",
	"w:rFonts",
	"w:b",
	"w:bCs",
	"w:i",
	"w:iCs",
	"w:caps",
	"w:smallCaps",
	"w:strike",
	"w:dstrike",
	"w:outline",
	"w:shadow",
	"w:emboss",
	"w:imprint",
	"w:noProof",
	"w:snapToGrid",
	"w:vanish",
	"w:webHidden",
	"w:color",
	"w:spacing",
	"w:w",
	"w:kern",
	"w:position",
	"w:sz",
	"w:szCs",
	"w:highlight",
	"w:u",
	"w:effect",
	"w:bdr",
	"w:shd",
	"w:fitText",
	"w:vertAlign",
	"w:rtl",
	"w:cs",
	"w:em",
	"w:lang",
	"w:eastAsianLayout",
	"w:specVanish",
	"w:oMath",
] as const;

/** Rank an rPr child by CT_RPr position. Unknown tags rank last so an
 *  in-order insert of a known property (rFonts/sz/…) still lands ahead of them. */
function rprChildRank(tag: string): number {
	const index = RPR_CHILD_ORDER.indexOf(
		tag as (typeof RPR_CHILD_ORDER)[number],
	);
	return index >= 0 ? index : RPR_CHILD_ORDER.length;
}

/** Splice `child` into `rPr.children` at its canonical CT_RPr position: before
 *  the first existing child that ranks after it. Use this instead of `push`/
 *  `unshift`/`insertAfter` when adding to an already-built rPr; see
 *  `RPR_CHILD_ORDER`. */
export function insertRprChildInOrder(rPr: XmlNode, child: XmlNode): void {
	const rank = rprChildRank(child.tag);
	const at = rPr.children.findIndex(
		(existing) => rprChildRank(existing.tag) > rank,
	);
	if (at < 0) rPr.children.push(child);
	else rPr.children.splice(at, 0, child);
}

/** A paragraph rendered as a horizontal rule — empty body with a bottom border.
 * Word renders this as a thin line spanning the page width.
 *
 * @public Staged for the S8 markdown walker's `---` thematic breaks; not yet
 * wired into a CLI verb, so it has no internal caller today. */
export function HorizontalRule(): XmlNode {
	return (
		<w.p>
			<w.pPr>
				<w.pBdr>
					<w.bottom w-val="single" w-sz="6" w-space="1" w-color="auto" />
				</w.pBdr>
			</w.pPr>
		</w.p>
	);
}

function TextRunElement({ run }: { run: TextRun }): XmlNode {
	return (
		<w.r>
			<RunProperties run={run} />
			<w.t {...{ "xml:space": "preserve" }}>{run.text}</w.t>
		</w.r>
	);
}

// The fields whose presence means "emit a `<w:rPr>`". Deliberately excludes
// `colorThemeTint`/`colorThemeShade` (modify a `<w:color>`, never standalone)
// and `underlineColor` (rides on `<w:u>`) — including them would emit a stray
// empty `<w:rPr/>` for a run that carries only a dependent modifier.
const FORMATTING_KEYS = [
	"runStyle",
	"font",
	"bold",
	"italic",
	"allCaps",
	"smallCaps",
	"strike",
	"color",
	"colorTheme",
	"sizeHalfPoints",
	"highlight",
	"underline",
	"shade",
	"vertAlign",
] as const satisfies readonly (keyof TextRun)[];

function RunProperties({ run }: { run: TextRun }): NullableXmlNode {
	const isEmpty = FORMATTING_KEYS.every((key) => run[key] == null);
	if (isEmpty) return null;
	// `<w:rPr>` children follow CT_RPr order (ECMA-376 §17.3.2.28):
	// rStyle → rFonts → b → i → caps → smallCaps → strike → color → sz →
	// highlight → u → shd → vertAlign. Kept byte-identical to the markdown-import
	// emitter in `core/markdown/inline.tsx`.
	// Word/LibreOffice tolerate other orderings, but matching the schema keeps
	// validators happy, minimizes the round-trip diff against Word, and keeps
	// this emitter byte-identical to the markdown-import `RunProperties` in
	// `core/markdown/inline.tsx` — the two MUST converge.
	return (
		<w.rPr>
			{run.runStyle && <w.rStyle w-val={run.runStyle} />}
			{run.font && <w.rFonts w-ascii={run.font} w-hAnsi={run.font} />}
			{run.bold && <w.b />}
			{run.italic && <w.i />}
			{run.allCaps && <w.caps />}
			{run.smallCaps && <w.smallCaps />}
			{run.strike && <w.strike />}
			{(run.color || run.colorTheme) && (
				<w.color
					w-val={run.color ?? "auto"}
					w-themeColor={run.colorTheme}
					w-themeTint={run.colorThemeTint}
					w-themeShade={run.colorThemeShade}
				/>
			)}
			{run.sizeHalfPoints !== undefined && (
				<w.sz w-val={String(run.sizeHalfPoints)} />
			)}
			{run.highlight && <w.highlight w-val={run.highlight} />}
			{run.underline && (
				<w.u w-val={run.underline} w-color={run.underlineColor} />
			)}
			{run.shade && <w.shd w-val="clear" w-color="auto" w-fill={run.shade} />}
			{run.vertAlign && <w.vertAlign w-val={run.vertAlign} />}
		</w.rPr>
	);
}

function ParagraphProperties({
	options,
}: {
	options: ParagraphOptions;
}): NullableXmlNode {
	const hasTabs = options.tabs !== undefined && options.tabs.length > 0;
	const spacingAttrs = spacingAttributes(options.spacing);
	const indentAttrs = indentAttributes(options.indent);
	if (
		!options.style &&
		!options.alignment &&
		!options.list &&
		!hasTabs &&
		!spacingAttrs &&
		!indentAttrs
	)
		return null;
	// Schema order (CT_PPrBase, ECMA-376 §17.3.1.26):
	// pStyle → numPr → tabs → spacing → ind → jc.
	return (
		<w.pPr>
			{options.style && <w.pStyle w-val={options.style} />}
			{options.list && (
				<w.numPr>
					<w.ilvl w-val={String(options.list.level)} />
					<w.numId w-val={String(options.list.numId)} />
				</w.numPr>
			)}
			{hasTabs && (
				<w.tabs>
					{options.tabs?.map((stop) => (
						<w.tab w-val={stop.align} w-pos={String(Math.round(stop.pos))} />
					))}
				</w.tabs>
			)}
			{spacingAttrs && <w.spacing {...spacingAttrs} />}
			{indentAttrs && <w.ind {...indentAttrs} />}
			{options.alignment && <w.jc w-val={options.alignment} />}
		</w.pPr>
	);
}

/** Build the `<w:spacing>` attribute map from a `ParagraphSpacing`, or null if it
 *  sets nothing. `before`/`after` are emitted in twips; `line` carries its
 *  `lineRule` (defaulting to `auto`, where `line` is 240ths of a line). */
export function spacingAttributes(
	spacing: ParagraphSpacing | undefined,
): Record<string, string> | null {
	if (!spacing) return null;
	const attrs: Record<string, string> = {};
	if (spacing.before !== undefined)
		attrs["w:before"] = String(Math.round(spacing.before));
	if (spacing.after !== undefined)
		attrs["w:after"] = String(Math.round(spacing.after));
	if (spacing.line !== undefined) {
		attrs["w:line"] = String(Math.round(spacing.line));
		attrs["w:lineRule"] = spacing.lineRule ?? "auto";
	}
	return Object.keys(attrs).length > 0 ? attrs : null;
}

/** Build the `<w:ind>` attribute map from a `ParagraphIndent`, or null if empty.
 *  All values are twips. `firstLine` and `hanging` are mutually exclusive. */
export function indentAttributes(
	indent: ParagraphIndent | undefined,
): Record<string, string> | null {
	if (!indent) return null;
	const attrs: Record<string, string> = {};
	if (indent.left !== undefined)
		attrs["w:left"] = String(Math.round(indent.left));
	if (indent.right !== undefined)
		attrs["w:right"] = String(Math.round(indent.right));
	if (indent.firstLine !== undefined)
		attrs["w:firstLine"] = String(Math.round(indent.firstLine));
	if (indent.hanging !== undefined)
		attrs["w:hanging"] = String(Math.round(indent.hanging));
	return Object.keys(attrs).length > 0 ? attrs : null;
}
