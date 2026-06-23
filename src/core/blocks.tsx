import type { Run, TextRun } from "./ast/types";
import { w } from "./jsx";
import { type NullableXmlNode, XmlNode } from "./parser";
import { TaskCheckbox } from "./task-list";

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
			<ParagraphProperties options={{ style, alignment, list, tabs }} />
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
};

/** One `<w:tab>` entry: `align` is the `w:val` (left/right/center/…), `pos` the
 *  `w:pos` in twips (absolute from the left text margin). */
export type TabStop = { align: string; pos: number };

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
	if (!options.style && !options.alignment && !options.tabs) return;
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
	if (!options.style && !options.alignment && !options.list && !hasTabs)
		return null;
	// Schema order (CT_PPrBase, ECMA-376 §17.3.1.26): pStyle → numPr → tabs → jc.
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
			{options.alignment && <w.jc w-val={options.alignment} />}
		</w.pPr>
	);
}
