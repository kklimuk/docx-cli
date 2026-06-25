import { type Block, baselineCatalog, iterateBlocks } from "@core";
import type { XmlNode } from "@core/parser";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

/** `docx styles` (the read default): list the catalog, describe one style, or list
 *  the built-in catalog. The mutating subverbs (`set`/`create`/`set-default-font`)
 *  live in sibling files; the dispatcher in `index.ts` routes to them. */
export async function runStylesRead(
	args: string[],
	help: string,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			used: { type: "boolean" },
			catalog: { type: "boolean" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		help,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	// `--catalog` lists the built-in styles docx-cli can provision on demand
	// (Title, Heading1–9, Quote, …). It's a static catalog — no document needed —
	// so an agent can discover valid `--style` values even with no FILE to hand.
	if (parsed.values.catalog) {
		const metas = baselineCatalog().map(styleMeta);
		if (parsed.values.json) {
			await respond(metas);
			return EXIT.OK;
		}
		await writeStdout(formatList(metas));
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;
	const styles = document.styles;
	const json = Boolean(parsed.values.json);

	const at = parsed.values.at as string | undefined;
	if (at !== undefined) {
		const node = styles?.getStyle(at);
		if (!node) {
			return fail(
				"BLOCK_NOT_FOUND",
				`No style with id "${at}"`,
				"Run `docx styles FILE` to list available style ids.",
			);
		}
		const detail = describeStyle(node);
		if (json) {
			await respond(detail);
			return EXIT.OK;
		}
		await writeStdout(formatDetail(detail));
		return EXIT.OK;
	}

	const ids = styles?.listStyleIds() ?? [];
	let metas = ids
		.map((id) => styles?.getStyle(id))
		.filter((node): node is XmlNode => node !== undefined)
		.map(styleMeta);

	if (parsed.values.used) {
		const used = appliedStyleIds(document);
		metas = metas.filter((meta) => used.has(meta.id));
	}

	if (json) {
		await respond(metas);
		return EXIT.OK;
	}
	await writeStdout(formatList(metas));
	return EXIT.OK;
}

type StyleMeta = {
	id: string;
	type: string;
	name?: string;
	basedOn?: string;
};

/** A style's own (directly-declared) key formatting — every property `styles set`
 *  can write, so the write-read loop holds: after `styles set --at X --bold
 *  --space-before 12`, `styles --at X` shows them. Not resolved up the `basedOn`
 *  chain (an agent can follow `basedOn`). Only properties present on THIS
 *  `<w:style>` are included. */
type StyleDetail = StyleMeta & {
	next?: string;
	font?: string;
	sizePt?: number;
	color?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: string;
	strike?: boolean;
	caps?: boolean;
	smallCaps?: boolean;
	highlight?: string;
	shade?: string;
	vertAlign?: string;
	alignment?: string;
	spaceBeforePt?: number;
	spaceAfterPt?: number;
	lineSpacing?: string;
	indentLeftIn?: number;
	indentRightIn?: number;
	firstLineIn?: number;
	hangingIn?: number;
};

function styleMeta(node: XmlNode): StyleMeta {
	const meta: StyleMeta = {
		id: node.getAttribute("w:styleId") ?? "",
		type: node.getAttribute("w:type") ?? "",
	};
	const name = node.findChild("w:name")?.getAttribute("w:val");
	if (name) meta.name = name;
	const basedOn = node.findChild("w:basedOn")?.getAttribute("w:val");
	if (basedOn) meta.basedOn = basedOn;
	return meta;
}

const TWIPS_PER_POINT = 20;
const TWIPS_PER_INCH = 1440;

function describeStyle(node: XmlNode): StyleDetail {
	const detail: StyleDetail = styleMeta(node);
	// `next` (the following-paragraph style) is settable via `styles set/create`, so
	// surface it here too — keeping the write-read loop whole. (It rides the detail,
	// not the list, since it's only relevant when inspecting one style.)
	const next = node.findChild("w:next")?.getAttribute("w:val");
	if (next) detail.next = next;
	const rPr = node.findChild("w:rPr");
	if (rPr) describeRunProperties(rPr, detail);
	const pPr = node.findChild("w:pPr");
	if (pPr) describeParagraphProperties(pPr, detail);
	return detail;
}

function describeRunProperties(rPr: XmlNode, detail: StyleDetail): void {
	const font = rPr.findChild("w:rFonts")?.getAttribute("w:ascii");
	if (font) detail.font = font;
	const sizeHalf = rPr.findChild("w:sz")?.getAttribute("w:val");
	if (sizeHalf) {
		const parsed = Number.parseInt(sizeHalf, 10);
		if (Number.isFinite(parsed)) detail.sizePt = parsed / 2;
	}
	const color = rPr.findChild("w:color")?.getAttribute("w:val");
	if (color && color !== "auto") detail.color = color.toUpperCase();
	const bold = rPr.findChild("w:b");
	if (bold) detail.bold = toggleOn(bold);
	const italic = rPr.findChild("w:i");
	if (italic) detail.italic = toggleOn(italic);
	const strike = rPr.findChild("w:strike");
	if (strike) detail.strike = toggleOn(strike);
	const caps = rPr.findChild("w:caps");
	if (caps) detail.caps = toggleOn(caps);
	const smallCaps = rPr.findChild("w:smallCaps");
	if (smallCaps) detail.smallCaps = toggleOn(smallCaps);
	const underline = rPr.findChild("w:u")?.getAttribute("w:val");
	if (underline && underline !== "none") detail.underline = underline;
	const highlight = rPr.findChild("w:highlight")?.getAttribute("w:val");
	if (highlight) detail.highlight = highlight;
	const shade = rPr.findChild("w:shd")?.getAttribute("w:fill");
	if (shade && shade !== "auto") detail.shade = shade.toUpperCase();
	const vertAlign = rPr.findChild("w:vertAlign")?.getAttribute("w:val");
	if (vertAlign && vertAlign !== "baseline") detail.vertAlign = vertAlign;
}

function describeParagraphProperties(pPr: XmlNode, detail: StyleDetail): void {
	const jc = pPr.findChild("w:jc")?.getAttribute("w:val");
	if (jc) detail.alignment = jc;
	const spacing = pPr.findChild("w:spacing");
	if (spacing) {
		const before = numberAttr(spacing, "w:before");
		if (before !== undefined) detail.spaceBeforePt = before / TWIPS_PER_POINT;
		const after = numberAttr(spacing, "w:after");
		if (after !== undefined) detail.spaceAfterPt = after / TWIPS_PER_POINT;
		const line = numberAttr(spacing, "w:line");
		if (line !== undefined) {
			const rule = spacing.getAttribute("w:lineRule");
			detail.lineSpacing =
				rule === "exact" || rule === "atLeast"
					? `${line / TWIPS_PER_POINT}pt ${rule}`
					: String(line / 240);
		}
	}
	const ind = pPr.findChild("w:ind");
	if (ind) {
		const left = numberAttr(ind, "w:left");
		if (left !== undefined) detail.indentLeftIn = left / TWIPS_PER_INCH;
		const right = numberAttr(ind, "w:right");
		if (right !== undefined) detail.indentRightIn = right / TWIPS_PER_INCH;
		const firstLine = numberAttr(ind, "w:firstLine");
		if (firstLine !== undefined)
			detail.firstLineIn = firstLine / TWIPS_PER_INCH;
		const hanging = numberAttr(ind, "w:hanging");
		if (hanging !== undefined) detail.hangingIn = hanging / TWIPS_PER_INCH;
	}
}

function numberAttr(node: XmlNode, attr: string): number | undefined {
	const raw = node.getAttribute(attr);
	if (raw === undefined) return undefined;
	const parsed = Number.parseInt(raw, 10);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** Resolve an OOXML toggle property (`<w:b>`, `<w:i>`) to its boolean value.
 * A bare element (or `w:val="1"/"true"/"on"`) is ON; `w:val="0"/"false"/"off"`
 * is an explicit OFF (e.g. a style that un-bolds a bold parent). Presence alone
 * is NOT "on" — reporting `bold:true` for `<w:b w:val="0"/>` would lie. */
function toggleOn(el: XmlNode): boolean {
	const val = el.getAttribute("w:val");
	if (!val) return true;
	return !/^(0|false|off)$/i.test(val);
}

/** The set of style ids applied anywhere in the body — paragraph styles
 * (`pStyle`), character/run styles (`rStyle`), and table styles (`tblStyle`) —
 * so `--used` shows only what the document actually relies on (a much shorter
 * list than the full catalog). `iterateBlocks` descends into table cells, so
 * styles applied only inside a cell count too. */
function appliedStyleIds(document: { body: { blocks: Block[] } }): Set<string> {
	const used = new Set<string>();
	for (const block of iterateBlocks(document.body.blocks)) {
		if (block.type === "table") {
			if (block.style) used.add(block.style);
			continue;
		}
		if (block.type !== "paragraph") continue;
		if (block.style) used.add(block.style);
		for (const run of block.runs) {
			if (run.type === "text" && run.runStyle) used.add(run.runStyle);
		}
	}
	return used;
}

function formatList(metas: StyleMeta[]): string {
	if (metas.length === 0) return "(no styles)\n";
	const idWidth = Math.max(...metas.map((meta) => meta.id.length), 2);
	const typeWidth = Math.max(...metas.map((meta) => meta.type.length), 4);
	const lines = metas.map((meta) => {
		const id = meta.id.padEnd(idWidth);
		const type = meta.type.padEnd(typeWidth);
		const name = meta.name ?? "";
		const basedOn = meta.basedOn ? `  ← ${meta.basedOn}` : "";
		return `${id}  ${type}  ${name}${basedOn}`.trimEnd();
	});
	return `${lines.join("\n")}\n`;
}

function formatDetail(detail: StyleDetail): string {
	const lines = [`${detail.id} (${detail.type})`];
	const add = (label: string, value: string | number | boolean | undefined) => {
		if (value !== undefined) lines.push(`  ${label.padEnd(12)} ${value}`);
	};
	add("name:", detail.name);
	add("basedOn:", detail.basedOn);
	add("next:", detail.next);
	add("font:", detail.font);
	add("size:", detail.sizePt !== undefined ? `${detail.sizePt}pt` : undefined);
	add("color:", detail.color);
	add("bold:", detail.bold);
	add("italic:", detail.italic);
	add("underline:", detail.underline);
	add("strike:", detail.strike);
	add("caps:", detail.caps);
	add("smallCaps:", detail.smallCaps);
	add("highlight:", detail.highlight);
	add("shade:", detail.shade);
	add("vertAlign:", detail.vertAlign);
	add("align:", detail.alignment);
	add(
		"space-before:",
		detail.spaceBeforePt !== undefined
			? `${detail.spaceBeforePt}pt`
			: undefined,
	);
	add(
		"space-after:",
		detail.spaceAfterPt !== undefined ? `${detail.spaceAfterPt}pt` : undefined,
	);
	add("line-spacing:", detail.lineSpacing);
	add(
		"indent-left:",
		detail.indentLeftIn !== undefined ? `${detail.indentLeftIn}in` : undefined,
	);
	add(
		"indent-right:",
		detail.indentRightIn !== undefined
			? `${detail.indentRightIn}in`
			: undefined,
	);
	add(
		"first-line:",
		detail.firstLineIn !== undefined ? `${detail.firstLineIn}in` : undefined,
	);
	add(
		"hanging:",
		detail.hangingIn !== undefined ? `${detail.hangingIn}in` : undefined,
	);
	return `${lines.join("\n")}\n`;
}
