import { type Block, iterateBlocks } from "@core";
import type { XmlNode } from "@core/parser";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx styles — list the styles available to apply, or describe one

Usage:
  docx styles FILE [--used] [--at STYLEID] [--json]

The style catalog lives in word/styles.xml, not in the document body — so
unlike everything else, you can't see it by reading the doc. Use this to learn
what \`--style NAME\` values exist (for \`insert --style\` / \`edit --style\`) and
what a style actually looks like before applying it.

Options:
  --at STYLEID   Describe one style (id, type, name, basedOn, key formatting)
  --used         List only the styles actually applied somewhere in the body
                 (paragraph styles + character/run styles)
  --json         Structured output (a JSON array for the list; an object for --at)
  -h, --help     Show this help

Output:
  Default: a text table (id, type, name, basedOn), one style per line. --at
  prints a style's detail. --json emits the structured form. Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx styles report.docx                 # full catalog
  docx styles report.docx --used          # only styles the doc uses
  docx styles report.docx --at Caption    # what does Caption look like?
  docx styles report.docx --json | jq '.[].id'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			used: { type: "boolean" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

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

type StyleDetail = StyleMeta & {
	font?: string;
	sizePt?: number;
	color?: string;
	bold?: boolean;
	italic?: boolean;
	alignment?: string;
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

/** A style's own (directly-declared) key formatting — not resolved up the
 * `basedOn` chain (an agent can follow `basedOn` via the catalog). Only the
 * properties present on this `<w:style>` are included. */
function describeStyle(node: XmlNode): StyleDetail {
	const detail: StyleDetail = styleMeta(node);
	const rPr = node.findChild("w:rPr");
	const font = rPr?.findChild("w:rFonts")?.getAttribute("w:ascii");
	if (font) detail.font = font;
	const sizeHalf = rPr?.findChild("w:sz")?.getAttribute("w:val");
	if (sizeHalf) {
		const parsed = Number.parseInt(sizeHalf, 10);
		if (Number.isFinite(parsed)) detail.sizePt = parsed / 2;
	}
	const color = rPr?.findChild("w:color")?.getAttribute("w:val");
	if (color && color !== "auto") detail.color = color.toUpperCase();
	const boldEl = rPr?.findChild("w:b");
	if (boldEl) detail.bold = toggleOn(boldEl);
	const italicEl = rPr?.findChild("w:i");
	if (italicEl) detail.italic = toggleOn(italicEl);
	const jc = node.findChild("w:pPr")?.findChild("w:jc")?.getAttribute("w:val");
	if (jc) detail.alignment = jc;
	return detail;
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
		if (value !== undefined) lines.push(`  ${label.padEnd(9)} ${value}`);
	};
	add("name:", detail.name);
	add("basedOn:", detail.basedOn);
	add("font:", detail.font);
	add("size:", detail.sizePt !== undefined ? `${detail.sizePt}pt` : undefined);
	add("color:", detail.color);
	add("bold:", detail.bold);
	add("italic:", detail.italic);
	add("align:", detail.alignment);
	return `${lines.join("\n")}\n`;
}
