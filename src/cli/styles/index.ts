import { type Block, baselineCatalog, Fonts, iterateBlocks } from "@core";
import type { XmlNode } from "@core/parser";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx styles — list the styles available to apply, describe one, or set the document font

Usage:
  docx styles FILE [--used] [--at STYLEID] [--json]
  docx styles --catalog [--json]
  docx styles set-default-font FILE "Font Name" [--size N] [--all]   # set the document-wide font

The style catalog lives in word/styles.xml, not in the document body — so
unlike everything else, you can't see it by reading the doc. Use this to learn
what \`--style NAME\` values exist (for \`insert --style\` / \`edit --style\`) and
what a style actually looks like before applying it.

Options:
  --at STYLEID   Describe one style (id, type, name, basedOn, key formatting)
  --used         List only the styles actually applied somewhere in the body
                 (paragraph styles + character/run styles)
  --catalog      List the built-in styles docx-cli can apply on demand — every
                 \`--style NAME\` value that \`insert\`/\`edit\` will auto-provision
                 (Title, Subtitle, Heading1–9, Quote, IntenseQuote, Code, …)
                 even when the doc doesn't contain them yet. No FILE needed.
  --json         Structured output (a JSON array for the list; an object for --at)
  -h, --help     Show this help

Output:
  Default: a text table (id, type, name, basedOn), one style per line. --at
  prints a style's detail. --json emits the structured form. Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx styles report.docx                 # styles defined in this doc
  docx styles report.docx --used          # only styles the doc uses
  docx styles --catalog                    # built-ins you can apply via --style
  docx styles report.docx --at Caption    # what does Caption look like?
  docx styles report.docx --json | jq '.[].id'
  docx styles set-default-font report.docx "Times New Roman"   # whole-doc font
  docx styles set-default-font report.docx "Georgia" --all     # incl. explicit fonts

See \`docx styles set-default-font --help\` for the font-setting details.
`;

export async function run(args: string[]): Promise<number> {
	// `styles` is read-by-default; `set-default-font` is its one mutating subverb
	// (keeps the read command from going dual-natured, and inherits --dry-run /
	// -o / the ack confirmation as a normal mutator).
	if (args[0] === "set-default-font") {
		return runSetDefaultFont(args.slice(1));
	}

	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			used: { type: "boolean" },
			catalog: { type: "boolean" },
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

const FONT_HELP = `docx styles set-default-font — set the document-wide default font

Usage:
  docx styles set-default-font FILE "Font Name" [--size N] [--all] [options]

A document font lives in TWO places at once — word/styles.xml (<w:docDefaults>)
and the theme font scheme (word/theme/theme1.xml, major + minor) — and setting
only one silently loses to the other. This sets both, so body text AND
theme-following headings adopt the font. Styles/runs that pin their OWN font
(e.g. a code block's monospace, a deliberately-Arial run) are preserved; pass
--all to repoint those too.

Options:
  --size N           Also set the default font size, in points (e.g. 12).
  --all              Repoint EVERY explicit font — styles, body runs, and notes —
                     onto FONT too, for a guaranteed-uniform document (overrides
                     even code monospace and per-run font choices).
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write the file
  -v, --verbose      Print the full success ack JSON
  -h, --help         Show this help

Examples:
  docx styles set-default-font report.docx "Times New Roman"
  docx styles set-default-font report.docx "Calibri" --size 11
  docx styles set-default-font report.docx "Georgia" --all
`;

async function runSetDefaultFont(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{ size: { type: "string" }, all: { type: "boolean" }, ...SAVE_FLAGS },
		FONT_HELP,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(FONT_HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", FONT_HELP);
	const fontName = parsed.positionals[1];
	if (!fontName) {
		return fail(
			"USAGE",
			'Missing FONT name (e.g. "Times New Roman")',
			FONT_HELP,
		);
	}

	let sizeHalfPoints: number | undefined;
	const sizeRaw = parsed.values.size as string | undefined;
	if (sizeRaw !== undefined) {
		// Strict decimal only — `Number.parseFloat` would silently accept trailing
		// garbage ("11pt"→11, "1e3"→1000, "12abc"→12) and write a wrong size.
		if (
			!/^\s*\d+(\.\d+)?\s*$/.test(sizeRaw) ||
			Number.parseFloat(sizeRaw) <= 0
		) {
			return fail(
				"USAGE",
				`--size must be a positive number of points, got "${sizeRaw}"`,
			);
		}
		sizeHalfPoints = Math.round(Number.parseFloat(sizeRaw) * 2);
	}

	const all = Boolean(parsed.values.all);
	const outputPath = parsed.values.output as string | undefined;
	const dryRun = Boolean(parsed.values["dry-run"]);

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const result = await new Fonts(document).setDefault(fontName, {
		sizeHalfPoints,
		all,
	});

	if (dryRun) {
		await respond({
			operation: "styles.set-default-font",
			dryRun: true,
			path: filePath,
			font: fontName,
			...(sizeHalfPoints !== undefined ? { sizePt: sizeHalfPoints / 2 } : {}),
			all,
			themeUpdated: result.themeUpdated,
			// Preview the decision-relevant fact: which styles stay off-font (or,
			// under --all, how many fonts would be repointed) — same as the real run.
			...(all
				? { repointed: result.repointed }
				: { explicitStyles: result.explicitStyles }),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	await document.save(outputPath);

	// Tell the agent what kept its own font (so "why is the heading still blue
	// Calibri?" has an answer) and how to force it.
	const count = result.explicitStyles.length;
	const leftover =
		!all && count > 0
			? `${count === 1 ? "1 style keeps" : `${count} styles keep`} their own font (${formatStyleList(result.explicitStyles)}); pass --all to override them too.`
			: undefined;

	await respondAck(
		{
			ok: true,
			operation: "styles.set-default-font",
			path: outputPath ?? filePath,
			font: fontName,
			...(sizeHalfPoints !== undefined ? { sizePt: sizeHalfPoints / 2 } : {}),
			themeUpdated: result.themeUpdated,
			all,
			...(all
				? { repointed: result.repointed }
				: { explicitStyles: result.explicitStyles }),
		},
		leftover,
	);
	return EXIT.OK;
}

function formatStyleList(ids: string[]): string {
	const shown = ids.slice(0, 5).join(", ");
	return ids.length > 5 ? `${shown}, …` : shown;
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
