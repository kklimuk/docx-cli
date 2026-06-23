import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/word-formatted.docx — paragraphs with mixed
 * bold/italic/color runs at known token positions, used to verify B3's
 * `edit --text` formatting preservation across word-level edits.
 *
 *   p0: "The quick brown fox jumps over the lazy dog." (plain)
 *   p1: "The MESSENGER is fatally flawed."
 *       — "MESSENGER" bold (purple)
 *       — "fatally flawed" italic
 *       — surrounding text plain
 *   p2: "Bold then italic then plain."
 *       — "Bold" bold
 *       — "italic" italic
 *       — surrounding text plain
 *   p3: "Rating: The messenger is fatally flawed."
 *       — every run italic, but split mid-word at run boundaries
 *         ("Rating: The me" + "ssenger is fatally f" + "lawed.")
 *       — exercises the regression where Word-produced run splits inside
 *         a word would shred tokens during diff alignment, producing
 *         spurious del+ins pairs for words that are actually unchanged
 *   p4: every run-level rPr property the CLI emits (direct + theme color,
 *       highlight, shading fill, underline plain/styled/colored, super/
 *       subscript, small/all caps, custom font, custom size). Absorbed from
 *       the former run-formatting.docx fixture; appended last so p0-p3 stay
 *       byte-stable. This is the run-formatting LibreOffice round-trip surface.
 *
 * Authoring channel — dogfoods the `edit` SET-run-formatting verbs (the inverse
 * of `--clear`): p1 and p2 are authored as PLAIN text via `create`/`insert
 * --text`, then their formatting is applied in place with
 * `edit --at <span> --bold/--italic/--color`. (See tests/fixtures/setup/CLAUDE.md
 * — re-author with the CLI once the verbs exist.) Two paragraphs stay on `--runs`
 * because the SET verbs can't reproduce their shapes:
 *   - p3's identical-rPr mid-word splits (a SET would coalesce them to one run).
 *   - p4, the emitter's full-rPr round-trip surface: it needs a theme color (no
 *     `--color` theme token) and a wavyDouble+colored underline (`--underline`
 *     sets `single` only); and authoring its OTHER runs as plain-then-`edit`
 *     wouldn't work either — `insert` blends PLAIN inserted runs into the anchor
 *     paragraph's formatting (here p3 is all-italic), so plain seed runs would
 *     inherit italic. Every run carrying its own rPr (the `--runs` form) is the
 *     only shape that keeps p4 italic-free.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/word-formatted.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

/** Apply run formatting to a span in place via the SET verbs. */
async function setFormat(span: string, ...flags: string[]): Promise<void> {
	await cli("edit", out, "--at", span, ...flags);
}

mkdirSync(dirname(out), { recursive: true });

// p0: plain.
await cli(
	"create",
	out,
	"--title",
	"Word-formatted fixture",
	"--author",
	"docx-cli",
	"--text",
	"The quick brown fox jumps over the lazy dog.",
	"--force",
);

// p1: plain text, then SET MESSENGER bold+purple and "fatally flawed" italic.
//     "The " 0-4 | "MESSENGER" 4-13 | " is " 13-17 | "fatally flawed" 17-31 | "." 31-32
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--text",
	"The MESSENGER is fatally flawed.",
);
await setFormat("p1:4-13", "--bold", "--color", "800080");
await setFormat("p1:17-31", "--italic");

// p2: plain text, then SET "Bold" bold and "italic" italic.
//     "Bold" 0-4 | " then " 4-10 | "italic" 10-16 | " then plain." 16-28
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--text",
	"Bold then italic then plain.",
);
await setFormat("p2:0-4", "--bold");
await setFormat("p2:10-16", "--italic");

// p3: every run italic, but the splits land mid-word. This is the shape Word
// produces after iterative editing — adjacent runs with identical rPr that just
// happen to break across letters of "messenger" and "flawed". A SET edit would
// coalesce these into one run, so this shape can only be authored via --runs.
// The B3 token extractor must concatenate before tokenizing or it will see
// ["Rating:", " ", "The", " ", "me"] from run 0 and ["ssenger", ...] from run 1,
// never matching new text's clean ["messenger"] token.
const p3Runs = JSON.stringify([
	{ type: "text", text: "Rating: The me", italic: true },
	{ type: "text", text: "ssenger is fatally f", italic: true },
	{ type: "text", text: "lawed.", italic: true },
]);
await cli("insert", out, "--after", "p2", "--runs", p3Runs);

// p4: every run-level rPr property the CLI emits through core/blocks.tsx —
// direct + theme color, named highlight, shading fill, underline (plain +
// styled + colored), super/subscript, small/all caps, a custom font, a custom
// size. Each run carries its own rPr (see the docstring for why plain-then-edit
// would inherit p3's italic) — this is the emitter's full round-trip surface.
// Appended last so the p0-p3 locators the edit tests hard-pin do not shift.
const p4Runs = JSON.stringify([
	{ type: "text", text: "color ", color: "FF0000" },
	{
		type: "text",
		text: "theme ",
		color: "4472C4",
		colorTheme: "accent1",
		colorThemeTint: "99",
	},
	{ type: "text", text: "highlight ", highlight: "yellow" },
	{ type: "text", text: "shade ", shade: "FFE599" },
	{ type: "text", text: "underline ", underline: "single" },
	{
		type: "text",
		text: "wavy ",
		underline: "wavyDouble",
		underlineColor: "FF0000",
	},
	{ type: "text", text: "x", sizeHalfPoints: 22 },
	{ type: "text", text: "2", vertAlign: "superscript" },
	{ type: "text", text: " plus H", sizeHalfPoints: 22 },
	{ type: "text", text: "2", vertAlign: "subscript" },
	{ type: "text", text: "O ", sizeHalfPoints: 22 },
	{ type: "text", text: "smallcaps ", smallCaps: true },
	{ type: "text", text: "allcaps ", allCaps: true },
	{ type: "text", text: "courier ", font: "Courier New" },
	{ type: "text", text: "big", sizeHalfPoints: 36 },
]);
await cli("insert", out, "--after", "p3", "--runs", p4Runs);

const verifyJson = await cli("read", out, "--ast");
const doc = JSON.parse(verifyJson) as {
	blocks: Array<{
		id: string;
		type: string;
		runs?: Array<{
			type: string;
			text?: string;
			bold?: boolean;
			italic?: boolean;
			color?: string;
			highlight?: string;
			underline?: string;
			font?: string;
			sizeHalfPoints?: number;
			vertAlign?: string;
			smallCaps?: boolean;
			allCaps?: boolean;
		}>;
	}>;
};
const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
console.log(`Wrote ${out}`);
for (const paragraph of paragraphs) {
	const annotated = (paragraph.runs ?? [])
		.map((run) => {
			const flags: string[] = [];
			if (run.bold) flags.push("b");
			if (run.italic) flags.push("i");
			if (run.underline) flags.push(`u:${run.underline}`);
			if (run.color) flags.push(`#${run.color}`);
			if (run.highlight) flags.push(`hl:${run.highlight}`);
			if (run.font) flags.push(run.font);
			if (run.sizeHalfPoints) flags.push(`${run.sizeHalfPoints / 2}pt`);
			if (run.vertAlign) flags.push(run.vertAlign);
			if (run.smallCaps) flags.push("smallCaps");
			if (run.allCaps) flags.push("allCaps");
			const tag = flags.length > 0 ? ` [${flags.join(",")}]` : "";
			return `"${run.text ?? ""}"${tag}`;
		})
		.join(" + ");
	console.log(`  ${paragraph.id}: ${annotated}`);
}
