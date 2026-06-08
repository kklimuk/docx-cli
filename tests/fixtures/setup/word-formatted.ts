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
 * Built by dogfooding the CLI's `--runs` JSON path; doubles as a smoke
 * test for `create` + `insert --runs`.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/word-formatted.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

mkdirSync(dirname(out), { recursive: true });

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

const p1Runs = JSON.stringify([
	{ type: "text", text: "The " },
	{ type: "text", text: "MESSENGER", bold: true, color: "800080" },
	{ type: "text", text: " is " },
	{ type: "text", text: "fatally flawed", italic: true },
	{ type: "text", text: "." },
]);
await cli("insert", out, "--after", "p0", "--runs", p1Runs);

const p2Runs = JSON.stringify([
	{ type: "text", text: "Bold", bold: true },
	{ type: "text", text: " then " },
	{ type: "text", text: "italic", italic: true },
	{ type: "text", text: " then plain." },
]);
await cli("insert", out, "--after", "p1", "--runs", p2Runs);

// p3: every run italic, but the splits land mid-word. This is the shape
// Word produces after iterative editing — adjacent runs with identical
// rPr that just happen to break across letters of "messenger" and
// "flawed". The B3 token extractor must concatenate before tokenizing or
// it will see ["Rating:", " ", "The", " ", "me"] from run 0 and
// ["ssenger", ...] from run 1, never matching new text's clean
// ["messenger"] token.
const p3Runs = JSON.stringify([
	{ type: "text", text: "Rating: The me", italic: true },
	{ type: "text", text: "ssenger is fatally f", italic: true },
	{ type: "text", text: "lawed.", italic: true },
]);
await cli("insert", out, "--after", "p2", "--runs", p3Runs);

// p4: every run-level rPr property the CLI emits through core/blocks.tsx —
// direct + theme color, named highlight, shading fill, underline (plain +
// styled + colored), super/subscript, small/all caps, a custom font, a custom
// size. Absorbed verbatim from the former run-formatting.docx fixture so this
// one doc carries both the edit-preservation layout (p0-p3) and the full
// run-formatting round-trip surface. Appended as a trailing paragraph so the
// p0-p3 locators edit-span/edit-formatting hard-pin do not shift.
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
			if (run.color) flags.push(`#${run.color}`);
			const tag = flags.length > 0 ? ` [${flags.join(",")}]` : "";
			return `"${run.text ?? ""}"${tag}`;
		})
		.join(" + ");
	console.log(`  ${paragraph.id}: ${annotated}`);
}
