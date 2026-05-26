import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

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
