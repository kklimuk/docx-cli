import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/normalize-query.docx — exercises every category of
 * `find` query normalization (smart quotes, em-dashes, balanced markdown
 * emphasis stripping) plus the conservative-non-stripping case for an
 * unmatched single asterisk used as multiplication, plus the editor-style
 * cross-paragraph `replace` (merge and split).
 *
 *   p0: 'The plan: “hello” world—ready to ship. The figure: 5 * 3 = 15.'
 *       — smart curly quotes around "hello"
 *       — em-dash inside "world—ready"
 *       — bare " * " between digits to verify the markdown-stripper does
 *         NOT collapse "5 * 3" to "5  3"
 *
 *   p1: 'plan: "hello" today.'
 *       — straight ASCII quotes around hello, used to verify the inverse
 *         direction of smart-quote canonicalization (smart-quote query
 *         matches straight-quote document text).
 *
 *   p2: 'gamma one two delta' — the RESULT of a cross-paragraph MERGE:
 *       authored as two paragraphs ("gamma one" / "two delta") and joined by
 *       `replace "one\ntwo" "one two"`.
 *
 *   p3: 'epsilon end' — the RESULT of a paragraph SPLIT: a "\n" in the
 *       replacement (`replace "delta" "delta\nepsilon end"`) minted this
 *       paragraph out of p2's tail.
 *
 * Built by dogfooding the CLI; doubles as an end-to-end smoke test for
 * `create` / `insert` with smart-quote / em-dash text and for the
 * cross-paragraph replace path, and the LibreOffice round-trip proves the
 * merged/split paragraph XML renders.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/normalize-query.docx");
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
	"Normalize-query fixture",
	"--author",
	"docx-cli",
	"--text",
	"The plan: “hello” world—ready to ship. The figure: 5 * 3 = 15.",
	"--force",
);

await cli("insert", out, "--after", "p0", "--text", 'plan: "hello" today.');

// Cross-paragraph replace dogfood: author two paragraphs, MERGE them with a
// "\n"-bearing pattern, then SPLIT the tail back off with a "\n"-bearing
// replacement. Final layout: p2 'gamma one two delta', p3 'epsilon end'.
await cli("insert", out, "--after", "p1", "--text", "gamma one");
await cli("insert", out, "--after", "p2", "--text", "two delta");
await cli("replace", out, "one\ntwo", "one two");
await cli("replace", out, "delta", "delta\nepsilon end");

const verifyJson = await cli("read", out, "--ast");
const doc = JSON.parse(verifyJson) as {
	blocks: Array<{ id: string; type: string; runs?: Array<{ text?: string }> }>;
};
const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
console.log(`Wrote ${out}`);
console.log("Paragraphs:");
for (const paragraph of paragraphs) {
	const text = (paragraph.runs ?? []).map((run) => run.text ?? "").join("");
	console.log(`  ${paragraph.id}: "${text}"`);
}
