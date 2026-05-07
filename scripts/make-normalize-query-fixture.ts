import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/normalize-query.docx — exercises every category of
 * `find` query normalization (smart quotes, em-dashes, balanced markdown
 * emphasis stripping) plus the conservative-non-stripping case for an
 * unmatched single asterisk used as multiplication.
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
 * Built by dogfooding the CLI; doubles as an end-to-end smoke test for
 * `create` and `insert` with smart-quote / em-dash text.
 */

const root = resolve(import.meta.dir, "..");
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

const verifyJson = await cli("read", out);
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
