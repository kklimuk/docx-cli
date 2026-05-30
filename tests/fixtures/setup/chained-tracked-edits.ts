import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/chained-tracked-edits.docx — a baseline doc with
 * track-changes ENABLED but NO tracked changes yet, used to verify that
 * chained replaces under tracking keep their offsets stable in accepted
 * view (the agent-feedback bug repro from issue #N).
 *
 *   p0: "Cost of living, anti-price-gouging, and housing reform."
 *       — exercises the chained-replace scenario where two consecutive
 *         `replace` calls land in the same paragraph. In the buggy
 *         version, the second match's offset was computed against a
 *         haystack that included the first replace's <w:ins>, slicing
 *         the splice mid-word inside the inserted run.
 *
 *   p1: "Old plan: ship Tuesday."
 *       — exercises the find-views contrast: under default (accepted)
 *         view, "Old" disappears from the haystack after a replace
 *         turns it into "New"; under --current view, "Old" still
 *         shows up because the <w:del> wrapper text is still on disk.
 *
 * Built by dogfooding the CLI; doubles as an end-to-end smoke test for
 * `create`, `insert`, and `track-changes on`.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/chained-tracked-edits.docx");
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
	"Chained tracked edits fixture",
	"--author",
	"docx-cli",
	"--text",
	"Cost of living, anti-price-gouging, and housing reform.",
	"--force",
);

await cli("insert", out, "--after", "p0", "--text", "Old plan: ship Tuesday.");

await cli("track-changes", out, "on");

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
