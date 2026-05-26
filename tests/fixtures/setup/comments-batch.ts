import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/comments-batch.docx — a three-paragraph doc used to
 * exercise B1's bulk + anchored comment surface (`comments add --anchor`,
 * `--batch`, `--occurrence`; `comments delete --id ... --id ...`;
 * `comments resolve --batch ...`).
 *
 *   p0: "Alpha is first."
 *   p1: "Beta is second."
 *   p2: "Gamma is third."
 *
 * Each paragraph has a unique opening word (Alpha / Beta / Gamma) for the
 * unique-anchor case, and shares the substring "is" so the multi-match /
 * --occurrence path is exercised.
 *
 * Built by dogfooding the CLI; doubles as an end-to-end smoke test for
 * `create` and `insert`.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/comments-batch.docx");
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
	"Comments batch fixture",
	"--author",
	"docx-cli",
	"--text",
	"Alpha is first.",
	"--force",
);

await cli("insert", out, "--after", "p0", "--text", "Beta is second.");
await cli("insert", out, "--after", "p1", "--text", "Gamma is third.");

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
