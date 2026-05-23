import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/code-blocks.docx — a doc that opens with prose, then a
 * multi-line TypeScript code block with lowlight syntax highlighting, then an
 * uncolored plaintext code block, then a paragraph with an inline code span.
 * Exercises:
 *   - `insert --code TEXT` (multi-line splits into N CodeBlock paragraphs)
 *   - `insert --code-file PATH` (file content path)
 *   - `--language typescript` (lowlight colors)
 *   - inline `runStyle: "Code"` via `--runs` JSON
 *   - `CodeBlock` paragraph style + `Code` character style both land in
 *     styles.xml via `ensureReferencedStyle`
 *
 * The fixture joins CORE_FIXTURES so LibreOffice round-trips the colored
 * runs, the rStyle references, and the multi-paragraph code block layout.
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/code-blocks.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Code blocks fixture.");
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--text",
	"Below: a TypeScript snippet with syntax highlighting.",
);
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--code",
	"function add(a: number, b: number): number {\n  // returns the sum\n  return a + b;\n}",
	"--language",
	"typescript",
);
// 4 code paragraphs were just inserted at p2..p5; section break is now s0
// at the tail. Add a plaintext code block (no language) after p5.
await cli(
	"insert",
	out,
	"--after",
	"p5",
	"--text",
	"And here is plain text without highlighting:",
);
await cli(
	"insert",
	out,
	"--after",
	"p6",
	"--code",
	"$ git log --oneline\nabc1234 fix: edge case\ndef5678 feat: new feature",
);
// 3 more code paragraphs (p7..p9), then an inline code span at p10.
await cli(
	"insert",
	out,
	"--after",
	"p9",
	"--runs",
	JSON.stringify([
		{ type: "text", text: "Run " },
		{ type: "text", text: "docx insert --code", runStyle: "Code" },
		{ type: "text", text: " to author code blocks." },
	]),
);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
