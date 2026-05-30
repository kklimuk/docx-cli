import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/tables.docx — three tables exercising the table feature
 * set, built entirely through CLI calls (a parity check that everything the
 * JSX emitters can do is reachable from the public surface):
 *   1. Plain 2×3 (autofit layout — the `insert --table` default)
 *   2. Custom column widths (1440 / 2880 / 4320 twips, fixed layout)
 *   3. 3×3 with merges — the top row's first cell spans 2 columns horizontally,
 *      and column 0 of rows 1-2 is vertically merged (via `tables merge`).
 *
 * The Heading2 style the section headings reference is auto-provisioned by
 * `insert --style` (ensureReferencedStyle).
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/tables.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

// Each insert appends after the previously-added block, so ids are predictable
// across the sequence (the CLI re-reads — and thus re-derives pN/tN — on every
// invocation).
await cli("create", out, "--force", "--text", "Tables fixture");

await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--text",
	"Plain 2×3",
	"--style",
	"Heading2",
); // p1
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--table",
	"--rows",
	"2",
	"--cols",
	"3",
); // t0 (autofit)

await cli(
	"insert",
	out,
	"--after",
	"t0",
	"--text",
	"Custom widths",
	"--style",
	"Heading2",
); // p2
await cli(
	"insert",
	out,
	"--after",
	"p2",
	"--table",
	"--rows",
	"2",
	"--cols",
	"3",
	"--widths",
	"1440,2880,4320",
); // t1 (fixed)

// Fill the custom-widths table cells via `edit --at tN:rRcC:pK`.
await cli("edit", out, "--at", "t1:r0c0:p0", "--text", "narrow");
await cli("edit", out, "--at", "t1:r0c1:p0", "--text", "medium");
await cli("edit", out, "--at", "t1:r0c2:p0", "--text", "wide");
await cli("edit", out, "--at", "t1:r1c0:p0", "--text", "1″");
await cli("edit", out, "--at", "t1:r1c1:p0", "--text", "2″");
await cli("edit", out, "--at", "t1:r1c2:p0", "--text", "3″");

await cli(
	"insert",
	out,
	"--after",
	"t1",
	"--text",
	"Merges (gridSpan + vMerge)",
	"--style",
	"Heading2",
); // p3

// 3. Merge table — 3×3 filled then collapsed via `tables merge`. Content is set
//    on the anchor cells before merging (merge keeps the leftmost/top cell's
//    content and empties the rest), then a horizontal merge on the top row and
//    a vertical merge on column 0.
await cli(
	"insert",
	out,
	"--after",
	"p3",
	"--table",
	"--rows",
	"3",
	"--cols",
	"3",
	"--widths",
	"2880,2880,2880",
); // t2

await cli("edit", out, "--at", "t2:r0c0:p0", "--text", "Spans cols 0–1");
await cli("edit", out, "--at", "t2:r0c2:p0", "--text", "col 2");
await cli("edit", out, "--at", "t2:r1c0:p0", "--text", "Spans rows 1–2");
await cli("edit", out, "--at", "t2:r1c1:p0", "--text", "r1c1");
await cli("edit", out, "--at", "t2:r1c2:p0", "--text", "r1c2");
await cli("edit", out, "--at", "t2:r2c1:p0", "--text", "r2c1");
await cli("edit", out, "--at", "t2:r2c2:p0", "--text", "r2c2");

await cli("tables", "merge", out, "--at", "t2:r0c0-r0c1"); // horizontal span
await cli("tables", "merge", out, "--at", "t2:r1c0-r2c0"); // vertical merge

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
