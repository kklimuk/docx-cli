import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";
import { openDocView, saveDocView } from "../src/core/ast/doc-view";
import { Paragraph } from "../src/core/blocks";
import type { XmlNode } from "../src/core/parser";
import { ensureStyle } from "../src/core/styles";
import { Table, TableCell, TableRow } from "../src/core/table";

/**
 * Build tests/fixtures/tables.docx — three tables exercising the S4a feature
 * set:
 *   1. Plain 2×3 (autofit layout — the `insert --table` default)
 *   2. Custom column widths (1440 / 2880 / 4320 twips, fixed layout)
 *   3. 3×3 with merges — top row's first cell spans 2 cols horizontally, and
 *      column 0 of rows 1-2 is vertically merged.
 *
 * Tables 1 and 2 (plus all cell content and the section headings) are built
 * entirely through CLI calls — this doubles as a parity check that everything
 * the JSX emitters can do for plain/widthed tables is reachable from the
 * public CLI surface.
 *
 * Table 3 is built via the JSX emitter directly: gridSpan / vMerge merges have
 * no CLI surface yet (that's the S4b `docx tables merge` / `unmerge` verbs).
 * Once S4b lands this can be rebuilt from CLI calls too. The Heading2 style the
 * CLI headings reference is provisioned here via ensureStyle — `insert --style`
 * sets the pStyle reference but doesn't define the style (also pending wider
 * style-provisioning work).
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/tables.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

// 1. Base doc + headings + tables 1-2, all via CLI. Each insert appends after
//    the previously-added block, so ids are predictable across the sequence
//    (the CLI re-reads — and thus re-derives pN/tN — on every invocation).
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
); // t1 (fixed via --widths)

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

// 2. Merge table — no CLI surface yet (S4b). Splice it in after the merges
//    heading via the JSX emitter. Also provision Heading2 so the CLI-inserted
//    headings actually render as headings.
const filledCell = (text: string): XmlNode => (
	<TableCell>
		<Paragraph text={text} />
	</TableCell>
);

const tableMerges = (
	<Table grid={[2880, 2880, 2880]} layout="fixed">
		<TableRow>
			{/* Top row: first cell spans 2 columns, then a normal third cell. */}
			<TableCell gridSpan={2}>
				<Paragraph text="Spans cols 0–1" />
			</TableCell>
			{filledCell("col 2")}
		</TableRow>
		<TableRow>
			<TableCell vMerge="restart">
				<Paragraph text="Spans rows 1–2" />
			</TableCell>
			{filledCell("r1c1")}
			{filledCell("r1c2")}
		</TableRow>
		<TableRow>
			<TableCell vMerge="continue">
				<Paragraph text="" />
			</TableCell>
			{filledCell("r2c1")}
			{filledCell("r2c2")}
		</TableRow>
	</Table>
);

const view = await openDocView(out);
ensureStyle(view, "Heading2");
const mergesHeading = view.blockReferences.get("p3");
if (!mergesHeading) throw new Error("expected p3 (merges heading)");
const insertIndex = mergesHeading.parent.indexOf(mergesHeading.node);
if (insertIndex === -1)
	throw new Error("merges heading not found in its parent");
mergesHeading.parent.splice(insertIndex + 1, 0, tableMerges);
await saveDocView(view);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
