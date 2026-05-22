import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/tables-mutations.docx — exercises the S4b `docx tables`
 * verb set, entirely through CLI calls (a parity check that the verbs are
 * reachable from the public surface).
 *
 * Table t0 (untracked structural edits): a 3×3 table that gets a horizontal
 * merge, a vertical merge, custom percentage widths, and double borders.
 *
 * Table t1 (tracked edits, left UNACCEPTED): a 2×2 table edited with
 * track-changes ON so the saved file carries every table revision marker —
 * <w:trPr><w:ins> (rowIns), <w:trPr><w:del> (rowDel), <w:tcPr><w:cellIns>
 * (column insert) and its paired <w:tblGridChange>, and <w:tcPr><w:cellDel>
 * (column delete). Keeping them unaccepted is the point: the LibreOffice
 * round-trip then validates that Word's table-revision XML survives a
 * load/save.
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/tables-mutations.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`
		.env({
			...process.env,
			DOCX_AUTHOR: "Fixture",
			DOCX_CLI_NOW: "2026-05-21T00:00:00Z",
		})
		.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Tables mutations fixture");

// t0 — untracked structural edits.
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--table",
	"--rows",
	"3",
	"--cols",
	"3",
);
await cli("tables", "merge", out, "--at", "t0:r0c0-r0c1"); // horizontal
await cli("tables", "merge", out, "--at", "t0:r1c0-r2c0"); // vertical
await cli("tables", "set-widths", out, "--at", "t0", "--widths", "25%,25%,50%");
await cli(
	"tables",
	"borders",
	out,
	"--at",
	"t0",
	"--style",
	"double",
	"--size",
	"8",
);

// t1 — tracked edits, left unaccepted so the markers persist in the fixture.
await cli(
	"insert",
	out,
	"--after",
	"t0",
	"--table",
	"--rows",
	"2",
	"--cols",
	"2",
);
await cli("track-changes", out, "on");
await cli("tables", "insert-row", out, "--at", "t1", "--cells", "added,row"); // rowIns
await cli("tables", "insert-column", out, "--at", "t1", "--position", "1"); // cellIns + tblGridChange
await cli("tables", "delete-row", out, "--at", "t1:r0"); // rowDel
await cli("tables", "delete-column", out, "--at", "t1:c2"); // cellDel (per row)

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
