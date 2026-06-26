import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/tables-mutations.docx — exercises the S4b `docx tables`
 * verb set, entirely through CLI calls (a parity check that the verbs are
 * reachable from the public surface).
 *
 * Table t0 (untracked structural edits + formatting): a 3×3 table that gets a
 * horizontal merge, a vertical merge, custom percentage widths, double borders,
 * and then `tables format` — header-row shading + vertical/horizontal centering
 * + repeat-header, a per-cell double bottom border, a taller last row, and the
 * whole table centered on the page. Dogfoods the format verb's <w:shd>/<w:vAlign>/
 * paragraph <w:jc>/<w:tcBorders>/<w:trHeight>/<w:tblHeader>/table-<w:jc> XML
 * through the LibreOffice round-trip.
 *
 * Table t1 (tracked edits, left UNACCEPTED): a 2×2 table edited with
 * track-changes ON so the saved file carries every table revision marker —
 * <w:trPr><w:ins> (rowIns), <w:trPr><w:del> (rowDel), <w:tcPr><w:cellIns>
 * (column insert) and its paired <w:tblGridChange>, and <w:tcPr><w:cellDel>
 * (column delete). Keeping them unaccepted is the point: the LibreOffice
 * round-trip then validates that Word's table-revision XML survives a
 * load/save.
 *
 * Table t2 (clean, unmerged 2×3 autofit): the equation-in-cell regression
 * host absorbed from the former tables.docx fixture — a stable, unmerged
 * `t2:r0c0:p0` for inserting an equation and toggling --display. Inserted
 * while track-changes is OFF (before t1's tracked edits) so it carries no
 * revision markers, and the generator turns track-changes back OFF at the
 * end so the untracked equation edit in equations.test.ts applies directly.
 */

const root = resolve(import.meta.dir, "../../..");
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
// t0 — cell/row/table FORMATTING (dogfoods `tables format`, so the LibreOffice
// round-trip validates the shading/vAlign/jc/tcBorders/trHeight/tblHeader XML).
// Header row: grey fill + centered both ways (broadcasts over the merged cell).
await cli(
	"tables",
	"format",
	out,
	"--at",
	"t0:r0",
	"--shade",
	"D9D9D9",
	"--valign",
	"center",
	"--halign",
	"center",
	"--repeat-header",
);
// A bottom rule under a plain cell, a taller last row, and the table centered.
await cli(
	"tables",
	"format",
	out,
	"--at",
	"t0:r2c2",
	"--cell-borders",
	"bottom",
	"--border-style",
	"double",
);
await cli("tables", "format", out, "--at", "t0:r2", "--row-height", "0.4in");
await cli("tables", "format", out, "--at", "t0", "--align", "center");

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

// t2 — clean, unmerged 2×3 autofit table (the equation-in-cell regression
// host migrated from the deleted tables.docx). Inserted while track-changes
// is still OFF so the cell carries no revision markers and `t2:r0c0:p0` is a
// stable insert target.
await cli(
	"insert",
	out,
	"--after",
	"t1",
	"--text",
	"Equation host (clean cell)",
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
); // t2 (clean autofit)

await cli("track-changes", out, "on");
await cli("tables", "insert-row", out, "--at", "t1", "--cells", "added,row"); // rowIns
await cli("tables", "insert-column", out, "--at", "t1", "--position", "1"); // cellIns + tblGridChange
await cli("tables", "delete-row", out, "--at", "t1:r0"); // rowDel
await cli("tables", "delete-column", out, "--at", "t1:c2"); // cellDel (per row)

// Restore the document toggle to OFF (the tracked-table revision markers from
// the four edits above persist — `track-changes off` only clears the settings
// flag). This keeps t2's cell a direct-edit target so the untracked
// `edit --display` in equations.test.ts flips display without being recorded
// as a tracked change.
await cli("track-changes", out, "off");

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
