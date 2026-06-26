import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import JSZip from "jszip";
import { runCli, tempWorkspace } from "./harness";
import { readDocumentXml, readMarkdown, trackedKinds } from "./helpers";

async function tableLayout(docPath: string): Promise<string | null> {
	const pkg = await Pkg.open(docPath);
	const xml = await pkg.readText("word/document.xml");
	const match = xml.match(/<w:tblLayout\s+w:type="(\w+)"\s*\/>/);
	return match?.[1] ?? null;
}

type Cell = {
	blocks: Array<{ id: string; type: string; runs?: Array<{ text?: string }> }>;
	gridSpan?: number;
	vMerge?: "restart" | "continue";
	width?: { value: number; unit: string };
	shading?: string;
};

type TableBlock = {
	id: string;
	type: "table";
	grid: number[];
	width?: { value: number; unit: string };
	borders?: string;
	rows: Array<{ cells: Cell[] }>;
};

async function tableFromDoc(docPath: string): Promise<TableBlock | undefined> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as {
		blocks: Array<{ id: string; type: string }>;
	};
	return doc.blocks.find(
		(block): block is TableBlock => block.type === "table",
	);
}

describe("docx insert --table", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("tables");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Before table");
	});

	test("basic --table --rows 2 --cols 3 produces 2×3 grid with even-width columns", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"3",
		);
		expect(result.exitCode).toBe(0);

		const table = await tableFromDoc(docPath);
		expect(table).toBeDefined();
		expect(table?.grid).toHaveLength(3);
		expect(table?.rows).toHaveLength(2);
		for (const row of table?.rows ?? []) {
			expect(row.cells).toHaveLength(3);
		}
		// Default 100% width
		expect(table?.width).toEqual({ value: 5000, unit: "pct" });
		// Even split of the 9360-twip page-content width across 3 cols (with
		// remainder bumping the first column).
		const sum = (table?.grid ?? []).reduce((acc, w) => acc + w, 0);
		expect(sum).toBe(9360);
	});

	test("--widths sets per-column grid widths", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"3",
			"--widths",
			"1440,2880,4320",
		);
		const table = await tableFromDoc(docPath);
		expect(table?.grid).toEqual([1440, 2880, 4320]);
	});

	test("--table-width 50% writes pct units (2500)", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"2",
			"--table-width",
			"50%",
		);
		const table = await tableFromDoc(docPath);
		expect(table?.width).toEqual({ value: 2500, unit: "pct" });
	});

	test("--table-width as twips writes dxa units", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"2",
			"--table-width",
			"4320",
		);
		const table = await tableFromDoc(docPath);
		expect(table?.width).toEqual({ value: 4320, unit: "dxa" });
	});

	test("default layout is autofit (no --widths, no --layout)", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"2",
		);
		expect(await tableLayout(docPath)).toBe("autofit");
	});

	test("--widths implies fixed layout", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"2",
			"--widths",
			"2000,3000",
		);
		expect(await tableLayout(docPath)).toBe("fixed");
	});

	test("--layout autofit overrides the widths-implies-fixed default", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"2",
			"--widths",
			"2000,3000",
			"--layout",
			"autofit",
		);
		expect(await tableLayout(docPath)).toBe("autofit");
	});

	test("--layout fixed without --widths", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"2",
			"--layout",
			"fixed",
		);
		expect(await tableLayout(docPath)).toBe("fixed");
	});

	test("invalid --layout is rejected", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
			"--layout",
			"sideways",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("--layout must be");
	});

	test("--widths length must equal --cols", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"3",
			"--widths",
			"1000,2000",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("--widths length");
	});

	test("--rows / --cols required when --table is set", async () => {
		const result = await runCli("insert", docPath, "--after", "p0", "--table");
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("--rows");
	});

	test("--rows without --table fails", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--rows",
			"2",
			"--text",
			"x",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("--rows requires --table");
	});

	test("--text and --table are mutually exclusive", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
			"--text",
			"hello",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("Pass only one");
	});

	test("inserting a table while track-changes is on is rejected", async () => {
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("track-changes");
	});
});

describe("docx edit --at tN:rRcC:pK", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("tables-edit");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Before table");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"2",
		);
	});

	test("replaces the paragraph in the targeted cell", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"t0:r0c1:p0",
			"--text",
			"TopRight",
		);
		expect(result.exitCode).toBe(0);

		const table = await tableFromDoc(docPath);
		const cell = table?.rows[0]?.cells[1];
		const text = cell?.blocks[0]?.runs?.[0]?.text;
		expect(text).toBe("TopRight");
		// Other cells stay empty.
		expect(table?.rows[0]?.cells[0]?.blocks[0]?.runs ?? []).toEqual([]);
		expect(table?.rows[1]?.cells[0]?.blocks[0]?.runs ?? []).toEqual([]);
	});

	test("preserves table grid / dimensions across a cell edit", async () => {
		await runCli("edit", docPath, "--at", "t0:r1c0:p0", "--text", "BotLeft");
		const table = await tableFromDoc(docPath);
		expect(table?.grid).toHaveLength(2);
		expect(table?.rows).toHaveLength(2);
		for (const row of table?.rows ?? []) {
			expect(row.cells).toHaveLength(2);
		}
	});
});

type MutCell = {
	blocks: Array<{ runs?: Array<{ text?: string }> }>;
	gridSpan?: number;
	vMerge?: "restart" | "continue";
	trackedChange?: { kind: string };
};

type MutTableBlock = {
	id: string;
	type: "table";
	grid: number[];
	width?: { value: number; unit: string };
	rows: Array<{ cells: MutCell[]; trackedChange?: { kind: string } }>;
};

async function table(docPath: string): Promise<MutTableBlock> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as { blocks: Array<{ type: string }> };
	const found = doc.blocks.find((block) => block.type === "table");
	if (!found) throw new Error("no table in document");
	return found as MutTableBlock;
}

function cellText(cell: MutCell): string {
	return cell.blocks[0]?.runs?.[0]?.text ?? "";
}

async function newTableDoc(
	label: string,
	rows: number,
	cols: number,
): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "out.docx");
	await runCli("create", docPath, "--text", "Before");
	await runCli(
		"insert",
		docPath,
		"--after",
		"p0",
		"--table",
		"--rows",
		String(rows),
		"--cols",
		String(cols),
	);
	return docPath;
}

describe("docx tables insert-row", () => {
	let docPath: string;
	beforeEach(async () => {
		docPath = await newTableDoc("ins-row", 2, 3);
	});

	test("appends a row at the end by default", async () => {
		const result = await runCli("tables", "insert-row", docPath, "--at", "t0");
		expect(result.exitCode).toBe(0);
		const t = await table(docPath);
		expect(t.rows).toHaveLength(3);
		expect(t.rows[2]?.cells).toHaveLength(3);
	});

	test("--position inserts at the given index", async () => {
		await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--position",
			"0",
			"--cells",
			"a,b,c",
		);
		const t = await table(docPath);
		expect(t.rows).toHaveLength(3);
		expect(t.rows[0]?.cells.map(cellText)).toEqual(["a", "b", "c"]);
	});

	test("inserting a row inside a vertical merge extends it (matches Word)", async () => {
		const merged = await newTableDoc("ins-row-merge", 3, 2);
		// col 0 merged down all three rows.
		await runCli("tables", "merge", merged, "--at", "t0:r0c0-r2c0");
		const result = await runCli(
			"tables",
			"insert-row",
			merged,
			"--at",
			"t0",
			"--position",
			"1",
		);
		expect(result.exitCode).toBe(0);
		const t = await table(merged);
		expect(t.rows).toHaveLength(4);
		// The new row joins the merge as a continuation rather than splitting it.
		expect(t.rows.map((row) => row.cells[0]?.vMerge)).toEqual([
			"restart",
			"continue",
			"continue",
			"continue",
		]);
	});

	test("inserting a row below a vertical merge leaves it a normal row", async () => {
		const merged = await newTableDoc("ins-row-below-merge", 3, 2);
		await runCli("tables", "merge", merged, "--at", "t0:r0c0-r2c0");
		await runCli("tables", "insert-row", merged, "--at", "t0"); // append at end
		const t = await table(merged);
		expect(t.rows).toHaveLength(4);
		expect(t.rows[3]?.cells[0]?.vMerge).toBeUndefined();
	});

	test("--cells exceeding column count is rejected", async () => {
		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--cells",
			"a,b,c,d",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("columns");
	});

	test("inserted row copies the sibling row's gridSpan pattern (invoice footgun)", async () => {
		// 3-grid-col table whose data rows merge cols 1+2 → 2 logical columns.
		const merged = await newTableDoc("ins-row-span", 2, 3);
		await runCli("tables", "merge", merged, "--at", "t0:r1c1-r1c2");
		// One value per LOGICAL cell (2), not per grid column (3).
		const result = await runCli(
			"tables",
			"insert-row",
			merged,
			"--at",
			"t0",
			"--position",
			"2",
			"--cells",
			"Label,Spanned",
		);
		expect(result.exitCode).toBe(0);
		const t = await table(merged);
		const newRow = t.rows[2];
		// Mirrors the sibling: 2 cells, the second spanning 2 grid columns — NOT
		// 3 flat cells that would shove values into the wrong columns.
		expect(newRow?.cells).toHaveLength(2);
		expect(newRow?.cells[1]?.gridSpan).toBe(2);
		expect(newRow?.cells.map(cellText)).toEqual(["Label", "Spanned"]);
		// Span-sum matches the grid (3 columns) — no malformed row.
		const span = (c: MutCell) => c.gridSpan ?? 1;
		expect(newRow?.cells.reduce((s, c) => s + span(c), 0)).toBe(3);
	});

	test("inserted row inherits the sibling cells' paragraph alignment (invoice jc footgun)", async () => {
		// The invoice defect: a numeric column is right-aligned in the data rows,
		// but `insert-row` cloned only the gridSpan/merge structure — not the cells'
		// `<w:jc>` — so the new row's number landed left-aligned, out of column.
		const doc = await newTableDoc("ins-row-jc", 2, 2);
		await runCli("edit", doc, "--at", "t0:r1c1:p0", "--alignment", "right");
		const result = await runCli(
			"tables",
			"insert-row",
			doc,
			"--at",
			"t0",
			"--position",
			"2",
		);
		expect(result.exitCode).toBe(0);
		const pkg = await Pkg.open(doc);
		const xml = await pkg.readText("word/document.xml");
		const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
		// Both the sibling data row AND the inserted row now carry jc=right — the
		// new row mirrors the alignment, not just the column structure.
		const rightAligned = rows.filter((row) =>
			/<w:jc w:val="right"/.test(row),
		).length;
		expect(rightAligned).toBe(2);
	});

	test("inherited alignment also applies under a gridSpan'd cell", async () => {
		// 3 grid columns; merge cols 1+2 of the data row into one spanned cell and
		// right-align it. The inserted row mirrors the span AND the alignment.
		const merged = await newTableDoc("ins-row-jc-span", 2, 3);
		await runCli("tables", "merge", merged, "--at", "t0:r1c1-r1c2");
		await runCli("edit", merged, "--at", "t0:r1c1:p0", "--alignment", "right");
		const result = await runCli(
			"tables",
			"insert-row",
			merged,
			"--at",
			"t0",
			"--position",
			"2",
		);
		expect(result.exitCode).toBe(0);
		const t = await table(merged);
		const newRow = t.rows[2];
		// The inserted row's spanned cell exists (gridSpan 2) — and the whole table
		// now carries two right-aligned cells (sibling + inserted), not one.
		expect(newRow?.cells[1]?.gridSpan).toBe(2);
		const xml = await (await Pkg.open(merged)).readText("word/document.xml");
		const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
		expect(rows.filter((row) => /<w:jc w:val="right"/.test(row)).length).toBe(
			2,
		);
	});

	test("inherited alignment survives a tracked insert-row", async () => {
		const doc = await newTableDoc("ins-row-jc-track", 2, 2);
		await runCli("edit", doc, "--at", "t0:r1c1:p0", "--alignment", "right");
		await runCli("track-changes", "on", doc);
		await runCli("tables", "insert-row", doc, "--at", "t0", "--position", "2");
		await runCli("track-changes", "accept", doc, "--all");
		const xml = await (await Pkg.open(doc)).readText("word/document.xml");
		const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
		expect(rows.filter((row) => /<w:jc w:val="right"/.test(row)).length).toBe(
			2,
		);
	});

	test("--cells limit counts logical (visible) columns on a merged table", async () => {
		const merged = await newTableDoc("ins-row-span-limit", 2, 3);
		await runCli("tables", "merge", merged, "--at", "t0:r1c1-r1c2");
		// 2 logical columns now; 3 values must be rejected.
		const result = await runCli(
			"tables",
			"insert-row",
			merged,
			"--at",
			"t0",
			"--position",
			"2",
			"--cells",
			"a,b,c",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("2 columns");
	});

	test("--cells with a shell-gutted currency value ('.00') is refused", async () => {
		// The invoice blocker: `--cells "...,$300.00,$600.00"` double-quoted in bash
		// becomes `...,.00,.00` before docx sees it. Refuse the gutted signature so a
		// weak agent doesn't ship `.00` cells.
		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--cells",
			"Calibration kit,.00",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		expect((result.parsed as { hint: string }).hint).toContain("--batch");
	});

	test("--cells with correct $300.00 is NOT a false positive", async () => {
		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--cells",
			"Calibration kit,$300.00",
		);
		expect(result.exitCode).toBe(0);
	});

	test("missing table is reported", async () => {
		const result = await runCli("tables", "insert-row", docPath, "--at", "t9");
		expect(result.exitCode).toBe(3);
	});
});

describe("docx tables delete-row", () => {
	test("removes the targeted row", async () => {
		const docPath = await newTableDoc("del-row", 3, 2);
		const result = await runCli(
			"tables",
			"delete-row",
			docPath,
			"--at",
			"t0:r1",
		);
		expect(result.exitCode).toBe(0);
		expect((await table(docPath)).rows).toHaveLength(2);
	});

	test("rejects deleting a row that orphans a vertical merge", async () => {
		const docPath = await newTableDoc("del-row-merge", 3, 2);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r1c0");
		const result = await runCli(
			"tables",
			"delete-row",
			docPath,
			"--at",
			"t0:r0",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("orphan");
	});
});

describe("docx tables insert-column", () => {
	let docPath: string;
	beforeEach(async () => {
		docPath = await newTableDoc("ins-col", 2, 2);
	});

	test("appends a column and grows the grid", async () => {
		await runCli(
			"tables",
			"insert-column",
			docPath,
			"--at",
			"t0",
			"--width",
			"1440",
		);
		const t = await table(docPath);
		expect(t.grid).toHaveLength(3);
		expect(t.grid[2]).toBe(1440);
		for (const row of t.rows) expect(row.cells).toHaveLength(3);
	});

	test("--position inserts mid-table", async () => {
		await runCli(
			"tables",
			"insert-column",
			docPath,
			"--at",
			"t0",
			"--position",
			"1",
			"--width",
			"1000",
		);
		const t = await table(docPath);
		expect(t.grid[1]).toBe(1000);
	});

	test("rejects bisecting a horizontal merge", async () => {
		const wide = await newTableDoc("ins-col-merge", 1, 3);
		await runCli("tables", "merge", wide, "--at", "t0:r0c0-r0c2");
		const result = await runCli(
			"tables",
			"insert-column",
			wide,
			"--at",
			"t0",
			"--position",
			"1",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("bisect");
	});
});

describe("docx tables delete-column", () => {
	test("removes the column and its grid entry", async () => {
		const docPath = await newTableDoc("del-col", 2, 3);
		await runCli("tables", "delete-column", docPath, "--at", "t0:c1");
		const t = await table(docPath);
		expect(t.grid).toHaveLength(2);
		for (const row of t.rows) expect(row.cells).toHaveLength(2);
	});

	test("rejects deleting the only column", async () => {
		const docPath = await newTableDoc("del-col-only", 2, 1);
		const result = await runCli(
			"tables",
			"delete-column",
			docPath,
			"--at",
			"t0:c0",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("only column");
	});

	test("rejects deleting through a horizontal merge", async () => {
		const docPath = await newTableDoc("del-col-merge", 1, 3);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c1");
		const result = await runCli(
			"tables",
			"delete-column",
			docPath,
			"--at",
			"t0:c0",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("horizontal merge");
	});
});

describe("docx tables set-widths", () => {
	let docPath: string;
	beforeEach(async () => {
		docPath = await newTableDoc("widths", 1, 3);
	});

	test("percentages convert to twips over the current total", async () => {
		// default grid sums to 9360
		await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"25%,25%,50%",
		);
		const t = await table(docPath);
		expect(t.grid).toEqual([2340, 2340, 4680]);
	});

	test("twips are written verbatim", async () => {
		await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"1000,2000,3000",
		);
		expect((await table(docPath)).grid).toEqual([1000, 2000, 3000]);
	});

	test("count mismatch is rejected", async () => {
		const result = await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"50%,50%",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("entries");
	});

	test("a text cell collapsed below ~0.2in is refused (invoice footgun)", async () => {
		// Put content in the column we're about to crush: 2% of the 9360 grid ≈
		// 187tw ≈ 0.13in, which after ~0.15in of cell margin fits under one char,
		// so Word wraps it one char per line. Refuse rather than return ok.
		await runCli("edit", docPath, "--at", "t0:r0c2:p0", "--text", "Amount");
		const result = await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"96%,2%,2%",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("Amount");
		expect(result.stdout).toContain("margins");
		// The original grid is untouched (no partial write of a broken layout).
		expect((await table(docPath)).grid).toEqual([3120, 3120, 3120]);
	});

	test("an EMPTY thin column is allowed (deliberate spacer, nothing to wrap)", async () => {
		// Same crushing widths, but the cells are empty — a thin spacer renders
		// fine, so the content-aware guard must NOT block it.
		const result = await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"96%,2%,2%",
		);
		expect(result.exitCode).toBe(0);
		expect((await table(docPath)).grid).toEqual([8986, 187, 187]);
	});

	test("count mismatch on a merged-cell table explains grid-vs-visible columns", async () => {
		// Merge c0+c1: the grid still has 3 <w:gridCol> but row 0 shows 2 cells.
		// A 2-value (visible-column) --widths must be told it needs 3 grid values.
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c1");
		const result = await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"50%,50%",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("merged cells");
		expect(result.stdout).toContain("grid columns");
	});
});

describe("docx tables merge / unmerge", () => {
	test("horizontal merge sets gridSpan on the leftmost cell", async () => {
		const docPath = await newTableDoc("merge-h", 1, 3);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c2");
		const t = await table(docPath);
		expect(t.rows[0]?.cells).toHaveLength(1);
		expect(t.rows[0]?.cells[0]?.gridSpan).toBe(3);
	});

	test("vertical merge sets restart/continue", async () => {
		const docPath = await newTableDoc("merge-v", 3, 2);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r2c0");
		const t = await table(docPath);
		expect(t.rows[0]?.cells[0]?.vMerge).toBe("restart");
		expect(t.rows[1]?.cells[0]?.vMerge).toBe("continue");
		expect(t.rows[2]?.cells[0]?.vMerge).toBe("continue");
	});

	test("unmerge restores a horizontal span", async () => {
		const docPath = await newTableDoc("unmerge-h", 1, 3);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c2");
		await runCli("tables", "unmerge", docPath, "--at", "t0:r0c0");
		const t = await table(docPath);
		expect(t.rows[0]?.cells).toHaveLength(3);
		expect(t.rows[0]?.cells[0]?.gridSpan).toBeUndefined();
	});

	test("unmerge restores a vertical merge", async () => {
		const docPath = await newTableDoc("unmerge-v", 3, 2);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r2c0");
		await runCli("tables", "unmerge", docPath, "--at", "t0:r0c0");
		const t = await table(docPath);
		for (const row of t.rows) expect(row.cells[0]?.vMerge).toBeUndefined();
	});

	test("merge rejects a single-cell region", async () => {
		const docPath = await newTableDoc("merge-one", 2, 2);
		const result = await runCli(
			"tables",
			"merge",
			docPath,
			"--at",
			"t0:r0c0-r0c0",
		);
		expect(result.exitCode).not.toBe(0);
	});

	test("unmerge rejects a non-merged cell", async () => {
		const docPath = await newTableDoc("unmerge-plain", 2, 2);
		const result = await runCli(
			"tables",
			"unmerge",
			docPath,
			"--at",
			"t0:r0c0",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("not a merged cell");
	});
});

describe("docx tables borders", () => {
	test("applies a border style to the table", async () => {
		const docPath = await newTableDoc("borders", 1, 2);
		const result = await runCli(
			"tables",
			"borders",
			docPath,
			"--at",
			"t0",
			"--style",
			"double",
			"--size",
			"8",
			"--color",
			"444444",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { style: string }).style).toBe("double");
	});

	test("rejects an invalid style", async () => {
		const docPath = await newTableDoc("borders-bad", 1, 2);
		const result = await runCli(
			"tables",
			"borders",
			docPath,
			"--at",
			"t0",
			"--style",
			"wavy",
		);
		expect(result.exitCode).not.toBe(0);
	});
});

describe("docx tables format", () => {
	test("shades a cell and round-trips through read", async () => {
		const docPath = await newTableDoc("fmt-shade", 2, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--shade",
			"D9D9D9",
		);
		expect(result.exitCode).toBe(0);
		expect(await readDocumentXml(docPath)).toContain('w:fill="D9D9D9"');
		expect(await readMarkdown(docPath)).toContain(
			'docx:cell t0:r0c0 shading="D9D9D9"',
		);
	});

	test("accepts a color name for --shade", async () => {
		const docPath = await newTableDoc("fmt-shade-name", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--shade",
			"grey",
		);
		expect(await readDocumentXml(docPath)).toContain('w:fill="808080"');
	});

	test("sets vertical alignment", async () => {
		const docPath = await newTableDoc("fmt-valign", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--valign",
			"center",
		);
		expect(await readDocumentXml(docPath)).toContain(
			'<w:vAlign w:val="center"/>',
		);
		expect(await readMarkdown(docPath)).toContain('vAlign="center"');
	});

	test("sets horizontal alignment as a paragraph property", async () => {
		const docPath = await newTableDoc("fmt-halign", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--halign",
			"center",
		);
		// halign is a paragraph <w:jc>, NOT a <w:tcPr> child.
		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:jc w:val="center"/>');
		expect(await readMarkdown(docPath)).toContain('halign="center"');
	});

	test("adds a per-cell border in the requested style", async () => {
		const docPath = await newTableDoc("fmt-border", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--cell-borders",
			"bottom",
			"--border-style",
			"double",
		);
		const xml = await readDocumentXml(docPath);
		expect(xml).toContain("<w:tcBorders>");
		expect(xml).toMatch(/<w:bottom w:val="double"/);
		expect(await readMarkdown(docPath)).toContain('borders="bottom:double"');
	});

	test("shades every cell when scoped to the whole table", async () => {
		const docPath = await newTableDoc("fmt-broadcast", 2, 3);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0",
			"--shade",
			"EEEEEE",
		);
		const fills = (await readDocumentXml(docPath)).match(/w:fill="EEEEEE"/g);
		expect(fills).toHaveLength(6); // 2 rows × 3 cols
	});

	test("centers the table and applies a style", async () => {
		const docPath = await newTableDoc("fmt-table", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0",
			"--align",
			"center",
			"--style",
			"LightGrid",
		);
		const md = await readMarkdown(docPath);
		expect(md).toContain('align="center"');
		expect(md).toContain('style="LightGrid"');
	});

	test("sets row height and repeat-header, surfaced in the table note", async () => {
		const docPath = await newTableDoc("fmt-row", 2, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0",
			"--row-height",
			"0.4in",
			"--repeat-header",
		);
		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:trHeight w:val="576" w:hRule="atLeast"\/>/);
		expect(xml).toContain("<w:tblHeader/>");
		const md = await readMarkdown(docPath);
		expect(md).toContain('repeat-header="r0"');
		expect(md).toContain('row-heights="r0:0.4in"');
	});

	test("rejects table-level flags on a cell scope", async () => {
		const docPath = await newTableDoc("fmt-scope-cell", 1, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--align",
			"center",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error: string }).error).toContain("whole table");
	});

	test("rejects row-level flags on a column scope", async () => {
		const docPath = await newTableDoc("fmt-scope-col", 2, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:c0",
			"--row-height",
			"0.5in",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error: string }).error).toContain("a row");
	});

	test("rejects an empty plan", async () => {
		const docPath = await newTableDoc("fmt-empty", 1, 2);
		const result = await runCli("tables", "format", docPath, "--at", "t0:r0c0");
		expect(result.exitCode).not.toBe(0);
	});

	test("rejects an invalid valign value", async () => {
		const docPath = await newTableDoc("fmt-bad-valign", 1, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--valign",
			"middle",
		);
		expect(result.exitCode).not.toBe(0);
	});

	test("rejects a unit-less --row-height (no silent inches default)", async () => {
		const docPath = await newTableDoc("fmt-bad-height", 1, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0",
			"--row-height",
			"28",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error: string }).error).toContain("unit");
	});

	test("rejects --height-rule without --row-height", async () => {
		const docPath = await newTableDoc("fmt-rule-orphan", 1, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0",
			"--repeat-header",
			"--height-rule",
			"exact",
		);
		expect(result.exitCode).not.toBe(0);
	});

	test("cell props track as a tcPrChange that reject restores", async () => {
		const docPath = await newTableDoc("fmt-track-cell", 2, 2);
		await runCli("track-changes", docPath, "on");
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r1c1",
			"--shade",
			"FFFF00",
		);
		expect(await trackedKinds(docPath)).toContain("tcPrChange");
		await runCli("track-changes", "reject", docPath, "--at", "tc0");
		expect(await readDocumentXml(docPath)).not.toContain('w:fill="FFFF00"');
		expect(await trackedKinds(docPath)).toHaveLength(0);
	});

	test("halign tracks as a pPrChange", async () => {
		const docPath = await newTableDoc("fmt-track-halign", 1, 2);
		await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0:r0c0",
			"--halign",
			"right",
			"--track",
		);
		expect(await trackedKinds(docPath)).toContain("pPrChange");
	});

	test("table-level changes under tracking emit an audit comment, not a revision", async () => {
		const docPath = await newTableDoc("fmt-track-table", 1, 2);
		const result = await runCli(
			"tables",
			"format",
			docPath,
			"--at",
			"t0",
			"--align",
			"center",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		// No tracked revision (Word won't revert a hand-authored tblPrChange)…
		expect(await trackedKinds(docPath)).toHaveLength(0);
		// …but an honest [docx-cli] audit comment instead.
		const comments = (await runCli("comments", "list", docPath))
			.parsed as Array<{
			text: string;
		}>;
		expect(
			comments.some((comment) => comment.text.includes("[docx-cli]")),
		).toBe(true);
	});
});

describe("docx tables — tracked changes", () => {
	async function trackedTableDoc(
		label: string,
		rows: number,
		cols: number,
	): Promise<string> {
		const docPath = await newTableDoc(label, rows, cols);
		await runCli("track-changes", docPath, "on");
		return docPath;
	}

	test("insert-row emits a rowIns the AST and list surface", async () => {
		const docPath = await trackedTableDoc("trk-ins-row", 2, 2);
		await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--cells",
			"x,y",
		);
		expect(await trackedKinds(docPath)).toContain("rowIns");
		const t = await table(docPath);
		expect(t.rows[2]?.trackedChange?.kind).toBe("rowIns");
	});

	test("reject of a tracked inserted row removes it", async () => {
		const docPath = await trackedTableDoc("trk-ins-row-rej", 2, 2);
		await runCli("tables", "insert-row", docPath, "--at", "t0");
		expect((await table(docPath)).rows).toHaveLength(3);
		await runCli("track-changes", "reject", docPath, "--all");
		expect((await table(docPath)).rows).toHaveLength(2);
	});

	test("accept of a tracked inserted row keeps it without markers", async () => {
		const docPath = await trackedTableDoc("trk-ins-row-acc", 2, 2);
		await runCli("tables", "insert-row", docPath, "--at", "t0");
		await runCli("track-changes", "accept", docPath, "--all");
		const t = await table(docPath);
		expect(t.rows).toHaveLength(3);
		expect(t.rows[2]?.trackedChange).toBeUndefined();
	});

	test("delete-row marks rowDel; accept removes, reject keeps", async () => {
		const accept = await trackedTableDoc("trk-del-row-acc", 3, 2);
		await runCli("tables", "delete-row", accept, "--at", "t0:r0");
		expect((await table(accept)).rows[0]?.trackedChange?.kind).toBe("rowDel");
		await runCli("track-changes", "accept", accept, "--all");
		expect((await table(accept)).rows).toHaveLength(2);

		const reject = await trackedTableDoc("trk-del-row-rej", 3, 2);
		await runCli("tables", "delete-row", reject, "--at", "t0:r0");
		await runCli("track-changes", "reject", reject, "--all");
		expect((await table(reject)).rows).toHaveLength(3);
	});

	test("insert-column emits cellIns per row plus a tblGridChange", async () => {
		const docPath = await trackedTableDoc("trk-ins-col", 2, 2);
		await runCli("tables", "insert-column", docPath, "--at", "t0");
		const kinds = await trackedKinds(docPath);
		expect(kinds.filter((k) => k === "cellIns")).toHaveLength(2);
		expect(kinds).toContain("tblGridChange");
	});

	test("reject of a tracked inserted column restores the prior grid", async () => {
		const docPath = await trackedTableDoc("trk-ins-col-rej", 2, 2);
		await runCli(
			"tables",
			"insert-column",
			docPath,
			"--at",
			"t0",
			"--position",
			"1",
		);
		expect((await table(docPath)).grid).toHaveLength(3);
		await runCli("track-changes", "reject", docPath, "--all");
		const t = await table(docPath);
		expect(t.grid).toHaveLength(2);
		for (const row of t.rows) expect(row.cells).toHaveLength(2);
	});

	test("accept of a tracked inserted column keeps it and clears markers", async () => {
		const docPath = await trackedTableDoc("trk-ins-col-acc", 2, 2);
		await runCli("tables", "insert-column", docPath, "--at", "t0");
		await runCli("track-changes", "accept", docPath, "--all");
		const t = await table(docPath);
		expect(t.grid).toHaveLength(3);
		for (const row of t.rows) {
			expect(row.cells).toHaveLength(3);
			for (const cell of row.cells) expect(cell.trackedChange).toBeUndefined();
		}
	});

	test("delete-column accept removes cells and resyncs the grid", async () => {
		const docPath = await trackedTableDoc("trk-del-col", 2, 3);
		await runCli("tables", "delete-column", docPath, "--at", "t0:c1");
		// cells marked, grid still intact pre-accept
		expect((await table(docPath)).grid).toHaveLength(3);
		await runCli("track-changes", "accept", docPath, "--all");
		const t = await table(docPath);
		expect(t.grid).toHaveLength(2);
		for (const row of t.rows) expect(row.cells).toHaveLength(2);
	});

	test("set-widths under tracking emits tblGridChange + per-cell tcPrChange", async () => {
		const docPath = await trackedTableDoc("trk-widths", 2, 3);
		await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			"t0",
			"--widths",
			"1000,2000,6000",
		);
		const kinds = await trackedKinds(docPath);
		expect(kinds).toContain("tblGridChange");
		// One tcPrChange per cell (2 rows × 3 cols).
		expect(kinds.filter((k) => k === "tcPrChange")).toHaveLength(6);
	});

	test("set-widths reject restores the prior grid; accept keeps the new", async () => {
		const reject = await trackedTableDoc("trk-widths-rej", 1, 3);
		const before = (await table(reject)).grid;
		await runCli(
			"tables",
			"set-widths",
			reject,
			"--at",
			"t0",
			"--widths",
			"1000,2000,6000",
		);
		await runCli("track-changes", "reject", reject, "--all");
		expect((await table(reject)).grid).toEqual(before);

		const accept = await trackedTableDoc("trk-widths-acc", 1, 3);
		await runCli(
			"tables",
			"set-widths",
			accept,
			"--at",
			"t0",
			"--widths",
			"1000,2000,6000",
		);
		await runCli("track-changes", "accept", accept, "--all");
		expect((await table(accept)).grid).toEqual([1000, 2000, 6000]);
	});

	// Word applies merges and border changes immediately even under tracking
	// (no revision marker it will revert), so we match that and note it with an
	// audit comment rather than emitting a revision.
	test("merge under tracking applies and leaves an audit comment", async () => {
		const docPath = await trackedTableDoc("trk-merge", 2, 3);
		await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c2");
		expect((await table(docPath)).rows[0]?.cells[0]?.gridSpan).toBe(3);
		expect(await trackedKinds(docPath)).toHaveLength(0);
		expect(await auditComments(docPath)).toBe(true);
	});

	test("borders under tracking applies and leaves an audit comment", async () => {
		const docPath = await trackedTableDoc("trk-borders", 2, 2);
		const result = await runCli(
			"tables",
			"borders",
			docPath,
			"--at",
			"t0",
			"--style",
			"double",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(docPath)).toHaveLength(0);
		expect(await auditComments(docPath)).toBe(true);
	});

	async function auditComments(docPath: string): Promise<boolean> {
		const comments = await runCli("comments", "list", docPath);
		const bodies = (comments.parsed as Array<{ text: string }>).map(
			(c) => c.text,
		);
		return bodies.some((text) => text.includes("[docx-cli]"));
	}
});

// track-flag covers `tables delete-row --track`; the toggle-off `--track`
// path for the remaining subverbs (a distinct code path from `track-changes
// on`) was otherwise only ever exercised via the global toggle. These pin it.
describe("docx tables — --track forces tracking with the toggle off", () => {
	// The subverbs whose edits ARE representable as OOXML revisions: row/cell
	// insert+delete and grid-width changes. (merge / unmerge / borders apply
	// structurally and are NOT tracked — even under the global toggle — see the
	// separate test below.)
	const trackable: Array<{ name: string; args: (doc: string) => string[] }> = [
		{
			name: "insert-row",
			args: (d) => ["tables", "insert-row", d, "--at", "t0", "--track"],
		},
		{
			name: "insert-column",
			args: (d) => ["tables", "insert-column", d, "--at", "t0", "--track"],
		},
		{
			name: "delete-column",
			args: (d) => ["tables", "delete-column", d, "--at", "t0:c1", "--track"],
		},
		{
			name: "set-widths",
			args: (d) => [
				"tables",
				"set-widths",
				d,
				"--at",
				"t0",
				"--widths",
				"1000,2000,6000",
				"--track",
			],
		},
	];

	for (const subverb of trackable) {
		test(`${subverb.name} --track records a tracked change with the toggle off`, async () => {
			const docPath = await newTableDoc(`track-flag-${subverb.name}`, 2, 3);
			const result = await runCli(...subverb.args(docPath));
			expect(result.exitCode).toBe(0);
			expect((await trackedKinds(docPath)).length).toBeGreaterThan(0);
		});
	}

	test("merge / unmerge / borders apply directly even with --track (not tracked revisions)", async () => {
		// These structural ops have no OOXML tracked-revision representation, so
		// --track is accepted but records nothing — consistent with the global
		// toggle (verified: `merge`/`borders` under `track-changes on` also emit
		// no revisions). The change still applies.
		const merged = await newTableDoc("track-flag-merge", 2, 3);
		expect(
			(
				await runCli(
					"tables",
					"merge",
					merged,
					"--at",
					"t0:r0c0-r0c2",
					"--track",
				)
			).exitCode,
		).toBe(0);
		expect(await trackedKinds(merged)).toHaveLength(0);
		expect((await tableFromDoc(merged))?.rows[0]?.cells[0]?.gridSpan).toBe(3);

		const bordered = await newTableDoc("track-flag-borders", 2, 3);
		const result = await runCli(
			"tables",
			"borders",
			bordered,
			"--at",
			"t0",
			"--style",
			"double",
			"--size",
			"8",
			"--color",
			"444444",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(bordered)).toHaveLength(0);
	});

	test("no --track on an untracked doc records nothing", async () => {
		const docPath = await newTableDoc("track-flag-control", 2, 3);
		await runCli("tables", "insert-row", docPath, "--at", "t0");
		expect(await trackedKinds(docPath)).toHaveLength(0);
	});
});

// Regression: Word emits compound rubric/layout cells as a wrapper paragraph
// plus a nested <w:tbl> holding the real text. Before the fix in
// `readCellBlocks`, the nested table was silently dropped from the AST, so its
// content vanished from `read --ast` and from the GFM render — and the user
// (or a downstream LLM) would see the cell as "empty" even though Word shows
// text. See the `Students Choice.docx` bug report.

const NESTED_CELL_TEXT = "Nested rubric text";

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr/>
<w:p/>
<w:tbl>
<w:tblPr><w:tblW w:w="2460" w:type="dxa"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="1230"/><w:gridCol w:w="1230"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr/><w:p><w:r><w:t>${NESTED_CELL_TEXT}</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr/><w:p><w:r><w:t>InnerRight</w:t></w:r></w:p></w:tc>
</w:tr>
</w:tbl>
<w:p/>
</w:tc></w:tr>
</w:tbl>
<w:sectPr/>
</w:body></w:document>`;

async function makeDoc(): Promise<string> {
	const workspace = tempWorkspace("tables-nested");
	const docPath = join(workspace, "out.docx");
	// Seed from minimal fixture so the package has rels / content types.
	await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
	const pkg = await Pkg.open(docPath);
	pkg.writeText("word/document.xml", DOC);
	await pkg.save();
	return docPath;
}

describe("nested tables inside a cell", () => {
	test("AST surfaces the nested table with chained ids", async () => {
		const docPath = await makeDoc();
		const result = await runCli("read", docPath, "--ast");
		expect(result.exitCode).toBe(0);
		const doc = result.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				rows?: Array<{
					cells: Array<{
						blocks: Array<{
							id: string;
							type: string;
							runs?: Array<{ text?: string }>;
							rows?: Array<{
								cells: Array<{
									blocks: Array<{
										id: string;
										type: string;
										runs?: Array<{ text?: string }>;
									}>;
								}>;
							}>;
						}>;
					}>;
				}>;
			}>;
		};
		const outerTable = doc.blocks.find((b) => b.id === "t0");
		expect(outerTable?.type).toBe("table");
		const outerCell = outerTable?.rows?.[0]?.cells[0];
		expect(outerCell?.blocks).toHaveLength(3);
		// wrapper p, nested table, trailing p (required after a nested tbl)
		expect(outerCell?.blocks[0]?.id).toBe("t0:r0c0:p0");
		expect(outerCell?.blocks[1]?.id).toBe("t0:r0c0:t0");
		expect(outerCell?.blocks[1]?.type).toBe("table");
		expect(outerCell?.blocks[2]?.id).toBe("t0:r0c0:p1");

		const nestedCell = outerCell?.blocks[1]?.rows?.[0]?.cells[0];
		const nestedParagraph = nestedCell?.blocks[0];
		expect(nestedParagraph?.id).toBe("t0:r0c0:t0:r0c0:p0");
		expect(nestedParagraph?.runs?.[0]?.text).toBe(NESTED_CELL_TEXT);
	});

	test("GFM render includes the nested table's text in the outer cell", async () => {
		const docPath = await makeDoc();
		const result = await runCli("read", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(NESTED_CELL_TEXT);
		// The nested-table id should appear as the run-comment marker so the
		// rendered cell is locator-addressable.
		expect(result.stdout).toContain("t0:r0c0:t0:r0c0:p0");
	});

	test("a nested-cell paragraph locator resolves for `read --from`", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"read",
			docPath,
			"--from",
			"t0:r0c0:t0:r0c0:p0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(NESTED_CELL_TEXT);
	});
});

// #6 regression: a cell whose ONLY block is a nested table (no paragraph at all)
// — cellAddress derived the docx:cell address from the first PARAGRAPH, so a
// shaded/merged nested-table-only cell emitted an ADDRESSLESS `<!-- docx:cell -->`,
// leaving an agent unable to map it to a --at locator. The address is derivable
// from the first block of any type (`t0:r0c0:t0` → `t0:r0c0`).
const NESTED_ONLY_DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00B0F0"/></w:tcPr>
<w:tbl>
<w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:tc>
<w:tc><w:tcPr/><w:p><w:r><w:t>right</w:t></w:r></w:p></w:tc>
</w:tr>
</w:tbl>
<w:sectPr/>
</w:body></w:document>`;

async function makeNestedOnlyDoc(): Promise<string> {
	const workspace = tempWorkspace("tables-nested-only");
	const docPath = join(workspace, "out.docx");
	await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
	const pkg = await Pkg.open(docPath);
	pkg.writeText("word/document.xml", NESTED_ONLY_DOC);
	await pkg.save();
	return docPath;
}

describe("docx:cell address with a nested-table-only cell", () => {
	test("the docx:cell note carries its address even with no paragraph in the cell", async () => {
		const docPath = await makeNestedOnlyDoc();
		const md = await readMarkdown(docPath);
		expect(md).toContain('<!-- docx:cell t0:r0c0 shading="00B0F0" -->');
		// Sanity: the outer cell's FIRST block really is the nested table (no
		// paragraph), which is what broke the paragraph-only address derivation.
		const ast = (await runCli("read", docPath, "--ast")).parsed as {
			blocks: Array<{
				rows?: Array<{
					cells: Array<{ blocks: Array<{ id: string; type: string }> }>;
				}>;
			}>;
		};
		const firstBlock = ast.blocks[0]?.rows?.[0]?.cells[0]?.blocks[0];
		expect(firstBlock?.type).toBe("table");
		expect(firstBlock?.id).toBe("t0:r0c0:t0");
	});
});

describe("wc with nested locators", () => {
	// "Nested rubric text" = 3 words, "InnerRight" = 1 word; nested table = 4.
	test("counts words in a nested-cell paragraph", async () => {
		const docPath = await makeDoc();
		const result = await runCli("wc", docPath, "t0:r0c0:t0:r0c0:p0");
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { scope: string; words: number };
		expect(ack.scope).toBe("paragraph");
		expect(ack.words).toBe(3);
	});

	test("counts words in a whole nested cell", async () => {
		const docPath = await makeDoc();
		const result = await runCli("wc", docPath, "t0:r0c0:t0:r0c0");
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { scope: string; words: number };
		expect(ack.scope).toBe("cell");
		expect(ack.words).toBe(3);
	});

	test("counts words in a whole nested table", async () => {
		const docPath = await makeDoc();
		const result = await runCli("wc", docPath, "t0:r0c0:t0");
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { scope: string; words: number };
		expect(ack.scope).toBe("table");
		expect(ack.words).toBe(4);
	});

	test("counts a span inside a nested-cell paragraph", async () => {
		const docPath = await makeDoc();
		// First 6 chars of "Nested rubric text" = "Nested" → 1 word.
		const result = await runCli("wc", docPath, "t0:r0c0:t0:r0c0:p0:0-6");
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { scope: string; words: number };
		expect(ack.scope).toBe("paragraphSpan");
		expect(ack.words).toBe(1);
	});
});

describe("insert / edit with nested locators", () => {
	test("insert --after a nested-cell paragraph drops a sibling in the same nested cell", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"t0:r0c0:t0:r0c0:p0",
			"--text",
			"AddedSibling",
		);
		expect(result.exitCode).toBe(0);
		const ast = await runCli("read", docPath, "--ast");
		const doc = ast.parsed as {
			blocks: Array<{
				id: string;
				rows?: Array<{
					cells: Array<{
						blocks: Array<{
							id: string;
							type: string;
							rows?: Array<{
								cells: Array<{
									blocks: Array<{
										id: string;
										runs?: Array<{ text?: string }>;
									}>;
								}>;
							}>;
						}>;
					}>;
				}>;
			}>;
		};
		const nestedCell = doc.blocks
			.find((b) => b.id === "t0")
			?.rows?.[0]?.cells[0]?.blocks.find((b) => b.id === "t0:r0c0:t0")
			?.rows?.[0]?.cells[0];
		expect(nestedCell?.blocks).toHaveLength(2);
		expect(nestedCell?.blocks[1]?.runs?.[0]?.text).toBe("AddedSibling");
	});

	test("edit --at a nested-cell paragraph rewrites its text", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"t0:r0c0:t0:r0c0:p0",
			"--text",
			"Rewritten",
		);
		expect(result.exitCode).toBe(0);
		const ast = await runCli("read", docPath, "--ast");
		const doc = ast.parsed as {
			blocks: Array<{
				id: string;
				rows?: Array<{
					cells: Array<{
						blocks: Array<{
							id: string;
							rows?: Array<{
								cells: Array<{
									blocks: Array<{
										id: string;
										runs?: Array<{ text?: string }>;
									}>;
								}>;
							}>;
						}>;
					}>;
				}>;
			}>;
		};
		const text = doc.blocks
			.find((b) => b.id === "t0")
			?.rows?.[0]?.cells[0]?.blocks.find((b) => b.id === "t0:r0c0:t0")
			?.rows?.[0]?.cells[0]?.blocks[0]?.runs?.[0]?.text;
		expect(text).toBe("Rewritten");
	});

	test("insert --after a wrapper paragraph adds another nested table next to it", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"t0:r0c0:p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		expect(result.exitCode).toBe(0);
		const ast = await runCli("read", docPath, "--ast");
		const doc = ast.parsed as {
			blocks: Array<{
				id: string;
				rows?: Array<{ cells: Array<{ blocks: Array<{ type: string }> }> }>;
			}>;
		};
		const outerCellBlocks = doc.blocks.find((b) => b.id === "t0")?.rows?.[0]
			?.cells[0]?.blocks;
		// wrapper p, NEW inserted table, original nested table, trailing p
		expect(outerCellBlocks?.filter((b) => b.type === "table")).toHaveLength(2);
	});
});

describe("tables verbs with nested locators", () => {
	const NESTED_TABLE = "t0:r0c0:t0";

	test("insert-row appends to a nested table", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			NESTED_TABLE,
			"--cells",
			"newL,newR",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string; position: number };
		expect(ack.table).toBe(NESTED_TABLE);
		expect(ack.position).toBe(1);
	});

	test("insert-column extends a nested table", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"tables",
			"insert-column",
			docPath,
			"--at",
			NESTED_TABLE,
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string; position: number };
		expect(ack.table).toBe(NESTED_TABLE);
		expect(ack.position).toBe(2);
	});

	test("set-widths rewrites a nested table's grid", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"tables",
			"set-widths",
			docPath,
			"--at",
			NESTED_TABLE,
			"--widths",
			"600,600",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string; widths: number[] };
		expect(ack.table).toBe(NESTED_TABLE);
		expect(ack.widths).toEqual([600, 600]);
	});

	test("borders apply to a nested table", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"tables",
			"borders",
			docPath,
			"--at",
			NESTED_TABLE,
			"--style",
			"single",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string };
		expect(ack.table).toBe(NESTED_TABLE);
	});

	test("delete-row drops a row from a nested table (after insert)", async () => {
		const docPath = await makeDoc();
		await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			NESTED_TABLE,
			"--cells",
			"x,y",
		);
		const result = await runCli(
			"tables",
			"delete-row",
			docPath,
			"--at",
			`${NESTED_TABLE}:r1`,
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string; row: number };
		expect(ack.table).toBe(NESTED_TABLE);
		expect(ack.row).toBe(1);
	});

	test("delete-column trims a column from a nested table", async () => {
		const docPath = await makeDoc();
		const result = await runCli(
			"tables",
			"delete-column",
			docPath,
			"--at",
			`${NESTED_TABLE}:c1`,
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as { table: string; column: number };
		expect(ack.table).toBe(NESTED_TABLE);
		expect(ack.column).toBe(1);
	});

	test("merge / unmerge round-trip on a nested table", async () => {
		const docPath = await makeDoc();
		await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			NESTED_TABLE,
			"--cells",
			"x,y",
		);
		const merge = await runCli(
			"tables",
			"merge",
			docPath,
			"--at",
			`${NESTED_TABLE}:r0c0-r1c0`,
		);
		expect(merge.exitCode).toBe(0);
		expect((merge.parsed as { table: string }).table).toBe(NESTED_TABLE);

		const unmerge = await runCli(
			"tables",
			"unmerge",
			docPath,
			"--at",
			`${NESTED_TABLE}:r0c0`,
		);
		expect(unmerge.exitCode).toBe(0);
		expect((unmerge.parsed as { table: string }).table).toBe(NESTED_TABLE);
	});
});

// range-edit covers deleting a table as part of a pN-pM range; the direct
// `delete --at tN` path (whole-table) — and its tracked rejection — was not
// covered on its own.
describe("docx delete --at tN (whole table)", () => {
	async function tableDoc(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "Before");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"2",
		);
		await runCli("insert", docPath, "--after", "t0", "--text", "After");
		return docPath;
	}

	test("untracked delete removes the table and keeps surrounding paragraphs", async () => {
		const docPath = await tableDoc("del-table");
		const result = await runCli("delete", docPath, "--at", "t0");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const types = (
			read.parsed as { blocks: Array<{ type: string }> }
		).blocks.map((block) => block.type);
		expect(types).not.toContain("table");
		expect(
			types.filter((type) => type === "paragraph").length,
		).toBeGreaterThanOrEqual(2);
	});

	test("tracked deletion of a whole table is rejected", async () => {
		const docPath = await tableDoc("del-table-tracked");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("delete", docPath, "--at", "t0");
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { code: string }).code).toBe(
			"TRACKED_CHANGE_CONFLICT",
		);
	});
});

describe("docx tables — -o parallel write", () => {
	test("insert-row -o writes to the output and leaves the source byte-unchanged", async () => {
		const src = await newTableDoc("tables-o", 2, 2);
		const before = await Bun.file(src).bytes();
		const out = join(tempWorkspace("tables-o-out"), "out.docx");

		const result = await runCli(
			"tables",
			"insert-row",
			src,
			"--at",
			"t0",
			"-o",
			out,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(out);
		expect(await Bun.file(src).bytes()).toEqual(before);

		// The output gained a row; the source still has the original two.
		expect((await tableFromDoc(out))?.rows).toHaveLength(3);
		expect((await tableFromDoc(src))?.rows).toHaveLength(2);
	});
});

describe("docx tables — --dry-run previews without writing", () => {
	test("insert-row --dry-run reports the preview and leaves the file byte-unchanged", async () => {
		const docPath = await newTableDoc("tables-dry", 2, 2);
		const before = await Bun.file(docPath).bytes();

		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t0",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await Bun.file(docPath).bytes()).toEqual(before);
		expect((await tableFromDoc(docPath))?.rows).toHaveLength(2);
	});

	test("--dry-run still validates the locator (stale tN errors)", async () => {
		const docPath = await newTableDoc("tables-dry-stale", 2, 2);
		const result = await runCli(
			"tables",
			"insert-row",
			docPath,
			"--at",
			"t9",
			"--dry-run",
		);
		expect(result.exitCode).not.toBe(0);
	});
});

// Read-side visibility of table structure GFM collapses: uneven column widths
// and borders (a leading docx:table note), merge + shading (a per-cell docx:cell
// note). All deviation-only read-time hints; the importer drops them.
describe("table visual structure surfaces as docx:table / docx:cell hints", () => {
	async function tableDoc(
		label: string,
		extra: string[] = [],
	): Promise<string> {
		const workspace = tempWorkspace(label);
		const doc = join(workspace, "out.docx");
		await runCli("create", doc, "--text", "intro");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--table",
			"--rows",
			"2",
			"--cols",
			"3",
			...extra,
		);
		return doc;
	}

	test("uneven column widths surface as docx:table widths in inches", async () => {
		const doc = await tableDoc("tbl-widths", ["--widths", "1440,2880,5040"]);
		expect(await readMarkdown(doc)).toContain(
			'<!-- docx:table t0 widths="1,2,3.5in"',
		);
	});

	test("even columns emit no widths (deviation-only)", async () => {
		const doc = await tableDoc("tbl-even");
		expect(await readMarkdown(doc)).not.toContain("widths=");
	});

	test("borders: single (the universal default) is suppressed; none/double surface", async () => {
		const single = await tableDoc("tbl-single");
		expect(await readMarkdown(single)).not.toContain("borders=");

		const none = await tableDoc("tbl-none");
		await runCli("tables", "borders", none, "--at", "t0", "--style", "none");
		expect(await readMarkdown(none)).toContain('borders="none"');

		const dbl = await tableDoc("tbl-double");
		await runCli("tables", "borders", dbl, "--at", "t0", "--style", "double");
		expect(await readMarkdown(dbl)).toContain('borders="double"');
	});

	test("merge: a docx:cell note carries gridSpan; the bare cell locator stays", async () => {
		const doc = await tableDoc("tbl-merge-h");
		await runCli("edit", doc, "--at", "t0:r0c0:p0", "--text", "merged");
		await runCli("tables", "merge", doc, "--at", "t0:r0c0-r0c1");
		const md = await readMarkdown(doc);
		expect(md).toContain('<!-- docx:cell t0:r0c0 gridSpan="2" -->');
		expect(md).toContain("<!-- t0:r0c0:p0 -->"); // bare locator unchanged
	});

	test("a vertical merge surfaces vMerge=restart + continue", async () => {
		const doc = await tableDoc("tbl-merge-v");
		await runCli("tables", "merge", doc, "--at", "t0:r0c0-r1c0");
		const md = await readMarkdown(doc);
		expect(md).toMatch(/docx:cell[^>]*vMerge="restart"/);
		expect(md).toMatch(/docx:cell[^>]*vMerge="continue"/);
	});

	test("cell shading surfaces as a docx:cell note + TableCell.shading in --ast", async () => {
		const doc = await tableDoc("tbl-shade");
		// No CLI verb sets shading yet — inject <w:shd w:fill> into the first cell.
		const zip = await JSZip.loadAsync(await Bun.file(doc).bytes());
		const entry = zip.file("word/document.xml");
		if (!entry) throw new Error("document.xml missing");
		const xml = (await entry.async("string")).replace(
			/<w:tc>(\s*)(<w:tcPr>)?/,
			(_m, ws2, tcPrOpen) =>
				tcPrOpen
					? `<w:tc>${ws2}<w:tcPr><w:shd w:val="clear" w:fill="FFE699"/>`
					: `<w:tc>${ws2}<w:tcPr><w:shd w:val="clear" w:fill="FFE699"/></w:tcPr>`,
		);
		zip.file("word/document.xml", xml);
		await Bun.write(doc, await zip.generateAsync({ type: "uint8array" }));

		expect(await readMarkdown(doc)).toMatch(/docx:cell[^>]*shading="FFE699"/);
		const table = await tableFromDoc(doc);
		expect(table?.rows[0]?.cells[0]?.shading).toBe("FFE699");
	});

	test("borders summary appears in --ast (Table.borders)", async () => {
		const doc = await tableDoc("tbl-borders-ast");
		await runCli("tables", "borders", doc, "--at", "t0", "--style", "double");
		const table = await tableFromDoc(doc);
		expect(table?.borders).toBe("double");
	});
});
