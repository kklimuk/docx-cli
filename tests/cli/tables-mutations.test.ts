import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

type Cell = {
	blocks: Array<{ runs?: Array<{ text?: string }> }>;
	gridSpan?: number;
	vMerge?: "restart" | "continue";
	trackedChange?: { kind: string };
};

type TableBlock = {
	id: string;
	type: "table";
	grid: number[];
	width?: { value: number; unit: string };
	rows: Array<{ cells: Cell[]; trackedChange?: { kind: string } }>;
};

async function table(docPath: string): Promise<TableBlock> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as { blocks: Array<{ type: string }> };
	const found = doc.blocks.find((block) => block.type === "table");
	if (!found) throw new Error("no table in document");
	return found as TableBlock;
}

function cellText(cell: Cell): string {
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

	async function listKinds(docPath: string): Promise<string[]> {
		const result = await runCli("track-changes", "list", docPath);
		return (result.parsed as Array<{ kind: string }>).map((c) => c.kind);
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
		expect(await listKinds(docPath)).toContain("rowIns");
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
		const kinds = await listKinds(docPath);
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
		const kinds = await listKinds(docPath);
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
		expect(await listKinds(docPath)).toHaveLength(0);
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
		expect(await listKinds(docPath)).toHaveLength(0);
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
