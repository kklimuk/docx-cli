import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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
};

type TableBlock = {
	id: string;
	type: "table";
	grid: number[];
	width?: { value: number; unit: string };
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
