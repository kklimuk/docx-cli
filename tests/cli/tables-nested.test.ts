import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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
