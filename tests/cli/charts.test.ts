import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import JSZip from "jszip";
import { runCli } from "./harness";
import { freshFixture, readMarkdown } from "./helpers";

const FIXTURES = join(import.meta.dir, "..", "fixtures");
const CHARTS = join(FIXTURES, "charts.docx");
const STRICT = join(FIXTURES, "strict-profile.docx");

describe("docx read — native charts", () => {
	test("each chart keeps its token and gains a docx:chart hint led by chartN", async () => {
		const markdown = await readMarkdown(CHARTS);
		expect(markdown).toContain(
			'`[chart]` <!-- docx:chart chart0 type="column" title="Revenue by Quarter" series="2" points="4" size="6x3.5in" data="none" -->',
		);
		// chartN numbering continues into table cells; a combo joins its types.
		expect(markdown).toContain(
			'| `[chart]` <!-- docx:chart chart1 type="column+line" series="2" points="4" size="6x3.5in" data="none" --> <!-- t0:r1c0:p0 --> |',
		);
	});

	test("an embedded-workbook chart omits data= (deviation-only), Strict namespaces included", async () => {
		const markdown = await readMarkdown(STRICT);
		expect(markdown).toContain(
			'`[chart]` <!-- docx:chart chart0 type="column" series="3" points="4" size="6x3.5in" -->',
		);
		expect(markdown).toContain("`[smartart]`");
	});

	test("--ast carries the parsed chart (everything read shows, and more)", async () => {
		const result = await runCli("read", CHARTS, "--ast");
		const ast = result.parsed as {
			blocks: { runs?: { type: string; id?: string; series?: unknown[] }[] }[];
		};
		const chart = ast.blocks
			.flatMap((block) => block.runs ?? [])
			.find((run) => run.type === "chart");
		expect(chart).toMatchObject({
			kind: "chart",
			id: "chart0",
			chartTypes: ["column"],
			title: "Revenue by Quarter",
			dataSource: "none",
		});
		expect(chart?.series).toHaveLength(2);
	});

	test("editing the paragraph around a chart keeps the chart and its part", async () => {
		const docPath = await freshFixture("charts-edit", CHARTS);
		const before = await JSZip.loadAsync(await Bun.file(docPath).arrayBuffer());
		const chartXml = await before
			.file("word/charts/chart1.xml")
			?.async("string");
		const result = await runCli("edit", docPath, "--at", "p2", "Revenue: ");
		expect(result.exitCode).toBe(0);
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain('docx:chart chart0 type="column"');
		expect(markdown).toContain("Revenue:");
		const after = await JSZip.loadAsync(await Bun.file(docPath).arrayBuffer());
		expect(await after.file("word/charts/chart1.xml")?.async("string")).toBe(
			chartXml,
		);
	});
});

describe("docx charts list", () => {
	test("prints a bare JSON array whose id is the --at handle", async () => {
		const result = await runCli("charts", "list", CHARTS);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([
			{
				id: "chart0",
				at: "p2",
				type: "column",
				title: "Revenue by Quarter",
				series: ["West", "East"],
				points: 4,
				data: "none",
			},
			{
				id: "chart1",
				at: "t0:r1c0:p0",
				type: "column+line",
				series: ["Units", "Margin"],
				points: 4,
				data: "none",
			},
		]);
	});

	test("an embedded chart names its workbook part", async () => {
		const result = await runCli("charts", "list", STRICT);
		expect(result.parsed).toEqual([
			{
				id: "chart0",
				at: "p1",
				type: "column",
				series: ["Series 1", "Series 2", "Series 3"],
				points: 4,
				data: "embedded",
				workbook: "word/embeddings/Microsoft_Excel_Worksheet.xlsx",
			},
		]);
	});

	test("a document without charts lists none", async () => {
		const result = await runCli(
			"charts",
			"list",
			join(FIXTURES, "minimal.docx"),
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([]);
	});
});

describe("docx charts extract", () => {
	test("CSV by default: header, one row per category, a gap is an empty cell", async () => {
		const result = await runCli("charts", "extract", CHARTS, "--at", "chart0");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe(
			"category,West,East\nQ1,120,90\nQ2,135,110\nQ3,150,\nQ4,170,140\n",
		);
	});

	test("--json prints the full series objects", async () => {
		const result = await runCli(
			"charts",
			"extract",
			CHARTS,
			"--at",
			"chart1",
			"--json",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			id: string;
			series: { name: string; values: number[]; valuesRef: string }[];
		};
		expect(payload.id).toBe("chart1");
		expect(payload.series[1]).toMatchObject({
			name: "Margin",
			values: [0.21, 0.24, 0.22, 0.27],
			valuesRef: "Sheet1!$C$2:$C$5",
		});
	});

	test("an unknown chart id is CHART_NOT_FOUND (exit 3)", async () => {
		const result = await runCli("charts", "extract", CHARTS, "--at", "chart9");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "CHART_NOT_FOUND" });
	});

	test("--at is required", async () => {
		const result = await runCli("charts", "extract", CHARTS);
		expect(result.exitCode).toBe(2);
	});
});
