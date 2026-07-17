import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import { readDocumentXml, readMarkdown } from "./helpers";

let workspace: string;
let docPath: string;

beforeEach(async () => {
	workspace = tempWorkspace("raw");
	docPath = join(workspace, "doc.docx");
	await runCli("create", docPath, "--text", "hello world");
});

type MintedAck = {
	ok: boolean;
	operation: string;
	locators: string[];
	warnings?: string[];
};

describe("raw get", () => {
	test("prints the exact XML of a block", async () => {
		const result = await runCli("raw", "get", docPath, "--at", "p0");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.trim().startsWith("<w:p")).toBe(true);
		expect(result.stdout).toContain("hello world");
	});

	test("--json wraps locator and xml", async () => {
		const result = await runCli("raw", "get", docPath, "--at", "s0", "--json");
		const parsed = result.parsed as { locator: string; xml: string };
		expect(parsed.locator).toBe("s0");
		expect(parsed.xml.startsWith("<w:sectPr")).toBe(true);
	});

	test("unknown locator exits 3", async () => {
		const result = await runCli("raw", "get", docPath, "--at", "p99");
		expect(result.exitCode).toBe(3);
	});

	test("--at ftrN prints the footer's <w:ftr> (the id from footers list / read)", async () => {
		await runCli("footers", "set", docPath, "--text", "Confidential");
		const result = await runCli("raw", "get", docPath, "--at", "ftr0");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.trim().startsWith("<w:ftr")).toBe(true);
		expect(result.stdout).toContain("Confidential");
	});

	test("--at hdrN prints the header's <w:hdr>", async () => {
		await runCli("headers", "set", docPath, "--text", "Title");
		const result = await runCli(
			"raw",
			"get",
			docPath,
			"--at",
			"hdr0",
			"--json",
		);
		const parsed = result.parsed as { locator: string; xml: string };
		expect(parsed.locator).toBe("hdr0");
		expect(parsed.xml.startsWith("<w:hdr")).toBe(true);
	});

	test("an unknown marginal id exits 3 with a list/read hint", async () => {
		const result = await runCli("raw", "get", docPath, "--at", "ftr9");
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { hint: string }).hint).toContain("footers list");
	});
});

describe("raw insert", () => {
	test("mints the new block's locator and survives the write-read loop", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			'<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/></w:pPr><w:r><w:t>D</w:t></w:r></w:p>',
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as MintedAck;
		expect(ack.locators).toEqual(["p1"]);

		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("<!-- docx:p p1 raw -->");
		const ast = await runCli("read", docPath, "--ast");
		const blocks = (ast.parsed as { blocks: { id: string; rawXml?: true }[] })
			.blocks;
		expect(blocks.find((block) => block.id === "p1")?.rawXml).toBe(true);
		expect(await readDocumentXml(docPath)).toContain("w:framePr");
	});

	test("inserting into a cell mints a chained locator", async () => {
		await runCli(
			"tables",
			"create",
			docPath,
			"--after",
			"p0",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"t0:r0c0:p0",
			"--xml",
			"<w:p><w:r><w:t>in cell</w:t></w:r></w:p>",
		);
		const ack = result.parsed as MintedAck;
		expect(ack.locators).toEqual(["t0:r0c0:p1"]);
	});

	test("a table left last in a cell gets a trailing paragraph appended", async () => {
		await runCli(
			"tables",
			"create",
			docPath,
			"--after",
			"p0",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"t0:r0c0:p0",
			"--xml",
			"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as MintedAck;
		expect(
			ack.warnings?.some((warning) =>
				warning.includes("appended an empty paragraph"),
			),
		).toBe(true);
	});

	test("--xml-file reads the fragment from disk", async () => {
		const fragmentPath = join(workspace, "fragment.xml");
		await Bun.write(fragmentPath, "<w:p><w:r><w:t>from file</w:t></w:r></w:p>");
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--at-end",
			"--xml-file",
			fragmentPath,
		);
		expect(result.exitCode).toBe(0);
		expect(await readMarkdown(docPath)).toContain("from file");
	});

	test("--track is rejected", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--track",
			"--xml",
			"<w:p/>",
		);
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe(
			"TRACKED_CHANGE_CONFLICT",
		);
	});

	test("document tracking ON applies untracked with a note and an audit comment", async () => {
		await runCli("track-changes", "on", docPath);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			"<w:p><w:r><w:t>untracked</w:t></w:r></w:p>",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as MintedAck;
		expect(
			ack.warnings?.some((warning) => warning.includes("NOT tracked")),
		).toBe(true);
		const comments = await runCli("comments", "list", docPath);
		const bodies = (comments.parsed as { text: string }[]).map(
			(comment) => comment.text,
		);
		expect(bodies.some((body) => body.includes("[docx-cli]"))).toBe(true);
	});

	test("--dry-run runs the gates but writes nothing", async () => {
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--dry-run",
			"--xml",
			"<w:p><w:r><w:t>preview</w:t></w:r></w:p>",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await readDocumentXml(docPath)).toBe(before);
	});
});

describe("raw insert gates", () => {
	test("malformed XML exits 2 with INVALID_XML", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			"<w:p><w:r><w:t>a",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("INVALID_XML");
	});

	test("non-addressable root exits 2", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			"<w:sdt><w:sdtContent/></w:sdt>",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { error: string }).error).toContain(
			"not addressable",
		);
	});

	test("schema gate: a deep bogus element fails VALIDATION_FAILED and writes nothing", async () => {
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			"<w:p><w:r><w:bogusChild/><w:t>x</w:t></w:r></w:p>",
		);
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe("VALIDATION_FAILED");
		expect((result.parsed as { error: string }).error).toContain("bogusChild");
		expect(await readDocumentXml(docPath)).toBe(before);
	});

	test("--no-validate skips only the schema gate", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--no-validate",
			"--xml",
			"<w:p><w:r><w:bogusChild/><w:t>x</w:t></w:r></w:p>",
		);
		expect(result.exitCode).toBe(0);
	});

	test("pre-existing schema errors never block a clean raw edit (baseline diff)", async () => {
		await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--no-validate",
			"--xml",
			"<w:p><w:r><w:bogusChild/><w:t>x</w:t></w:r></w:p>",
		);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--at-end",
			"--xml",
			"<w:p><w:r><w:t>clean</w:t></w:r></w:p>",
		);
		expect(result.exitCode).toBe(0);
	});
});

describe("raw replace", () => {
	test("get → modify → replace round-trips", async () => {
		const got = await runCli("raw", "get", docPath, "--at", "p0");
		const patched = got.stdout.trim().replace("hello world", "patched text");
		const result = await runCli(
			"raw",
			"replace",
			docPath,
			"--at",
			"p0",
			"--xml",
			patched,
		);
		expect(result.exitCode).toBe(0);
		expect(await readMarkdown(docPath)).toContain("patched text");
	});

	test("1-for-N replacement mints every replacement locator", async () => {
		const result = await runCli(
			"raw",
			"replace",
			docPath,
			"--at",
			"p0",
			"--xml",
			"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>",
		);
		const ack = result.parsed as MintedAck;
		expect(ack.locators).toEqual(["p0", "p1"]);
	});

	test("replacing sN takes exactly one w:sectPr (the section patch loop)", async () => {
		const result = await runCli(
			"raw",
			"replace",
			docPath,
			"--at",
			"s0",
			"--xml",
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:lnNumType w:countBy="1"/><w:cols w:space="720"/></w:sectPr>',
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as MintedAck).locators).toEqual(["s0"]);
		expect(await readDocumentXml(docPath)).toContain("w:lnNumType");

		const wrongRoot = await runCli(
			"raw",
			"replace",
			docPath,
			"--at",
			"s0",
			"--xml",
			"<w:p/>",
		);
		expect(wrongRoot.exitCode).toBe(2);
	});
});

describe("raw batch", () => {
	test("insert --batch: stacked entries keep entry order from one read", async () => {
		const batchPath = join(workspace, "batch.jsonl");
		await Bun.write(
			batchPath,
			`${JSON.stringify({ after: "p0", xml: "<w:p><w:r><w:t>first</w:t></w:r></w:p>" })}\n${JSON.stringify({ after: "p0", xml: "<w:p><w:r><w:t>second</w:t></w:r></w:p>" })}\n`,
		);
		const result = await runCli("raw", "insert", docPath, "--batch", batchPath);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as MintedAck).locators).toEqual(["p1", "p2"]);
		const markdown = await readMarkdown(docPath);
		expect(markdown.indexOf("first")).toBeLessThan(markdown.indexOf("second"));
	});

	test("replace --batch: locators address the doc AS READ; duplicates reject", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"second paragraph",
		);
		const batchPath = join(workspace, "replace.jsonl");
		await Bun.write(
			batchPath,
			`${JSON.stringify({ at: "p0", xml: "<w:p><w:r><w:t>ONE</w:t></w:r></w:p>" })}\n${JSON.stringify({ at: "p1", xml: "<w:p><w:r><w:t>TWO</w:t></w:r></w:p>" })}\n`,
		);
		const result = await runCli(
			"raw",
			"replace",
			docPath,
			"--batch",
			batchPath,
		);
		expect(result.exitCode).toBe(0);
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("ONE");
		expect(markdown).toContain("TWO");

		const dupPath = join(workspace, "dup.jsonl");
		await Bun.write(
			dupPath,
			`${JSON.stringify({ at: "p0", xml: "<w:p/>" })}\n${JSON.stringify({ at: "p0", xml: "<w:p/>" })}\n`,
		);
		const dup = await runCli("raw", "replace", docPath, "--batch", dupPath);
		expect(dup.exitCode).toBe(2);
	});

	test("--batch with a single-shot flag rejects", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--batch",
			"whatever.jsonl",
			"--after",
			"p0",
		);
		expect(result.exitCode).toBe(2);
	});

	test("a failing entry means nothing is written", async () => {
		const before = await readDocumentXml(docPath);
		const batchPath = join(workspace, "bad.jsonl");
		await Bun.write(
			batchPath,
			`${JSON.stringify({ after: "p0", xml: "<w:p><w:r><w:t>fine</w:t></w:r></w:p>" })}\n${JSON.stringify({ after: "p0", xml: "<w:p><w:r><w:t>broken" })}\n`,
		);
		const result = await runCli("raw", "insert", docPath, "--batch", batchPath);
		expect(result.exitCode).toBe(2);
		expect(await readDocumentXml(docPath)).toBe(before);
	});
});

describe("raw — relationships", () => {
	const HYPERLINK_TYPE =
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

	test("insert mints an rId, get reads it back, replace retargets it", async () => {
		const inserted = await runCli(
			"raw",
			"insert",
			docPath,
			"--xml",
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://old.example" TargetMode="External"/>`,
		);
		expect(inserted.exitCode).toBe(0);
		const rId = (inserted.parsed as { relationships: string[] })
			.relationships[0] as string;
		expect(rId).toMatch(/^rId\d+$/);

		const got = await runCli("raw", "get", docPath, "--at", rId);
		expect(got.exitCode).toBe(0);
		expect(got.stdout).toContain('Target="https://old.example"');

		const replaced = await runCli(
			"raw",
			"replace",
			docPath,
			"--at",
			rId,
			"--xml",
			`<Relationship Id="${rId}" Type="${HYPERLINK_TYPE}" Target="https://new.example" TargetMode="External"/>`,
		);
		expect(replaced.exitCode).toBe(0);
		expect(
			(replaced.parsed as { relationships: string[] }).relationships,
		).toEqual([rId]);

		const after = await runCli("raw", "get", docPath, "--at", rId, "--json");
		expect((after.parsed as { xml: string }).xml).toContain(
			"https://new.example",
		);
	});

	test("get --at rels prints the whole part (the discovery path)", async () => {
		const result = await runCli("raw", "get", docPath, "--at", "rels");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("<Relationships");
		expect(result.stdout).toContain("styles.xml");
	});

	test("the rel-then-body loop: mint an rId, then reference it from a raw block", async () => {
		const inserted = await runCli(
			"raw",
			"insert",
			docPath,
			"--xml",
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://docs.example" TargetMode="External"/>`,
		);
		const rId = (inserted.parsed as { relationships: string[] })
			.relationships[0] as string;
		const block = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			`<w:p><w:hyperlink r:id="${rId}"><w:r><w:t>docs</w:t></w:r></w:hyperlink></w:p>`,
		);
		expect(block.exitCode).toBe(0);
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("https://docs.example");
	});

	test("placement flags reject on a Relationship fragment", async () => {
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml",
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://x" TargetMode="External"/>`,
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { error: string }).error).toContain("unordered");
	});

	test("a missing rId is exit 3 with the rels-listing hint", async () => {
		for (const args of [
			["raw", "get", docPath, "--at", "rId999"],
			[
				"raw",
				"replace",
				docPath,
				"--at",
				"rId999",
				"--xml",
				'<Relationship Type="t" Target="https://x" TargetMode="External"/>',
			],
		]) {
			const result = await runCli(...(args as [string, ...string[]]));
			expect(result.exitCode).toBe(3);
			expect((result.parsed as { hint: string }).hint).toContain("--at rels");
		}
	});

	test("--dry-run gates without writing", async () => {
		const before = await runCli("raw", "get", docPath, "--at", "rels");
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--dry-run",
			"--xml",
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://x" TargetMode="External"/>`,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		const after = await runCli("raw", "get", docPath, "--at", "rels");
		expect(after.stdout).toBe(before.stdout);
	});

	test("Relationship fragments and rId targets aren't batchable", async () => {
		const insertBatch = join(workspace, "rel-insert.jsonl");
		await Bun.write(
			insertBatch,
			`${JSON.stringify({ after: "p0", xml: '<Relationship Type="t" Target="https://x" TargetMode="External"/>' })}\n`,
		);
		const insertResult = await runCli(
			"raw",
			"insert",
			docPath,
			"--batch",
			insertBatch,
		);
		expect(insertResult.exitCode).toBe(2);
		expect((insertResult.parsed as { error: string }).error).toContain(
			"batchable",
		);

		const replaceBatch = join(workspace, "rel-replace.jsonl");
		await Bun.write(
			replaceBatch,
			`${JSON.stringify({ at: "rId1", xml: '<Relationship Type="t" Target="https://x" TargetMode="External"/>' })}\n`,
		);
		const replaceResult = await runCli(
			"raw",
			"replace",
			docPath,
			"--batch",
			replaceBatch,
		);
		expect(replaceResult.exitCode).toBe(2);
		expect((replaceResult.parsed as { error: string }).error).toContain(
			"batchable",
		);
	});
});

describe("raw — parts", () => {
	const OLE_TYPE = "application/vnd.openxmlformats-officedocument.oleObject";

	test("the embedded-object loop: part add → relationship → get bytes back", async () => {
		const blob = join(workspace, "blob.bin");
		await Bun.write(blob, new Uint8Array([1, 2, 3, 4, 5]));
		const added = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/embeddings/object1.bin",
			"--from",
			blob,
			"--content-type",
			OLE_TYPE,
		);
		expect(added.exitCode).toBe(0);
		const ack = added.parsed as {
			part: string;
			contentType: string;
			warnings: string[];
		};
		expect(ack.part).toBe("word/embeddings/object1.bin");
		expect(ack.contentType).toBe(OLE_TYPE);
		expect(ack.warnings[0]).toContain("inert until a relationship");

		// The internal-Target gate that used to reject now passes: the part exists.
		const rel = await runCli(
			"raw",
			"insert",
			docPath,
			"--xml",
			'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/object1.bin"/>',
		);
		expect(rel.exitCode).toBe(0);

		const out = join(workspace, "out.bin");
		const got = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"word/embeddings/object1.bin",
			"--to-file",
			out,
		);
		expect(got.exitCode).toBe(0);
		expect(new Uint8Array(await Bun.file(out).arrayBuffer())).toEqual(
			new Uint8Array([1, 2, 3, 4, 5]),
		);
	});

	test("part list shows every part with its content type; parts aren't locators", async () => {
		const result = await runCli("raw", "part", "list", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("word/document.xml");
		expect(result.stdout).toContain("wordprocessingml.document.main+xml");

		for (const at of ["parts", "part:word/document.xml"]) {
			const guarded = await runCli("raw", "get", docPath, "--at", at);
			expect(guarded.exitCode).toBe(2);
			expect((guarded.parsed as { hint: string }).hint).toContain(
				"docx raw part",
			);
		}
	});

	test("xml part: add via the extension Default, whole-part replace, read back", async () => {
		const added = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"customXml/item1.xml",
			"--xml",
			'<data xmlns="urn:test"><v>1</v></data>',
		);
		expect(added.exitCode).toBe(0);
		const replaced = await runCli(
			"raw",
			"part",
			"replace",
			docPath,
			"--name",
			"customXml/item1.xml",
			"--xml",
			'<data xmlns="urn:test"><v>2</v></data>',
		);
		expect(replaced.exitCode).toBe(0);
		const got = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"customXml/item1.xml",
		);
		expect(got.stdout).toContain("<v>2</v>");
	});

	test("a WML-rooted part is schema-gated on add", async () => {
		const bad = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/extra.xml",
			"--xml",
			'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:bogusChild/></w:r></w:p></w:hdr>',
		);
		expect(bad.exitCode).toBe(1);
		expect((bad.parsed as { code: string }).code).toBe("VALIDATION_FAILED");
		const good = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/extra.xml",
			"--xml",
			'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:hdr>',
		);
		expect(good.exitCode).toBe(0);
	});

	test("guards: existing part, unknown extension, modeled replace, missing targets", async () => {
		const existing = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/styles.xml",
			"--xml",
			"<w:styles/>",
		);
		expect(existing.exitCode).toBe(2);
		expect((existing.parsed as { error: string }).error).toContain(
			"already exists",
		);

		const blob = join(workspace, "ct.bin");
		await Bun.write(blob, new Uint8Array([9]));
		const noCt = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/embeddings/x.bin",
			"--from",
			blob,
		);
		expect(noCt.exitCode).toBe(2);
		expect((noCt.parsed as { error: string }).error).toContain(
			"--content-type",
		);

		const bodyPart = await runCli(
			"raw",
			"part",
			"replace",
			docPath,
			"--name",
			"word/document.xml",
			"--xml",
			"<w:document/>",
		);
		expect(bodyPart.exitCode).toBe(2);
		expect((bodyPart.parsed as { hint: string }).hint).toContain(
			"block locators",
		);

		const missingReplace = await runCli(
			"raw",
			"part",
			"replace",
			docPath,
			"--name",
			"customXml/none.xml",
			"--xml",
			"<a/>",
		);
		expect(missingReplace.exitCode).toBe(3);
		const missingGet = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"customXml/none.xml",
		);
		expect(missingGet.exitCode).toBe(3);
	});

	test("a binary part get without --to-file is a usage error", async () => {
		const blob = join(workspace, "bin2.bin");
		await Bun.write(blob, new Uint8Array([7]));
		await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--name",
			"word/embeddings/bin2.bin",
			"--from",
			blob,
			"--content-type",
			OLE_TYPE,
		);
		const result = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"word/embeddings/bin2.bin",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { error: string }).error).toContain("--to-file");
	});

	test("--dry-run gates without writing", async () => {
		const result = await runCli(
			"raw",
			"part",
			"add",
			docPath,
			"--dry-run",
			"--name",
			"customXml/dry.xml",
			"--xml",
			'<data xmlns="urn:test"/>',
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		const after = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"customXml/dry.xml",
		);
		expect(after.exitCode).toBe(3);
	});
});

describe("raw edit", () => {
	test("patches a block in place (get → modify → replace in one call)", async () => {
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			"p0",
			"--find",
			"hello world",
			"--with",
			"patched world",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as MintedAck;
		expect(ack.operation).toBe("raw.edit");
		expect(ack.locators).toEqual(["p0"]);
		expect(
			ack.warnings?.some((warning) => warning.includes("1 occurrence")),
		).toBe(true);
		expect(await readMarkdown(docPath)).toContain("patched world");
	});

	test("patches a sectPr without shipping the whole element", async () => {
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			"s0",
			"--find",
			"<w:cols",
			"--with",
			'<w:lnNumType w:countBy="1" w:restart="continuous"/><w:cols',
		);
		expect(result.exitCode).toBe(0);
		const got = await runCli("raw", "get", docPath, "--at", "s0");
		expect(got.stdout).toContain("<w:lnNumType");
		const valid = await runCli("validate", docPath);
		expect(valid.exitCode).toBe(0);
	});

	test("patches a relationship target", async () => {
		const inserted = await runCli(
			"raw",
			"insert",
			docPath,
			"--xml",
			'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/>',
		);
		const rId = (inserted.parsed as { relationships: string[] })
			.relationships[0] as string;
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			rId,
			"--find",
			"https://old.example",
			"--with",
			"https://new.example",
		);
		expect(result.exitCode).toBe(0);
		const got = await runCli("raw", "get", docPath, "--at", rId);
		expect(got.stdout).toContain("https://new.example");
	});

	test("patches an unmodeled part (pkg route) and a view-backed part (view route)", async () => {
		// pkg route: fontTable.xml — swap a font name, still schema-valid.
		const fontTable = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"word/fontTable.xml",
		);
		expect(fontTable.stdout).toContain('w:name="Calibri"');
		const pkgEdit = await runCli(
			"raw",
			"part",
			"edit",
			docPath,
			"--name",
			"word/fontTable.xml",
			"--find",
			'w:name="Calibri"',
			"--with",
			'w:name="Carlito"',
		);
		expect(pkgEdit.exitCode).toBe(0);
		const afterPkg = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"word/fontTable.xml",
		);
		expect(afterPkg.stdout).toContain("Carlito");

		// view route: styles.xml — alias the Normal style; must survive the
		// StylesView reparse + serialization on save.
		const viewEdit = await runCli(
			"raw",
			"part",
			"edit",
			docPath,
			"--name",
			"word/styles.xml",
			"--find",
			'<w:name w:val="Normal"/>',
			"--with",
			'<w:name w:val="Normal"/><w:aliases w:val="Base"/>',
		);
		expect(viewEdit.exitCode).toBe(0);
		const afterView = await runCli(
			"raw",
			"part",
			"get",
			docPath,
			"--name",
			"word/styles.xml",
		);
		expect(afterView.stdout).toContain('<w:aliases w:val="Base"/>');
		const valid = await runCli("validate", docPath);
		expect(valid.exitCode).toBe(0);
	});

	test("zero matches is MATCH_NOT_FOUND and writes nothing", async () => {
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			"p0",
			"--find",
			"no such text anywhere",
			"--with",
			"x",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("MATCH_NOT_FOUND");
		expect((result.parsed as { hint: string }).hint).toContain("raw get");
		expect(await readDocumentXml(docPath)).toBe(before);
	});

	test("a patch that breaks the schema is rejected, nothing written", async () => {
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			"p0",
			"--find",
			'<w:t xml:space="preserve">hello world</w:t>',
			"--with",
			'<w:bogusChild/><w:t xml:space="preserve">hello world</w:t>',
		);
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe("VALIDATION_FAILED");
		expect(await readDocumentXml(docPath)).toBe(before);
	});

	test("--dry-run previews with the occurrence count, writes nothing", async () => {
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"edit",
			docPath,
			"--at",
			"p0",
			"--dry-run",
			"--find",
			"hello",
			"--with",
			"goodbye",
		);
		expect(result.exitCode).toBe(0);
		const parsed = result.parsed as { dryRun: boolean; warnings: string[] };
		expect(parsed.dryRun).toBe(true);
		expect(
			parsed.warnings.some((warning) => warning.includes("1 occurrence")),
		).toBe(true);
		expect(await readDocumentXml(docPath)).toBe(before);
	});
});
