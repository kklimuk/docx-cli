import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx track-changes", () => {
	test("on creates settings.xml and registers it", async () => {
		const workspace = tempWorkspace("track");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");

		const result = await runCli("track-changes", docPath, "on");
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes",
			mode: "on",
			previouslyOn: false,
		});

		const proc = Bun.spawn(["unzip", "-p", docPath, "word/settings.xml"], {
			stdout: "pipe",
		});
		const xml = await new Response(proc.stdout).text();
		await proc.exited;
		expect(xml).toContain("<w:trackChanges/>");
		expect(xml).toContain("xmlns:w=");
	});

	test("off removes the trackChanges element", async () => {
		const workspace = tempWorkspace("untrack");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("track-changes", docPath, "off");
		expect(result.parsed).toMatchObject({
			mode: "off",
			previouslyOn: true,
		});

		const proc = Bun.spawn(["unzip", "-p", docPath, "word/settings.xml"], {
			stdout: "pipe",
		});
		const xml = await new Response(proc.stdout).text();
		await proc.exited;
		expect(xml).not.toContain("<w:trackChanges/>");
	});

	test("rejects invalid mode", async () => {
		const workspace = tempWorkspace("invalid");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");
		const result = await runCli("track-changes", docPath, "maybe");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("list returns every <w:ins>/<w:del> with metadata", async () => {
		const result = await runCli(
			"track-changes",
			"list",
			"tests/fixtures/tracked-changes.docx",
		);
		expect(result.exitCode).toBe(0);
		const changes = result.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			date: string;
			blockId: string;
			text: string;
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "ins",
			author: "eng-dept",
			date: "2014-06-25T10:40:00Z",
			blockId: "p0",
			text: "two exciting ",
		});
	});

	test("list on a doc without tracked changes returns []", async () => {
		const result = await runCli(
			"track-changes",
			"list",
			"tests/fixtures/minimal.docx",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([]);
	});
});

describe("docx track-changes accept / reject", () => {
	async function freshFixture(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/tracked-changes.docx"));
		return docPath;
	}

	async function flatText(docPath: string): Promise<string> {
		const read = await runCli("read", docPath, "--ast");
		const blocks = (
			read.parsed as {
				blocks: Array<{ runs?: Array<{ type: string; text: string }> }>;
			}
		).blocks;
		return blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
	}

	async function buildMixed(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "doc.docx");
		await runCli("create", docPath, "--text", "The quick brown fox jumps.");
		await runCli("track-changes", docPath, "on");
		// Yields a <w:del> for "quick brown " and a <w:ins> for "old slow ".
		await runCli("replace", docPath, "quick brown ", "old slow ");
		return docPath;
	}

	test("accept --at tcN unwraps an insertion (text becomes plain)", async () => {
		const docPath = await freshFixture("accept-at-ins");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.accept",
			applied: [{ id: "tc0", kind: "ins", action: "unwrap" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
		expect(await flatText(docPath)).toBe(
			"This is a text with two exciting insertions.",
		);
	});

	test("reject --at tcN on an insertion deletes the inserted text", async () => {
		const docPath = await freshFixture("reject-at-ins");
		const result = await runCli(
			"track-changes",
			"reject",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.reject",
			applied: [{ id: "tc0", kind: "ins", action: "delete" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
		expect(await flatText(docPath)).toBe("This is a text with insertions.");
	});

	test("accept --all on mixed ins+del yields the post-accept text", async () => {
		const docPath = await buildMixed("accept-all-mixed");
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);
		const applied = (
			result.parsed as {
				applied: Array<{ id: string; kind: string; action: string }>;
			}
		).applied;
		expect(applied).toHaveLength(2);
		// Reverse-order processing: tc1 (ins) before tc0 (del).
		expect(applied.map((a) => a.id).sort()).toEqual(["tc0", "tc1"]);
		expect(applied.find((a) => a.kind === "del")?.action).toBe("delete");
		expect(applied.find((a) => a.kind === "ins")?.action).toBe("unwrap");

		expect(await flatText(docPath)).toBe("The old slow fox jumps.");
		expect((await runCli("track-changes", "list", docPath)).parsed).toEqual([]);
	});

	test("reject --all on mixed ins+del reverts to the baseline text", async () => {
		const docPath = await buildMixed("reject-all-mixed");
		const result = await runCli("track-changes", "reject", docPath, "--all");
		expect(result.exitCode).toBe(0);
		const applied = (
			result.parsed as { applied: Array<{ kind: string; action: string }> }
		).applied;
		expect(applied.find((a) => a.kind === "ins")?.action).toBe("delete");
		expect(applied.find((a) => a.kind === "del")?.action).toBe("unwrap");

		expect(await flatText(docPath)).toBe("The quick brown fox jumps.");
		expect((await runCli("track-changes", "list", docPath)).parsed).toEqual([]);
	});

	test("--dry-run reports the plan without mutating the file", async () => {
		const docPath = await freshFixture("dry-run-accept");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			dryRun: true,
			operation: "track-changes.accept",
			applied: [{ id: "tc0", kind: "ins", action: "unwrap" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toMatchObject([{ id: "tc0" }]);
	});

	test("--at tcN with unknown id returns TRACKED_CHANGE_NOT_FOUND", async () => {
		const docPath = await freshFixture("missing-id");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc99",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "TRACKED_CHANGE_NOT_FOUND",
		});
	});

	test("--at and --all together are rejected", async () => {
		const docPath = await freshFixture("conflict");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--all",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("missing --at and --all is rejected", async () => {
		const docPath = await freshFixture("missing-target");
		const result = await runCli("track-changes", "accept", docPath);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--all on a doc with no tracked changes succeeds with empty applied", async () => {
		const workspace = tempWorkspace("accept-empty");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.accept",
			applied: [],
		});
	});

	test("-o/--output writes to a parallel file, original untouched", async () => {
		const docPath = await freshFixture("accept-output");
		const outPath = `${docPath}.out.docx`;
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--output",
			outPath,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(outPath);

		// Original still has the change.
		expect(
			(
				(await runCli("track-changes", "list", docPath)).parsed as Array<{
					id: string;
				}>
			).length,
		).toBe(1);
		// Output has it accepted.
		expect((await runCli("track-changes", "list", outPath)).parsed).toEqual([]);
	});
});
