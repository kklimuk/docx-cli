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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});
