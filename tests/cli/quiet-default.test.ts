import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { spawnCli as rawCli, tempWorkspace } from "./harness";

// Mutators are silent on success unless --verbose is passed. The shared harness
// auto-injects --verbose for mutators (so the rest of the suite keeps asserting
// on JSON acks); these tests bypass it by spawning the REAL binary via
// spawnCli. They double as the binary smoke layer (real process boundary +
// exit codes) now that the bulk of the suite runs in-process — see
// binary-smoke.test.ts for the remaining process-level cases.

describe("docx mutators — quiet by default", () => {
	test("create succeeds with empty stdout when --verbose is omitted", async () => {
		const workspace = tempWorkspace("quiet-create");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli("create", docPath, "--text", "hello");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("");
	});

	test("create with --verbose prints the JSON ack", async () => {
		const workspace = tempWorkspace("verbose-create");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli(
			"create",
			docPath,
			"--text",
			"hello",
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "create" });
	});

	test("-v shorthand also enables the ack", async () => {
		const workspace = tempWorkspace("verbose-short");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli("create", docPath, "--text", "hello", "-v");
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true });
	});

	test("errors print regardless of --verbose", async () => {
		const result = await rawCli("create", "--text", "hello"); // missing FILE
		expect(result.exitCode).not.toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("--dry-run prints regardless of --verbose", async () => {
		const workspace = tempWorkspace("dry-run-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		const result = await rawCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"new para",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true, dryRun: true });
	});

	test("read commands stay loud (no --verbose needed)", async () => {
		const workspace = tempWorkspace("read-loud");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		const result = await rawCli("find", docPath, "hello");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.length).toBeGreaterThan(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "find" });
	});

	test("track-changes on is silent by default; --verbose prints the toggle ack", async () => {
		const workspace = tempWorkspace("tc-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		const quiet = await rawCli("track-changes", docPath, "on");
		expect(quiet.exitCode).toBe(0);
		expect(quiet.stdout).toBe("");

		await rawCli("track-changes", docPath, "off");
		const verbose = await rawCli("track-changes", docPath, "on", "--verbose");
		expect(verbose.exitCode).toBe(0);
		const payload = JSON.parse(verbose.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "track-changes" });
	});

	test("comments add is silent by default; --verbose returns commentId", async () => {
		const workspace = tempWorkspace("comments-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello world");

		const quiet = await rawCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"check",
		);
		expect(quiet.exitCode).toBe(0);
		expect(quiet.stdout).toBe("");

		const verbose = await rawCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"check 2",
			"--verbose",
		);
		expect(verbose.exitCode).toBe(0);
		const payload = JSON.parse(verbose.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "comments.add" });
		expect(payload).toHaveProperty("commentId");
	});
});
