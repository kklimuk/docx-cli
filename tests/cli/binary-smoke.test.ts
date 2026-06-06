import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { spawnCli, tempWorkspace } from "./harness";

// The bulk of the CLI suite runs in-process (see harness.ts runCli) for speed.
// These few tests spawn the REAL binary (`bun src/index.ts`) so the process
// boundary itself stays covered: argv handling, exit-code propagation via
// src/index.ts, and the 64 KB Bun.stdout truncation that respond.ts guards.
// quiet-default.test.ts also spawns (quiet/verbose acks); keep subprocess tests
// to this handful.

const FIXTURES = join(import.meta.dir, "..", "fixtures");

describe("binary smoke (real subprocess)", () => {
	test("--version prints the version and exits 0", async () => {
		const result = await spawnCli("--version");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^docx \d+\.\d+\.\d+/);
	});

	test("--help prints usage and exits 0", async () => {
		const result = await spawnCli("--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("Usage");
	});

	test("unknown command exits non-zero with an error payload", async () => {
		const result = await spawnCli("frobnicate");
		expect(result.exitCode).not.toBe(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("not-found locator propagates exit code 3", async () => {
		const workspace = tempWorkspace("smoke-exit");
		const docPath = join(workspace, "out.docx");
		await spawnCli("create", docPath, "--text", "hi");
		const result = await spawnCli("edit", docPath, "--at", "p9", "--text", "x");
		expect(result.exitCode).toBe(3);
	});

	test("large stdout (>64 KB) is not truncated", async () => {
		// academic-paper.docx --ast is ~77 KB — past the 64 KB write boundary that
		// respond.ts/writeStdout exist to handle. The output must arrive whole.
		const result = await spawnCli(
			"read",
			join(FIXTURES, "academic-paper.docx"),
			"--ast",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout.length).toBeGreaterThan(64 * 1024);
		// Complete + valid JSON end-to-end (truncation would break the parse).
		expect(() => JSON.parse(result.stdout)).not.toThrow();
	});
});
