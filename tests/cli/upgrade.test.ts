import { describe, expect, test } from "bun:test";
import installScript from "../../skills/docx-cli/scripts/install.sh" with {
	type: "text",
};
import { runCli } from "./harness";

describe("upgrade", () => {
	// The whole point of the design: `upgrade` reuses the installer rather than
	// reimplementing the platform table, download, and checksum logic in TypeScript.
	// If this drifts there are two implementations again — the exact problem the
	// install.sh/bootstrap.sh consolidation removed. A stubbed or empty embed (which
	// would still typecheck, since `*.sh` is declared as `string`) fails here too.
	test("embeds the skill's install.sh verbatim, not a copy of its logic", async () => {
		const onDisk = await Bun.file("skills/docx-cli/scripts/install.sh").text();
		expect(installScript).toBe(onDisk);
	});

	// `--to` is interpolated into install.sh's release URL and curl normalizes "..",
	// so an unconstrained value redirects the binary AND the SHA256SUMS it is verified
	// against to another repo — the checksum gate then blesses the attacker's manifest.
	// This CLI is driven by agents over untrusted document content, so these must not
	// reach the installer.
	test.each([
		"../../../../octocat/Hello-World/releases/latest",
		"..",
		"v1.0.0/../../../other/repo/releases/download/v1",
		"latest",
		"main",
		"https://evil.example/x",
		"",
	])("rejects --to %p rather than passing it into the release URL", async (tag) => {
		const result = await runCli("upgrade", "--to", tag);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("release tag");
	});

	test.each([
		"v0.23.0",
		"0.23.0",
		"1.2.3-rc.1",
	])("accepts the release tag %p", async (tag) => {
		// Reaches the install-method check (npm here), i.e. past validation.
		const result = await runCli("upgrade", "--to", tag);
		expect(result.exitCode).not.toBe(2);
	});

	// The suite runs from source under Bun, never as a compiled executable, so
	// isStandaloneBinary() is false here — which is exactly the npm/bun install case.
	// Both argv forms must refuse: the install-method check precedes --dry-run.
	test.each([
		[[]],
		[["--dry-run"]],
	])("a package-manager install is refused and directed to the package manager (%p)", async (argv) => {
		const result = await runCli("upgrade", ...(argv as string[]));
		// Nonzero: nothing was upgraded, and weak agents key off the exit code.
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("bun add -g bun-docx");
		expect(result.stdout).not.toContain("Downloading");
	});

	test("--help documents the install-method split and the output shapes", async () => {
		const result = await runCli("upgrade", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("--to");
		expect(result.stdout).toContain("--dry-run");
		expect(result.stdout).toContain("Output:");
		expect(result.stdout).toMatch(/npm|package manager/i);
	});

	test("rejects unknown flags rather than silently ignoring them", async () => {
		const result = await runCli("upgrade", "--force");
		expect(result.exitCode).toBe(2);
	});

	// `docx upgrade v0.23.0` is the shape an agent reaches for. Left as a positional it
	// would upgrade to LATEST and exit 0 — a confidently-wrong result at the one moment
	// a weak agent stops reading.
	test("rejects a bare version positional instead of silently installing latest", async () => {
		const result = await runCli("upgrade", "v0.23.0");
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("--to");
	});
});

// isStandaloneBinary() keys on `/$bunfs/`, an UNDOCUMENTED Bun internal path that a
// Bun upgrade could rename silently — and the failure is invisible from the source
// tree, since every standalone `docx upgrade` would just degrade to the npm advisory
// and no in-process test would go red. This is the only guard for that, so it builds
// the real artifact and asks it what it thinks it is.
describe("upgrade (compiled binary)", () => {
	test("a compiled binary detects itself as standalone", async () => {
		const dir = `${import.meta.dir}/../../node_modules/.cache/upgrade-smoke`;
		const binary = `${dir}/docx`;
		const build = Bun.spawnSync([
			"bun",
			"build",
			"--compile",
			"--outfile",
			binary,
			"src/index.ts",
		]);
		expect(build.exitCode).toBe(0);

		// --dry-run so it reports intent without touching the network or the binary.
		const result = Bun.spawnSync([binary, "upgrade", "--dry-run"]);
		expect(result.exitCode).toBe(0);
		const stdout = result.stdout.toString();
		expect(stdout).toContain('"dryRun":true');
		// It reports the binary it would replace — process.execPath, normalized.
		expect(stdout).toContain("upgrade-smoke/docx");
		// The npm advisory would mean detection broke.
		expect(stdout).not.toContain("bun add -g bun-docx");

		// install.sh always writes `${PREFIX}/docx`, so a RENAMED binary can't be
		// self-replaced: the install would drop a second file beside it and leave the
		// command the user runs stale, while upgrade reported success on the old path.
		const renamed = `${dir}/mydocx`;
		expect(Bun.spawnSync(["cp", binary, renamed]).exitCode).toBe(0);
		const refusal = Bun.spawnSync([renamed, "upgrade", "--dry-run"]);
		expect(refusal.exitCode).not.toBe(0);
		expect(refusal.stdout.toString()).toContain("mydocx");
	}, 120_000);
});
