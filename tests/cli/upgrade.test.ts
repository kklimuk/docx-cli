import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import installScript from "../../skills/docx-cli/scripts/install.sh" with {
	type: "text",
};
import { withoutStagingPathHint } from "../../src/cli/upgrade/staging-path-hint";
import { swapIn } from "../../src/cli/upgrade/swap-in";
import { runCli, tempWorkspace } from "./harness";

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

// Windows locks a running .exe against overwrite but not rename, so upgrade stages
// the verified binary and swaps it in. Plain file moves, so this runs anywhere.
describe("upgrade (Windows swap-in)", () => {
	// The previous upgrade's `.old` could not delete itself while it ran; by now it
	// isn't running, so the next upgrade clears it instead of failing on it.
	test("puts the staged binary in place and keeps the old one beside it", async () => {
		const workspace = tempWorkspace("swap-in");
		const binary = `${workspace}/docx.exe`;
		const staged = `${workspace}/.docx-upgrade/docx.exe`;
		await Bun.write(binary, "current");
		await Bun.write(`${binary}.old`, "previous");
		await Bun.write(staged, "new");

		await swapIn(staged, binary);
		expect(await Bun.file(binary).text()).toBe("new");
		expect(await Bun.file(`${binary}.old`).text()).toBe("current");
		expect(await Bun.file(staged).exists()).toBe(false);
	});

	// Antivirus can hold a freshly written .exe open for a moment, so the second
	// rename can fail after the first moved the running binary aside. The user
	// must not be left with only docx.exe.old.
	test("puts the running binary back when the staged one can't be moved in", async () => {
		const workspace = tempWorkspace("swap-in-fail");
		const binary = `${workspace}/docx.exe`;
		await Bun.write(binary, "current");

		await expect(
			swapIn(`${workspace}/.docx-upgrade/docx.exe`, binary),
		).rejects.toThrow();
		expect(await Bun.file(binary).text()).toBe("current");
	});

	// install.sh's PATH note names its PREFIX, which on Windows is the staging
	// folder upgrade deletes right after — so the note is dropped, nothing else.
	test("drops install.sh's PATH note about the staging folder, keeps the rest", async () => {
		const staging = "C:\\bin\\.docx-upgrade";
		const lines = [
			"→ Downloading docx-windows-x64.exe from https://example/v1",
			"✓ Verified SHA-256 (docx-windows-x64.exe)",
			`✓ Installed: ${staging}/docx.exe`,
			"",
			`  Note: ${staging} is not on your PATH.`,
			"  Add this to your shell profile (~/.zshrc, ~/.bashrc, etc):",
			`    export PATH="${staging}:$PATH"`,
			"",
			"docx 0.29.0",
		];
		let out = "";
		const filter = withoutStagingPathHint(async (text) => {
			out += text;
		}, staging);
		// Chunk boundaries fall mid-line, as they do on a real pipe.
		const text = `${lines.join("\n")}\n`;
		for (let start = 0; start < text.length; start += 7) {
			await filter.write(text.slice(start, start + 7));
		}
		await filter.flush();
		expect(out).not.toContain("not on your PATH");
		expect(out).not.toContain("export PATH");
		expect(out).toContain("✓ Verified SHA-256");
		expect(out).toContain("docx 0.29.0\n");
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
		// `bun build --compile` appends it on Windows, and install.sh names it the same.
		const exe = process.platform === "win32" ? ".exe" : "";
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
		expect(JSON.parse(stdout).path).toEndWith(
			join("upgrade-smoke", `docx${exe}`),
		);
		// The npm advisory would mean detection broke.
		expect(stdout).not.toContain("bun add -g bun-docx");

		// install.sh always writes `${PREFIX}/docx`, so a RENAMED binary can't be
		// self-replaced: the install would drop a second file beside it and leave the
		// command the user runs stale, while upgrade reported success on the old path.
		const renamed = `${dir}/mydocx${exe}`;
		expect(Bun.spawnSync(["cp", `${binary}${exe}`, renamed]).exitCode).toBe(0);
		const refusal = Bun.spawnSync([renamed, "upgrade", "--dry-run"]);
		expect(refusal.exitCode).not.toBe(0);
		expect(refusal.stdout.toString()).toContain("mydocx");
	}, 120_000);
});
