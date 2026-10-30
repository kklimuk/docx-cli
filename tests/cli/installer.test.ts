import { describe, expect, test } from "bun:test";
import { tempWorkspace } from "./harness";

const SKILL_INSTALLER = "skills/docx-cli/scripts/install.sh";
const ROOT_INSTALLER = "install.sh";
const BOOTSTRAP = "skills/docx-cli/scripts/bootstrap.sh";
const RELEASE_WORKFLOW = ".github/workflows/release.yml";

describe("installer drift guards", () => {
	// The skill folder is the single source: bootstrap.sh delegates to its sibling, and
	// the release workflow publishes that same file as the `install.sh` release asset.
	//
	// The repo-root copy is a DEPRECATION SHIM, not a second implementation. Bootstraps
	// already published to npm/skills.sh (<= v0.22.0) fetch
	// raw.githubusercontent.com/kklimuk/docx-cli/<tag>/install.sh, so a tag whose tree
	// lacks that path breaks every deployed skill copy's session-start check. Keep the
	// shim until those bootstraps have aged out, then delete it and this test.
	test("repo-root install.sh is byte-identical to the skill copy", async () => {
		const [root, skill] = await Promise.all([
			Bun.file(ROOT_INSTALLER).text(),
			Bun.file(SKILL_INSTALLER).text(),
		]);
		expect(root).toBe(skill);
	});

	// detect_target's asset names are what install.sh requests from the release; the
	// workflow matrix's `name:` values are what actually gets uploaded. If they disagree
	// the download 404s — and the release workflow only runs on a tag, so nothing else
	// catches it before the release is public.
	test("install.sh asset names match the release workflow build matrix", async () => {
		const [installer, workflow] = await Promise.all([
			Bun.file(SKILL_INSTALLER).text(),
			Bun.file(RELEASE_WORKFLOW).text(),
		]);

		const detectTarget = installer.match(
			/detect_target\(\) \{[\s\S]*?\n\}/,
		)?.[0];
		expect(detectTarget).toBeDefined();

		const requested = new Set(
			[...(detectTarget as string).matchAll(/"(docx-[\w.-]+)"/g)].map(
				(match) => match[1],
			),
		);
		const published = new Set(
			[...workflow.matchAll(/name:\s*(docx-[\w.-]+)\s*\}/g)].map(
				(match) => match[1],
			),
		);

		expect(requested.size).toBeGreaterThan(0);
		expect([...requested].sort()).toEqual([...published].sort());
	});

	test("both scripts stay free of the fetch-then-execute pattern", async () => {
		// The Snyk CRITICAL that motivated this layout was bootstrap.sh downloading
		// install.sh from raw.githubusercontent.com and running it. Executing a shipped
		// local sibling is fine; fetching code and running it is not.
		for (const path of [SKILL_INSTALLER, ROOT_INSTALLER, BOOTSTRAP]) {
			const source = await Bun.file(path).text();
			expect(source).not.toMatch(/raw\.githubusercontent\.com/);
			expect(source).not.toMatch(/\|\s*(sh|bash)\b/);
			expect(source).not.toMatch(/(sh|bash)\s+-c\s+"\$\(/);
		}
	});

	// `docx upgrade` embeds the checked-out bytes and pipes them to `sh -s`. Only a
	// Windows checkout can produce CRLF here, so the Windows CI job is what guards this
	// (v0.28.0's Windows binary shipped a CRLF copy).
	test("shell scripts check out with LF line endings", async () => {
		for (const path of [SKILL_INSTALLER, ROOT_INSTALLER, BOOTSTRAP]) {
			expect(await Bun.file(path).text()).not.toContain("\r");
		}
	});
});

describe("installer sha256_of", () => {
	// GNU sha256sum escapes a file name holding a backslash by prefixing its output line
	// with `\`, which used to land in the digest and fail every Windows checksum. Each
	// fallback is run on its own, since a machine normally only reaches the first.
	test("every checksum tool prints the bare digest for a backslash path", async () => {
		const installer = await Bun.file(SKILL_INSTALLER).text();
		const definitions = [
			...installer.matchAll(/sha256_of\(\) \{ (\S+) .*\}/g),
		].map((match) => ({ definition: match[0], tool: match[1] as string }));
		expect(definitions.map(({ tool }) => tool)).toEqual([
			"sha256sum",
			"shasum",
			"openssl",
		]);

		const path = `${tempWorkspace("sha256-of")}/back\\slash`;
		await Bun.write(path, "abc");
		const expected = new Bun.CryptoHasher("sha256").update("abc").digest("hex");

		for (const { definition, tool } of definitions) {
			if (!Bun.which(tool)) continue;
			const proc = Bun.spawn(
				["sh", "-c", `${definition}\nsha256_of "$1"`, "sh", path],
				{ stdout: "pipe" },
			);
			const stdout = await new Response(proc.stdout).text();
			expect(await proc.exited).toBe(0);
			expect(stdout.trim()).toBe(expected);
		}
	});
});
