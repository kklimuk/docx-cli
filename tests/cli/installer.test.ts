import { describe, expect, test } from "bun:test";

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
});
