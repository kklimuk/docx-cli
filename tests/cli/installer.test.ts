import { describe, expect, test } from "bun:test";
import {
	buildThirdPartyLicenses,
	collectProductionPackages,
	THIRD_PARTY_LICENSES_FILE,
} from "../../scripts/third-party-licenses";
import { tempWorkspace } from "./harness";

const SKILL_INSTALLER = "skills/docx-cli/scripts/install.sh";
const ROOT_INSTALLER = "install.sh";
const BOOTSTRAP = "skills/docx-cli/scripts/bootstrap.sh";
const RELEASE_WORKFLOW = ".github/workflows/release.yml";
const NOTICES = "NOTICES";

describe("release attribution drift guards", () => {
	// Every package in the production closure is redistributed inside dist/index.js
	// and the binaries, and its license requires its notice to travel with it. The
	// generator throws on a missing license text or a non-permissive license, so this
	// also stops a copyleft dependency from landing unnoticed. Offline: it reads the
	// installed node_modules, never the network.
	test("THIRD-PARTY-LICENSES matches the installed production dependencies", async () => {
		const committed = (
			await Bun.file(THIRD_PARTY_LICENSES_FILE).text()
		).replace(/\r\n/g, "\n");
		// Stale? Run `bun run licenses` and commit the result.
		expect(committed).toBe(buildThirdPartyLicenses("."));
	});

	// MIT/ISC terms are satisfied by the LICENSE text alone. Anything else (BSD
	// notices, Zlib, a dual license we elect from) gets a named entry in NOTICES too.
	test("NOTICES names every production package that isn't plain MIT or ISC", async () => {
		const notices = await Bun.file(NOTICES).text();
		// An entry OPENS a line: `----  name (…)` for a reproduced notice, or the
		// two-space-indented elections table. A bare substring would let `diff`
		// pass on the word "different".
		const hasEntry = (name: string) =>
			new RegExp(
				`^(?:----\\s+|  )${name.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&")}\\s`,
				"m",
			).test(notices);
		const unnamed = [...collectProductionPackages(".").values()]
			.filter((licensed) => !["MIT", "ISC"].includes(licensed.license))
			.filter((licensed) => !hasEntry(licensed.name))
			.map((licensed) => `${licensed.name} (${licensed.license})`);
		expect(unnamed).toEqual([]);
	});

	// The attribution has to reach the people who get the bytes: the npm tarball
	// (package.json `files`) and every release binary (the release staging step).
	test("the npm package and the release artifacts both ship the attribution files", async () => {
		const [manifest, workflow] = await Promise.all([
			Bun.file("package.json").json() as Promise<{ files: string[] }>,
			Bun.file(RELEASE_WORKFLOW).text(),
		]);
		const staging = workflow.match(
			/run: cp LICENSE [^\n]*\.\/artifacts\//,
		)?.[0];
		expect(staging).toBeDefined();
		for (const file of ["LICENSE", NOTICES, THIRD_PARTY_LICENSES_FILE]) {
			expect(manifest.files).toContain(file);
			expect(staging).toContain(file);
		}
		// The binaries embed the Bun runtime (LGPL JavaScriptCore + tinycc), so the
		// LGPL texts NOTICES points at travel with them; the npm package embeds no
		// Bun, so they aren't required there.
		const notices = await Bun.file(NOTICES).text();
		for (const file of ["LGPL-2.0.txt", "LGPL-2.1.txt"]) {
			expect(await Bun.file(file).exists()).toBe(true);
			expect(notices).toContain(file);
			expect(staging).toContain(file);
		}
	});
});

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
			// Asks sh, as install.sh does: on Windows Bun.which also finds a shasum.bat
			// that sh can't run. An empty stdout therefore means the tool is absent.
			const script = `command -v ${tool} >/dev/null 2>&1 || exit 0\n${definition}\nsha256_of "$1"`;
			const proc = Bun.spawn(["sh", "-c", script, "sh", path], {
				stdout: "pipe",
			});
			const stdout = (await new Response(proc.stdout).text()).trim();
			expect(await proc.exited).toBe(0);
			if (stdout) expect(stdout).toBe(expected);
		}
	});
});
