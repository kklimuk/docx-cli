#!/usr/bin/env bun
/**
 * Stage the COMPETITOR for the A/B bake-off: Anthropic's official, bundled "docx"
 * Agent Skill (python + raw-OOXML), provisioned with its FULL intended toolset.
 *
 * Fairness is the whole ballgame for the bake-off — an under-equipped competitor
 * would void the comparison. So this script (a) fetches the real skill from
 * anthropics/skills, and (b) installs every dependency the skill's SKILL.md relies
 * on (python-docx, lxml, the Node `docx` library, pandoc), then VERIFIES each one is
 * actually usable. It exits non-zero if any REQUIRED piece is missing, so a human can
 * confirm a green competitor setup before spending a bake-off run.
 *
 * Usage:
 *   stage-competitor.ts <SKILL_DEST> [RUN_DIR]
 *     SKILL_DEST  where to place the skill (will contain SKILL.md + scripts/). Pass
 *                 this same path to the workflow as args.competitorDir.
 *     RUN_DIR     optional: the bake-off run dir. The Node `docx` library is installed
 *                 here so every scenario subfolder (<RUN_DIR>/<key>/) resolves it via
 *                 Node's upward module resolution.
 *
 * Idempotent: re-running skips work that's already done.
 */

import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO_URL = "https://github.com/anthropics/skills.git";
const SKILL_SUBPATH = "skills/docx";
// The skill's required artifacts — used for BOTH the "already staged?" skip guard and
// the post-copy verify, so a partial copy can't pass the guard yet fail the verify
// (and then never re-fetch to self-heal).
const REQUIRED = [
	"SKILL.md",
	"scripts/office/unpack.py",
	"scripts/office/pack.py",
	"scripts/comment.py",
];

let failed = false;
const note = (message: string) => console.log(`[ok]   ${message}`);
const warn = (message: string) => console.log(`[warn] ${message}`);
const fail = (message: string) => {
	console.log(`[FAIL] ${message}`);
	failed = true;
};

const [skillDest, runDir] = Bun.argv.slice(2);
if (!skillDest) {
	console.error("usage: stage-competitor.ts <SKILL_DEST> [RUN_DIR]");
	process.exit(2);
}

console.log("=== Staging competitor: Anthropic docx skill ===");
console.log(`    skill dest: ${skillDest}`);
if (runDir) console.log(`    run dir:    ${runDir}`);
console.log("");

await fetchSkill(skillDest);
for (const required of REQUIRED) {
	if (await Bun.file(join(skillDest, required)).exists()) {
		note(`present: ${required}`);
	} else {
		fail(`missing: ${join(skillDest, required)}`);
	}
}
await installPythonDeps();
await installNodeDocx(skillDest, runDir);
await ensurePandoc();
await ensureImageTooling();

console.log("");
if (failed) {
	console.log(
		"=== COMPETITOR STAGING INCOMPLETE — fix the [FAIL]s above before running the competitor arm. ===",
	);
	console.log("    (A handicapped competitor would void the bake-off.)");
	process.exit(1);
}
console.log("=== Competitor staged. Pass this to the workflow: ===");
console.log(`    arm: "anthropic-docx-skill", competitorDir: "${skillDest}"`);
if (runDir) {
	console.log(
		`    (Node \`docx\` installed under ${runDir} for scenario-folder resolution.)`,
	);
}

// ---------------------------------------------------------------------------
// 1. Fetch the real skill (SKILL.md + scripts/) into skillDest.
// ---------------------------------------------------------------------------
async function fetchSkill(dest: string): Promise<void> {
	if (await haveAllRequired(dest)) {
		note(`skill already present at ${dest} (skipping clone)`);
		return;
	}
	if (!Bun.which("git")) {
		fail("git not found — cannot fetch the skill");
		process.exit(1);
	}
	const clone = mkdtempSync(join(tmpdir(), "docx-skill-clone-"));
	// process.exit() SKIPS finally blocks, so every hard-exit path below must clean
	// the clone itself (the old .sh version did this with `trap … EXIT`).
	const bailOut = (message: string): never => {
		fail(message);
		rmSync(clone, { recursive: true, force: true });
		process.exit(1);
	};
	try {
		console.log(`    cloning ${REPO_URL} (shallow) ...`);
		// Try a sparse partial clone first (fast); fall back to a full clone if ANY
		// part of the sparse path fails — the clone, the sparse-checkout, OR an empty
		// resulting tree (old git, cone-mode quirks, or a server refusing the blob
		// filter).
		const sparse =
			(await Bun.$`git clone --depth 1 --filter=blob:none --sparse ${REPO_URL} ${clone}`
				.quiet()
				.nothrow()
				.then((result) => result.exitCode)) === 0 &&
			(await Bun.$`git -C ${clone} sparse-checkout set ${SKILL_SUBPATH}`
				.quiet()
				.nothrow()
				.then((result) => result.exitCode)) === 0 &&
			(await Bun.file(join(clone, SKILL_SUBPATH, "SKILL.md")).exists());
		if (!sparse) {
			rmSync(clone, { recursive: true, force: true });
			const full = await Bun.$`git clone --depth 1 ${REPO_URL} ${clone}`
				.quiet()
				.nothrow();
			if (full.exitCode !== 0) {
				bailOut(`git clone of ${REPO_URL} failed (network? auth?)`);
			}
		}
		if (!(await Bun.file(join(clone, SKILL_SUBPATH, "SKILL.md")).exists())) {
			bailOut(`${SKILL_SUBPATH} not found in the clone — repo layout may have changed`);
		}
		await Bun.$`mkdir -p ${dest}`.quiet();
		await Bun.$`cp -R ${join(clone, SKILL_SUBPATH)}/. ${dest}/`.quiet();
		note(`copied skill into ${dest}`);
	} finally {
		rmSync(clone, { recursive: true, force: true });
	}
}

async function haveAllRequired(dest: string): Promise<boolean> {
	for (const required of REQUIRED) {
		if (!(await Bun.file(join(dest, required)).exists())) return false;
	}
	return true;
}

// ---------------------------------------------------------------------------
// 2. Python deps (python-docx, lxml, defusedxml, pillow, numpy) for EVERY python
//    the agents might invoke. macOS commonly has MULTIPLE Homebrew pythons (e.g.
//    `python`=3.13 and `python3`=3.14) with SEPARATE site-packages, and the skill's
//    docs say `python scripts/...` — so the deps must be importable by BOTH `python`
//    and `python3`, or an agent hits ModuleNotFoundError and a PEP-668-blocked
//    `pip install` it can't recover from. Install with --break-system-packages (the
//    reliable option on externally-managed Homebrew python) and VERIFY each.
// ---------------------------------------------------------------------------
async function installPythonDeps(): Promise<void> {
	const IMPORT_CHECK = "import docx, lxml, defusedxml.minidom, PIL, numpy";
	const PACKAGES = ["python-docx", "lxml", "defusedxml", "pillow", "numpy"];
	let found = 0;
	for (const python of ["python", "python3"]) {
		if (!Bun.which(python)) continue;
		found += 1;
		const importable = () =>
			Bun.$`${python} -c ${IMPORT_CHECK}`
				.quiet()
				.nothrow()
				.then((result) => result.exitCode === 0);
		if (await importable()) {
			note(`${python}: python deps already importable`);
			continue;
		}
		console.log(`    installing ${PACKAGES.join(" + ")} for ${python} ...`);
		const attempts = [
			["--quiet", "--break-system-packages"],
			["--quiet", "--user", "--break-system-packages"],
			["--quiet", "--user"],
		];
		for (const flags of attempts) {
			const install =
				await Bun.$`${python} -m pip install ${flags} ${PACKAGES}`
					.quiet()
					.nothrow();
			if (install.exitCode === 0) break;
		}
		if (await importable()) {
			note(`${python}: python deps installed and importable`);
		} else {
			fail(
				`${python}: python-docx/lxml/defusedxml NOT importable — install them manually, then re-run`,
			);
		}
	}
	if (found === 0) {
		fail("no python/python3 on PATH — the skill's scripts cannot run");
	}
}

// ---------------------------------------------------------------------------
// 3. Node `docx` library (the skill creates documents with it). Install where the
//    weak agents can resolve it: RUN_DIR (so every scenario subfolder finds it) and,
//    as a fallback, SKILL_DEST.
// ---------------------------------------------------------------------------
async function installNodeDocx(
	dest: string,
	runDirArg: string | undefined,
): Promise<void> {
	if (!Bun.which("npm")) {
		warn("npm not found — the skill's Node `docx` create flow will be unavailable");
		return;
	}
	for (const target of [runDirArg, dest]) {
		if (!target) continue;
		if (!(await Bun.file(join(target, "node_modules/docx/package.json")).exists())) {
			await Bun.$`mkdir -p ${target}`.quiet();
			if (!(await Bun.file(join(target, "package.json")).exists())) {
				await Bun.$`npm init -y`.cwd(target).quiet().nothrow();
			}
			await Bun.$`npm install --silent --no-audit --no-fund docx`
				.cwd(target)
				.quiet()
				.nothrow();
		}
		if (await Bun.file(join(target, "node_modules/docx/package.json")).exists()) {
			note(`Node \`docx\` available at ${target}/node_modules`);
		} else {
			warn(
				`could not install Node \`docx\` at ${target} (create-from-scratch flow may be limited)`,
			);
		}
	}
}

// ---------------------------------------------------------------------------
// 4. pandoc (the skill's read/text-extraction path). Strongly recommended.
// ---------------------------------------------------------------------------
async function ensurePandoc(): Promise<void> {
	if (Bun.which("pandoc")) {
		const version = await Bun.$`pandoc --version`.quiet().nothrow().text();
		note(`pandoc present (${version.split("\n")[0]})`);
		return;
	}
	if (!Bun.which("brew")) {
		warn(
			"pandoc NOT found and no brew to install it — the skill's read path will be degraded",
		);
		return;
	}
	console.log("    installing pandoc via brew ...");
	await Bun.$`brew install pandoc`.quiet().nothrow();
	if (Bun.which("pandoc")) {
		note("pandoc installed via brew");
	} else {
		warn(
			"pandoc install failed — the skill's read path will be degraded (install manually for full fairness)",
		);
	}
}

// ---------------------------------------------------------------------------
// 5. Image tooling for SVG figure scenarios. The skill documents LibreOffice for
//    PDF, but weak agents reach for ImageMagick `convert` + librsvg to rasterize
//    SVG figures (e.g. a frontispiece or an SVG logo). Provide them so an image
//    task isn't blocked by a missing system tool (the docx-cli arm handles SVG
//    natively, so withholding this would handicap only the competitor).
// ---------------------------------------------------------------------------
async function ensureImageTooling(): Promise<void> {
	if (Bun.which("convert") && Bun.which("rsvg-convert")) {
		note("image tooling present (convert + rsvg-convert)");
		return;
	}
	if (!Bun.which("brew")) {
		warn(
			"imagemagick/librsvg NOT found and no brew — SVG figure scenarios may be limited for the competitor",
		);
		return;
	}
	console.log("    installing imagemagick + librsvg via brew (may take a few min) ...");
	await Bun.$`brew install imagemagick librsvg`.quiet().nothrow();
	if (Bun.which("convert")) {
		note("imagemagick + librsvg installed");
	} else {
		warn("image tooling install failed — SVG figure scenarios may be limited for the competitor");
	}
}
