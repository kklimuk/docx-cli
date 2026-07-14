#!/usr/bin/env bun
/**
 * Stage ONE scenario folder into a run workspace: copy the pristine scenario dir,
 * strip the judge-only criteria.md, and verify the required inputs landed. This is
 * the ONE staging path — the workflow's Stage agent runs it per scenario, and
 * run-local-corpus.ts imports stageScenario() for the local-harness arm, so every
 * arm stages identically.
 *
 * Usage:
 *   stage-scenario.ts <srcDir> <dstDir> [--require-doc FILE]
 *     --require-doc FILE   also verify the fixture .docx landed (edit scenarios)
 *
 * Prints a one-line JSON verdict {key, staged, missing} and exits nonzero when
 * anything required is absent or the answer key leaked into the destination.
 */

import { basename } from "node:path";

if (import.meta.main) {
	const positional: string[] = [];
	let requireDoc: string | null = null;
	const argv = Bun.argv.slice(2);
	for (let index = 0; index < argv.length; index += 1) {
		const arg = argv[index] ?? "";
		if (arg === "--require-doc") {
			requireDoc = argv[++index] ?? null;
			continue;
		}
		positional.push(arg);
	}
	const [srcDir, dstDir] = positional;
	if (!srcDir || !dstDir) {
		await Bun.write(
			Bun.stderr,
			"Usage: stage-scenario.ts <srcDir> <dstDir> [--require-doc FILE]\n",
		);
		process.exit(2);
	}
	const verdict = await stageScenario(srcDir, dstDir, requireDoc);
	await Bun.write(Bun.stdout, `${JSON.stringify(verdict)}\n`);
	process.exit(verdict.staged ? 0 : 1);
}

export async function stageScenario(
	srcDir: string,
	dstDir: string,
	requireDoc?: string | null,
): Promise<StageVerdict> {
	const key = basename(srcDir.replace(/\/+$/, ""));
	const missing: string[] = [];
	if (!(await Bun.file(`${srcDir}/task.md`).exists())) {
		missing.push(
			`${key}: ${srcDir}/task.md not found — is this a scenario folder?`,
		);
		return { key, staged: false, missing };
	}

	// Copy the folder CONTENTS (task.md, the fixture, assets/) and then delete the
	// answer key — criteria.md is the judge's rubric and must NEVER reach the
	// agent's workspace (the judge reads it from the pristine source instead).
	// .nothrow() so a shell failure (permissions, disk full) degrades to a `missing`
	// verdict + exit 1 instead of a thrown ShellError — the corpus runner loops over
	// scenarios and must be able to skip-and-continue past a bad one.
	const copy =
		await Bun.$`mkdir -p ${dstDir} && cp -R ${srcDir}/. ${dstDir}/ && rm -f ${dstDir}/criteria.md`
			.quiet()
			.nothrow();
	if (copy.exitCode !== 0) {
		missing.push(
			`${key}: staging shell failed (exit ${copy.exitCode}): ${copy.stderr.toString().trim() || "(no stderr)"}`,
		);
	}

	if (!(await Bun.file(`${dstDir}/task.md`).exists())) {
		missing.push(`${key}: ${dstDir}/task.md missing after copy`);
	}
	if (requireDoc && !(await Bun.file(`${dstDir}/${requireDoc}`).exists())) {
		missing.push(`${key}: fixture ${dstDir}/${requireDoc} missing after copy`);
	}
	if (await Bun.file(`${dstDir}/criteria.md`).exists()) {
		missing.push(
			`${key}: criteria.md still present in ${dstDir} — answer-key leak`,
		);
	}
	return { key, staged: missing.length === 0, missing };
}

export type StageVerdict = { key: string; staged: boolean; missing: string[] };
