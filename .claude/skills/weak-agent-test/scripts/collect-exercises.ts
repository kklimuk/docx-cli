#!/usr/bin/env bun
/**
 * Deterministically collect the local corpus runner's per-scenario exercise.json
 * results into the `args.exercises` array the weak-agent-test workflow accepts INLINE.
 *
 * Why this exists: the workflow's JS is sandboxed with no filesystem access, so the
 * only way it can obtain the local results is (a) an in-workflow LOAD agent — an LLM
 * that reads the files and hands them back, or (b) `args.exercises` passed in from the
 * caller. This script is path (b): it reads the files with plain code, so the
 * code-computed `status` (completed | failed) and the rest of the account reach the
 * workflow verbatim instead of being laundered through a model. Pass its output as
 * `args.exercises` and the workflow skips the LOAD agent entirely.
 *
 * Only the fields the workflow consumes are emitted (status/completed, summary,
 * deadEnds, frictions, outputPath) — small and low-risk to pass through. The big
 * `docxCommands` / `_local` ledger blocks stay on disk in exercise.json, which is the
 * deterministic source of truth for anything that needs them.
 *
 * Usage: collect-exercises.ts <runDir> [key...]
 *   Prints a JSON array to stdout — pass it verbatim as the workflow's args.exercises.
 */

const [runDir, ...keys] = Bun.argv.slice(2);
if (!runDir) {
	await Bun.write(Bun.stderr, "Usage: collect-exercises.ts <runDir> [key...]\n");
	process.exit(2);
}

const wanted = new Set(keys);
const exercises: Record<string, unknown>[] = [];
for await (const rel of new Bun.Glob("*/exercise.json").scan({ cwd: runDir })) {
	const key = rel.slice(0, rel.indexOf("/"));
	if (wanted.size && !wanted.has(key)) continue;
	try {
		const entry = await Bun.file(`${runDir}/${rel}`).json();
		exercises.push({
			key: entry.key ?? key,
			// `status` is the code-computed process lifecycle (new files); `completed`
			// is only present on files produced before the switch. JSON.stringify drops
			// whichever is undefined.
			status: entry.status,
			completed: entry.completed,
			summary: entry.summary,
			deadEnds: entry.deadEnds,
			frictions: entry.frictions,
			outputPath: entry.outputPath,
		});
	} catch (error) {
		await Bun.write(
			Bun.stderr,
			`[${key}] skipped (unreadable exercise.json): ${error}\n`,
		);
	}
}
// Stable order for reproducibility.
exercises.sort((a, b) =>
	String(a.key) < String(b.key) ? -1 : String(a.key) > String(b.key) ? 1 : 0,
);
await Bun.write(Bun.stdout, JSON.stringify(exercises));
