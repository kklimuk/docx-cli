#!/usr/bin/env bun
import { resolve } from "node:path";
import { writeStdout } from "../../src/cli/respond";

const [directory] = Bun.argv.slice(2);
if (!directory) throw new Error("Usage: bun benchmarks/officecli/verify.ts RUN_ROOT");
const root = resolve(directory);
const manifest = await Bun.file(resolve(root, "manifest.json")).json();
const checks: Array<[string, string]> = [
	[manifest.docx.binary, manifest.docx.sha256],
	[manifest.officecli.binary, manifest.officecli.sha256],
	[resolve(manifest.officecli.source, "SKILL.md"), manifest.officecli.skillSha256],
	[manifest.harness.scriptPath, manifest.harness.workflowSha256],
	[resolve(root, "harness/weak-exercise.md"), manifest.harness.agentSha256],
];
for (const [path, hash] of Object.entries(manifest.fixtures)) {
	checks.push([resolve(root, "harness/skill/scenarios", path), String(hash)]);
}
for (const [path, hash] of Object.entries(manifest.harness.files ?? {})) {
	checks.push([resolve(root, "harness", path), String(hash)]);
}
for (const [path, expected] of checks) {
	const actual = new Bun.CryptoHasher("sha256").update(await Bun.file(path).arrayBuffer()).digest("hex");
	if (actual !== expected) throw new Error(`Benchmark input changed: ${path}`);
}
await writeStdout(`Verified ${checks.length} frozen benchmark inputs. This verifies inputs, not model results or run isolation.\n`);
