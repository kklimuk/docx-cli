#!/usr/bin/env bun
import { chmod, cp, mkdir } from "node:fs/promises";
import { resolve } from "node:path";
import { writeStdout } from "../../src/cli/respond";

const [officeDirectory, destination] = Bun.argv.slice(2);
if (!officeDirectory || !destination) {
	throw new Error("Usage: bun benchmarks/officecli/prepare.ts OFFICECLI_STAGE EMPTY_RUN_ROOT");
}
const repo = resolve(import.meta.dir, "../..");
const root = resolve(destination);
const office = await Bun.file(resolve(officeDirectory, "manifest.json")).json();
if (await sha256(office.binary) !== office.sha256) throw new Error("OfficeCLI binary changed since staging");
if (await sha256(resolve(office.source, "SKILL.md")) !== office.skillSha256) {
	throw new Error("OfficeCLI skill changed since staging");
}
await mkdir(root, { recursive: false });
await command(["bun", "run", "build:binary"]);
const binary = resolve(root, process.platform === "win32" ? "docx.exe" : "docx");
await Bun.write(binary, Bun.file(resolve(repo, "dist/docx")));
await chmod(binary, 0o755);
const version = (await command([binary, "--version"])).trim();
const expected = (await Bun.file(resolve(repo, "package.json")).json()).version;
if (version !== `docx ${expected}`) throw new Error(`Version mismatch: ${version}`);
await command([binary, "render", "--help"]);
const diff = await command(["git", "diff", "HEAD", "--binary"]);
await Bun.write(resolve(root, "worktree.patch"), diff);
const status = await command(["git", "status", "--short"]);
// Snapshot everything the runtime reads, so another working session cannot
// change tasks or grading between arms or repetitions.
const harnessDir = resolve(root, "harness");
await mkdir(harnessDir);
await cp(resolve(repo, ".claude/skills/weak-agent-test"), resolve(harnessDir, "skill"), { recursive: true });
const scriptPath = resolve(harnessDir, "weak-agent-test.js");
await Bun.write(scriptPath, Bun.file(resolve(repo, ".claude/workflows/weak-agent-test.js")));
const agentDefinition = await Bun.file(resolve(repo, ".claude/agents/weak-exercise.md")).text();
await Bun.write(resolve(harnessDir, "weak-exercise.md"), agentDefinition);
await Bun.write(resolve(harnessDir, "agents.json"), JSON.stringify({
	"weak-exercise": {
		description: "Isolated document-task benchmark exercise agent",
		prompt: agentDefinition.replace(/^---[\s\S]*?---\s*/, ""),
		tools: ["Bash", "Read", "Write", "Glob", "Grep"],
		disallowedTools: ["Skill"],
	},
}, null, 2));
const harnessFiles: Record<string, string> = {};
for await (const entry of new Bun.Glob("**/*").scan({ cwd: harnessDir, onlyFiles: true })) {
	if (entry.includes(".DS_Store") || entry.endsWith(".gitkeep")) continue;
	harnessFiles[entry] = await sha256(resolve(harnessDir, entry));
}
const fixtures = {} as Record<string, string>;
const scenariosDir = resolve(harnessDir, "skill/scenarios");
for await (const entry of new Bun.Glob("**/*").scan({ cwd: scenariosDir, onlyFiles: true })) {
	if (entry.includes(".DS_Store") || entry.endsWith(".gitkeep")) continue;
	fixtures[entry] = await sha256(resolve(scenariosDir, entry));
}
await Bun.write(resolve(root, "manifest.json"), `${JSON.stringify({
	createdAt: new Date().toISOString(),
	docx: { version, binary, sha256: await sha256(binary), commit: (await command(["git", "rev-parse", "HEAD"])).trim(), status, patch: "worktree.patch" },
	officecli: office, fixtures,
	harness: { scriptPath, files: harnessFiles, workflowSha256: await sha256(scriptPath),
		agentSha256: await sha256(resolve(harnessDir, "weak-exercise.md")) },
	model: "haiku", judge: "opus", repetitions: 3,
	status: "prepared; no model runs or quality results yet",
}, null, 2)}\n`);
// Alternate order across repetitions to reduce systematic order/cache effects.
for (let repeat = 1; repeat <= 3; repeat++) {
	const arms = repeat % 2 ? ["docx-cli", "officecli"] : ["officecli", "docx-cli"];
	for (const arm of arms) {
		const runDir = resolve(root, `${arm}-r${repeat}`);
		await mkdir(runDir);
		const args = { runDir, binary, scenariosDir, model: "haiku", arm,
			...(arm === "officecli" ? { competitorDir: office.source, competitorBinary: office.binary } : {}),
		};
		await Bun.write(resolve(root, `${arm}-r${repeat}.json`), `${JSON.stringify(args, null, 2)}\n`);
	}
}
await writeStdout(`Prepared six fresh run configurations in ${root}\n`);

async function sha256(path: string) {
	return new Bun.CryptoHasher("sha256").update(await Bun.file(path).arrayBuffer()).digest("hex");
}

async function command(argv: string[]) {
	const child = Bun.spawn(argv, { cwd: repo, stdout: "pipe", stderr: "pipe" });
	const [stdout, stderr, status] = await Promise.all([
		new Response(child.stdout).text(), new Response(child.stderr).text(), child.exited,
	]);
	if (status !== 0) throw new Error(`${argv[0]} failed (${status}): ${stderr}`);
	return stdout;
}
