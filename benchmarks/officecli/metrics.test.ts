import { expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

test("metrics exclude same-model stage/render agents and retain legacy exercises", async () => {
	const root = await mkdtemp(join(tmpdir(), "docx-metrics-"));
	try {
		const transcripts = join(root, "transcripts");
		const run = join(root, "run");
		await mkdir(transcripts);
		await mkdir(run);
		for (const [name, description, prompt] of [
			["exercise", "exercise:mnda", "Your task (scenario: mnda): edit mnda.docx"],
			["render", "render:mnda", "Render mnda.docx"],
			["stage", "stage:inputs", "Stage mnda.docx and invoice.docx"],
			["legacy", null, "Your task (scenario: invoice): edit invoice.docx"],
			["legacy-render", null, "Render invoice.docx"],
		] as const) {
			const path = join(transcripts, `agent-${name}`);
			if (description) await Bun.write(`${path}.meta.json`, JSON.stringify({ description }));
			const usage = { input_tokens: 100, output_tokens: 10, cache_read_input_tokens: 20, cache_creation_input_tokens: 30 };
			await Bun.write(`${path}.jsonl`, [
				{ type: "user", message: { role: "user", content: prompt } },
				{ type: "assistant", message: { role: "assistant", id: name, model: "claude-haiku-test", usage, content: [{ type: "text", text: "Done." }] } },
				{ type: "assistant", message: { role: "assistant", id: name, model: "claude-haiku-test", usage, content: [{ type: "text", text: "OK." }] } },
			].map((entry) => JSON.stringify(entry)).join("\n"));
		}
		const collector = new URL("../../.claude/skills/weak-agent-test/scripts/exercise-metrics.ts", import.meta.url).pathname;
		const child = Bun.spawn(["bun", collector, transcripts, run, "/tools/docx", "haiku"], { stdout: "pipe", stderr: "pipe" });
		const [status, , stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()]);
		expect(status, stderr).toBe(0);
		const metrics = await Bun.file(join(run, "exercise-metrics.json")).json();
		expect(metrics.perScenario.map((row: { scenario: string }) => row.scenario)).toEqual(["invoice", "mnda"]);
		expect(metrics.totals.agents).toBe(2);
		expect(metrics.totals.inputTokens).toBe(200);
		expect(metrics.totals.cacheWriteTokens).toBe(60);
		expect(metrics.totals.cacheReadTokens).toBe(40);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});
