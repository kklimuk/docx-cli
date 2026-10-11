import { describe, expect, test } from "bun:test";

// Exercise the actual workflow's configuration and prompt builders without a paid
// model call. Function declarations remain hoisted when we stop before staging.
const source = (await Bun.file(new URL("../../.claude/workflows/weak-agent-test.js", import.meta.url)).text())
	.replace("export const meta", "const meta")
	.replace('phase("Stage");', `return {
		exercise: exercisePrompt(SCENARIOS[0]),
		render: renderPrompt({ key: "mnda", output: "/run/mnda/mnda.docx" }),
		judgeNote: ARMS[arm].judgeNote,
	};\nphase("Stage");`);
const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
const inspect = new AsyncFunction("args", "log", source);
const base = { runDir: "/run", binary: "/tools/docx", scenariosDir: "/private/scenarios" };

describe("OfficeCLI workflow arm", () => {
	test("requires both the pinned binary and matching skill source", async () => {
		await expect(inspect({ ...base, arm: "officecli" }, () => {})).rejects.toThrow("competitorBinary");
		await expect(inspect({ ...base, arm: "officecli", competitorBinary: "/tools/officecli" }, () => {})).rejects.toThrow("competitorDir");
	});

	test("uses native instructions, withholds rubric, and flushes before rendering", async () => {
		const result = await inspect({ ...base, arm: "officecli", competitorBinary: "/tools/officecli", competitorDir: "/tools/source" }, () => {});
		expect(result.exercise).toContain("/tools/source/SKILL.md");
		expect(result.exercise).toContain("/run/mnda/task.md");
		expect(result.exercise).not.toContain("/private/scenarios");
		expect(result.exercise).toContain("OFFICECLI_SKIP_UPDATE=1");
		expect(result.exercise).toContain("END your final message with ONE fenced");
		expect(result.exercise).toContain("SCRATCH FILES GO IN /run/mnda/scratch/");
		expect(result.exercise).toContain("close /run/mnda/mnda.docx");
		expect(result.render).toContain("/tools/officecli close /run/mnda/mnda.docx");
		expect(result.judgeNote).toContain("reader is not an OfficeCLI failure");
	});

	test("retains the original default and Anthropic arms", async () => {
		const defaultArm = await inspect(base, () => {});
		expect(defaultArm.exercise).toContain("/tools/docx info locators");
		expect(defaultArm.render).not.toContain("OFFICECLI_SKIP_UPDATE");
		const anthropic = await inspect({ ...base, arm: "anthropic-docx-skill", competitorDir: "/tools/anthropic" }, () => {});
		expect(anthropic.exercise).toContain("/tools/anthropic/SKILL.md");
		expect(anthropic.exercise).not.toContain("OFFICECLI_SKIP_UPDATE");
	});
});
