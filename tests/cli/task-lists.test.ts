import { describe, expect, test } from "bun:test";
import { copyFileSync } from "node:fs";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = join(import.meta.dir, "..", "fixtures", "task-lists.docx");

type Block = {
	id: string;
	type: string;
	list?: { level: number; numId: string };
	taskState?: "checked" | "unchecked";
	runs?: Array<{ type: string; text?: string }>;
};

async function blocksOf(docPath: string): Promise<Block[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

function textOf(block: Block): string {
	return (block.runs ?? []).map((run) => run.text ?? "").join("");
}

describe("task lists — read", () => {
	test("AST: leading SDT checkbox populates taskState; SDT + space stripped from runs", async () => {
		const blocks = await blocksOf(FIXTURE);
		const tasks = blocks.filter(
			(b) => b.type === "paragraph" && b.taskState !== undefined,
		);
		// Three top-level tasks + one nested.
		expect(tasks).toHaveLength(4);

		const states = tasks.map((b) => b.taskState);
		expect(states).toEqual(["unchecked", "checked", "unchecked", "checked"]);

		// Runs carry only the task text — no leading space, no ☐/☒ glyph.
		const texts = tasks.map(textOf);
		expect(texts).toEqual([
			"buy groceries",
			"pay rent",
			"call dentist",
			"nested done item",
		]);
	});

	test("plain list paragraphs without SDT are NOT marked as tasks", async () => {
		const blocks = await blocksOf(FIXTURE);
		const plainBullet = blocks.find(
			(b) =>
				b.type === "paragraph" &&
				b.taskState === undefined &&
				b.list !== undefined,
		);
		expect(plainBullet).toBeDefined();
		expect(textOf(plainBullet as Block)).toBe("regular reminder");
	});

	test("markdown render emits GFM `- [ ]` / `- [x]`", async () => {
		const result = await runCli("read", FIXTURE);
		expect(result.stdout).toContain("- [ ] buy groceries");
		expect(result.stdout).toContain("- [x] pay rent");
		expect(result.stdout).toContain("- [ ] call dentist");
		// Plain bullet — no task marker.
		expect(result.stdout).toContain("- regular reminder");
		// Nested + checked, indented two spaces (level 1).
		expect(result.stdout).toContain("  - [x] nested done item");
	});
});

describe("task lists — round-trip", () => {
	test("save + re-read preserves every checkbox", async () => {
		const workspace = tempWorkspace("task-list-rt");
		const docPath = join(workspace, "task-lists.docx");
		copyFileSync(FIXTURE, docPath);

		// Mutate to force a Document.save path; the SDT lives in the XmlNode tree
		// and should ride through the re-serializer untouched.
		await runCli("insert", docPath, "--after", "p5", "--text", "after task");

		const blocks = await blocksOf(docPath);
		const tasks = blocks.filter(
			(b) => b.type === "paragraph" && b.taskState !== undefined,
		);
		expect(tasks.map((b) => b.taskState)).toEqual([
			"unchecked",
			"checked",
			"unchecked",
			"checked",
		]);
		expect(tasks.map(textOf)).toEqual([
			"buy groceries",
			"pay rent",
			"call dentist",
			"nested done item",
		]);
	});
});

describe("task lists — Word for Web shape (Wingdings ☐ bullet + strike)", () => {
	const WEB_FIXTURE = join(
		import.meta.dir,
		"..",
		"fixtures",
		"task-lists-web.docx",
	);

	test("AST: Wingdings ☐ bullet list with strike-on-pPr-mark populates taskState", async () => {
		const blocks = await blocksOf(WEB_FIXTURE);
		const tasks = blocks.filter(
			(b) => b.type === "paragraph" && b.taskState !== undefined,
		);
		// All four list items in the web-style fixture are tasks (no plain bullets
		// mixed in — Word for Web's whole list is the "checklist").
		expect(tasks).toHaveLength(4);
		expect(tasks.map((b) => b.taskState)).toEqual([
			"unchecked",
			"checked",
			"unchecked",
			"checked",
		]);
		expect(tasks.map(textOf)).toEqual([
			"buy groceries",
			"pay rent",
			"call dentist",
			"nested done item",
		]);
	});

	test("markdown render emits GFM `- [ ]` / `- [x]` even for Wingdings-bullet shape", async () => {
		const result = await runCli("read", WEB_FIXTURE);
		expect(result.stdout).toContain("- [ ] buy groceries");
		expect(result.stdout).toContain("- [x]");
		expect(result.stdout).toContain("pay rent");
	});
});

describe("task lists — tracked checkbox toggles", () => {
	const TRACKED_FIXTURE = join(
		import.meta.dir,
		"..",
		"fixtures",
		"task-lists-tracked.docx",
	);

	test("`track-changes list` surfaces toggles as `checkboxToggle` (one tcN per toggle)", async () => {
		const result = await runCli("track-changes", "list", TRACKED_FIXTURE);
		const changes = result.parsed as Array<{
			id: string;
			kind: string;
			blockId: string;
			author: string;
		}>;
		const toggles = changes.filter((c) => c.kind === "checkboxToggle");
		expect(toggles).toHaveLength(2);
		expect(toggles.map((t) => t.blockId)).toEqual(["p2", "p3"]);
		// Metadata is pulled from the inner <w:ins>, not the SDT itself.
		for (const t of toggles) expect(t.author).toBe("Kirill Klimuk");
	});

	test("accept a toggle keeps the new state; the SDT becomes untracked", async () => {
		const workspace = tempWorkspace("task-tracked-accept");
		const docPath = join(workspace, "out.docx");
		copyFileSync(TRACKED_FIXTURE, docPath);
		await runCli("track-changes", "accept", docPath, "--at", "tc0");
		// p2 was a ☐→☒ toggle; accept keeps the checked state.
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("- [x] pay rent");
		// No more tracking on tc0.
		const after = await runCli("track-changes", "list", docPath);
		const remaining = (
			after.parsed as Array<{ id: string; kind: string }>
		).filter((c) => c.kind === "checkboxToggle");
		expect(remaining.map((r) => r.id)).toEqual(["tc0"]); // formerly tc1, re-numbered
	});

	test("reject a toggle restores the prior glyph AND flips w14:checked back", async () => {
		const workspace = tempWorkspace("task-tracked-reject");
		const docPath = join(workspace, "out.docx");
		copyFileSync(TRACKED_FIXTURE, docPath);
		// p3 was a ☒→☐ un-toggle; reject restores ☒ (checked).
		await runCli("track-changes", "reject", docPath, "--at", "tc1");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("- [x] call dentist");
	});

	test("--all over a mix of toggles flips both correctly", async () => {
		const workspace = tempWorkspace("task-tracked-all");
		const docPath = join(workspace, "out.docx");
		copyFileSync(TRACKED_FIXTURE, docPath);
		await runCli("track-changes", "reject", docPath, "--all");
		const result = await runCli("read", docPath);
		// p2 (pay rent) was ☐→☒; reject restores ☐.
		expect(result.stdout).toContain("- [ ] pay rent");
		// p3 (call dentist) was ☒→☐; reject restores ☒.
		expect(result.stdout).toContain("- [x] call dentist");
	});
});

describe("task lists — tasks add", () => {
	test("creates a fresh task list paragraph with state, allocating a numId", async () => {
		const workspace = tempWorkspace("tasks-add-fresh");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--unchecked",
			"--text",
			"buy groceries",
		);
		const blocks = await blocksOf(docPath);
		const task = blocks.find((b) => b.id === "p1") as Block;
		expect(task.taskState).toBe("unchecked");
		expect(task.list).toBeDefined();
		expect(textOf(task)).toBe("buy groceries");
	});

	test("consecutive tasks add inherit the anchor's numId (contiguous list)", async () => {
		const workspace = tempWorkspace("tasks-add-inherit");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--unchecked",
			"--text",
			"A",
		);
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p1",
			"--checked",
			"--text",
			"B",
		);
		const blocks = await blocksOf(docPath);
		const a = blocks.find((b) => b.id === "p1") as Block;
		const b = blocks.find((b) => b.id === "p2") as Block;
		expect(a.list?.numId).toBe(b.list?.numId);
		expect(a.taskState).toBe("unchecked");
		expect(b.taskState).toBe("checked");
	});

	test("--checked / --unchecked set the state; default is unchecked", async () => {
		const workspace = tempWorkspace("tasks-add-state");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("tasks", "add", docPath, "--after", "p0", "--text", "default");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p1",
			"--checked",
			"--text",
			"done",
		);
		const blocks = await blocksOf(docPath);
		const p1 = blocks.find((b) => b.id === "p1") as Block;
		const p2 = blocks.find((b) => b.id === "p2") as Block;
		expect(p1.taskState).toBe("unchecked"); // default
		expect(p2.taskState).toBe("checked");
	});

	test("--checked and --unchecked together is rejected", async () => {
		const workspace = tempWorkspace("tasks-add-both");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const result = await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--checked",
			"--unchecked",
			"--text",
			"X",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("at most one of --checked or --unchecked");
	});

	test("without --text or --runs is rejected", async () => {
		const workspace = tempWorkspace("tasks-add-nocontent");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const result = await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--checked",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("exactly one of --text or --runs");
	});

	test("--runs authors a task item's label", async () => {
		const workspace = tempWorkspace("tasks-add-runs");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--checked",
			"--runs",
			'[{"type":"text","text":"ship it"}]',
		);
		const blocks = await blocksOf(docPath);
		const p1 = blocks.find((b) => b.id === "p1") as Block;
		expect(p1.taskState).toBe("checked");
		expect(textOf(p1)).toBe("ship it");
	});

	test("--list bullet still creates a plain list paragraph (no checkbox)", async () => {
		const workspace = tempWorkspace("insert-list-bullet");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--list",
			"bullet",
			"--text",
			"plain bullet",
		);
		const blocks = await blocksOf(docPath);
		const p1 = blocks.find((b) => b.id === "p1") as Block;
		expect(p1.list).toBeDefined();
		expect(p1.taskState).toBeUndefined();
		expect(textOf(p1)).toBe("plain bullet");
	});

	test("--list-level N nests a task item at the given level", async () => {
		const workspace = tempWorkspace("tasks-add-level");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--unchecked",
			"--text",
			"top",
		);
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p1",
			"--checked",
			"--list-level",
			"1",
			"--text",
			"nested",
		);
		const blocks = await blocksOf(docPath);
		const nested = blocks.find((b) => b.id === "p2") as Block;
		expect(nested.list?.level).toBe(1);
		expect(nested.taskState).toBe("checked");
	});

	test("under tracking surfaces as ordinary ins, the SDT survives untouched", async () => {
		const workspace = tempWorkspace("tasks-add-tracked");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"tasks",
			"add",
			docPath,
			"--after",
			"p0",
			"--unchecked",
			"--text",
			"buy groceries",
			"--author",
			"Tester",
		);
		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{ kind: string; text: string }>;
		// Two ins: runs (with task text) + paragraph-mark.
		const ins = changes.filter((c) => c.kind === "ins");
		expect(ins.length).toBeGreaterThanOrEqual(1);
		const textIns = ins.find((c) => c.text.includes("buy groceries"));
		expect(textIns).toBeDefined();
	});
});

describe("task lists — tasks check / uncheck", () => {
	test("check flips a task's state in place (untracked)", async () => {
		const workspace = tempWorkspace("tasks-check-untracked");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		// p1 in the fixture is "buy groceries" — unchecked.
		await runCli("tasks", "check", docPath, "--at", "p1");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("- [x] buy groceries");
	});

	test("uncheck clears a task's state in place (untracked)", async () => {
		const workspace = tempWorkspace("tasks-uncheck-untracked");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		// p2 in the fixture is "pay rent" — checked.
		await runCli("tasks", "uncheck", docPath, "--at", "p2");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("- [ ] pay rent");
	});

	test("check under tracking emits Word's canonical toggle XML (round-trip listable)", async () => {
		const workspace = tempWorkspace("tasks-check-tracked");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		await runCli("track-changes", docPath, "on");
		await runCli("tasks", "check", docPath, "--at", "p1", "--author", "Tester");
		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			kind: string;
			blockId: string;
			author: string;
		}>;
		const toggle = changes.find(
			(c) => c.kind === "checkboxToggle" && c.blockId === "p1",
		);
		expect(toggle).toBeDefined();
		expect(toggle?.author).toBe("Tester");
	});

	test("--track records a tracked checkboxToggle even when the doc toggle is OFF", async () => {
		const workspace = tempWorkspace("tasks-check-force-track");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		// The fixture's track-changes toggle is OFF; --track forces this one
		// invocation to record a tracked change anyway (the honesty check).
		await runCli(
			"tasks",
			"check",
			docPath,
			"--at",
			"p1",
			"--track",
			"--author",
			"Tester",
		);
		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			kind: string;
			blockId: string;
			author: string;
		}>;
		const toggle = changes.find(
			(c) => c.kind === "checkboxToggle" && c.blockId === "p1",
		);
		expect(toggle).toBeDefined();
		expect(toggle?.author).toBe("Tester");
	});

	test("on a non-task paragraph rejects with a clear hint", async () => {
		const workspace = tempWorkspace("tasks-check-bad");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		// p0 is the heading — not a task list item.
		const result = await runCli("tasks", "check", docPath, "--at", "p0");
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("requires a task-list paragraph");
	});
});

describe("task lists — command extraction redirects", () => {
	test("insert --task redirects to `docx tasks add`", async () => {
		const workspace = tempWorkspace("insert-task-redirect");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--task",
			"checked",
			"--text",
			"X",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("docx tasks add");
	});

	test("edit --task redirects to `docx tasks check`", async () => {
		const workspace = tempWorkspace("edit-task-redirect");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--task",
			"checked",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("docx tasks check");
	});

	test("edit --batch with a task key redirects to `docx tasks check`", async () => {
		const workspace = tempWorkspace("edit-batch-task-redirect");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		const result = await runCli(
			"edit",
			docPath,
			"--batch",
			'{"at":"p1","task":"checked"}',
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("docx tasks check");
	});
});

describe("task lists — schema", () => {
	test("`info schema` exposes taskState on Paragraph", async () => {
		const result = await runCli("info", "schema");
		const schema = result.parsed as {
			$defs: { Paragraph: { properties: { taskState?: unknown } } };
		};
		expect(schema.$defs.Paragraph.properties.taskState).toEqual({
			enum: ["checked", "unchecked"],
		});
	});

	test("`info schema --ts` reflects the TypeScript field", async () => {
		const result = await runCli("info", "schema", "--ts");
		expect(result.stdout).toContain('taskState?: "checked" | "unchecked";');
	});
});
