import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// `--track` forces tracked emission for one command even when the document's
// global <w:trackChanges/> toggle is OFF (the Task-2 trap: tracked corrections
// that silently weren't tracked).

async function docFrom(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const docPath = join(workspace, "out.docx");
	expect((await runCli("create", docPath, "--from", src)).exitCode).toBe(0);
	return docPath;
}

async function changes(
	docPath: string,
): Promise<Array<{ kind: string; author: string }>> {
	const list = await runCli("track-changes", "list", docPath);
	return list.parsed as Array<{ kind: string; author: string }>;
}

describe("--track forces tracked emission with the global toggle off", () => {
	test("edit --track wraps the edit in del/ins (author defaults to Reviewer)", async () => {
		const docPath = await docFrom("track-edit", "The old plan is here.\n");
		const find = await runCli("find", docPath, "old");
		const locator = (find.parsed as { matches: Array<{ locator: string }> })
			.matches[0]?.locator;
		if (!locator) throw new Error("expected a match for 'old'");
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					locator,
					"--text",
					"new",
					"--track",
				)
			).exitCode,
		).toBe(0);
		const list = await changes(docPath);
		expect(list.map((change) => change.kind).sort()).toEqual(["del", "ins"]);
		expect(list.every((change) => change.author === "Reviewer")).toBe(true);
	});

	test("replace --track records substitutions as tracked changes", async () => {
		const docPath = await docFrom(
			"track-replace",
			"Acme Corp and Acme Corp.\n",
		);
		expect(
			(
				await runCli(
					"replace",
					docPath,
					"Acme Corp",
					"Acme Industries",
					"--track",
					"--all",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).length).toBeGreaterThan(0);
	});

	test("delete --track wraps the paragraph in a tracked deletion", async () => {
		const docPath = await docFrom(
			"track-delete",
			"First para.\n\nSecond para.\n",
		);
		expect(
			(await runCli("delete", docPath, "--at", "p0", "--track")).exitCode,
		).toBe(0);
		expect((await changes(docPath)).some((c) => c.kind === "del")).toBe(true);
	});

	test("tables delete-row --track emits a rowDel revision", async () => {
		const docPath = await docFrom(
			"track-table",
			"| A | B |\n| - | - |\n| 1 | 2 |\n| 3 | 4 |\n",
		);
		expect(
			(
				await runCli(
					"tables",
					"delete-row",
					docPath,
					"--at",
					"t0:r2",
					"--track",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).some((c) => c.kind === "rowDel")).toBe(
			true,
		);
	});

	test("without --track and global off, a mutation leaves no tracked changes", async () => {
		const docPath = await docFrom("track-none", "Acme Corp.\n");
		expect(
			(
				await runCli(
					"replace",
					docPath,
					"Acme Corp",
					"Acme Industries",
					"--all",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).length).toBe(0);
	});
});
