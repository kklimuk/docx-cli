import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// B2 — atomic batch accept/reject. `--at` is a `multiple: true` flag, so
// `accept --at tc1 --at tc2 --at tc3` resolves all targets against the
// pre-mutation tree. Mid-batch renumbering doesn't shift the still-pending
// ids out from under the agent.

const FIXTURE = "tests/fixtures/multi-tracked.docx";
// Layout (built by scripts/make-multi-tracked-fixture.ts):
//   p0: "Aleph is first."  — tc0 (del "Alpha") + tc1 (ins "Aleph")
//   p1: "Bet is second."   — tc2 (del "Beta")  + tc3 (ins "Bet")
//   p2: "Gimel is third."  — tc4 (del "Gamma") + tc5 (ins "Gimel")
// Six tracked changes, ids tc0..tc5 in document order. Picking
// non-adjacent ids (tc0, tc2, tc4) targets one wrapper per paragraph.

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "doc.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

async function listTracked(
	docPath: string,
): Promise<Array<{ id: string; kind: string }>> {
	const result = await runCli("track-changes", "list", docPath);
	return result.parsed as Array<{ id: string; kind: string }>;
}

async function paragraphText(
	docPath: string,
	blockId: string,
): Promise<string> {
	const read = await runCli("read", docPath, "--ast");
	const blocks = (
		read.parsed as {
			blocks: Array<{
				id: string;
				runs?: Array<{ type: string; text: string }>;
			}>;
		}
	).blocks;
	const block = blocks.find((candidate) => candidate.id === blockId);
	return (block?.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text)
		.join("");
}

describe("docx track-changes accept --at (batch)", () => {
	test("repeated --at accepts each id atomically against the pre-mutation tree", async () => {
		const docPath = await freshCopy("batch-accept");
		const before = await listTracked(docPath);
		// 3 replaces × 2 wrappers each = 6 tracked changes.
		expect(before).toHaveLength(6);
		const ids = before.map((change) => change.id);
		const targets = [ids[0], ids[2], ids[4]].filter(
			(id): id is string => id !== undefined,
		);
		expect(targets).toHaveLength(3);

		// Pick three non-adjacent ids that span all three paragraphs. In the
		// buggy world (one-at-a-time, no batch), accepting tc0 would shift
		// tc2/tc4's ids and the agent's pre-fetched list would be wrong.
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			targets[0] as string,
			"--at",
			targets[1] as string,
			"--at",
			targets[2] as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			applied: Array<{ id: string }>;
		};
		expect(payload.applied.map((entry) => entry.id)).toEqual(targets);

		// 3 of the original 6 changes accepted → 3 remain.
		const after = await listTracked(docPath);
		expect(after).toHaveLength(3);
	});

	test("dedupes repeated ids", async () => {
		const docPath = await freshCopy("batch-dedupe");
		const before = await listTracked(docPath);
		const firstId = before[0]?.id;
		expect(firstId).toBeDefined();

		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			firstId as string,
			"--at",
			firstId as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			applied: Array<{ id: string }>;
		};
		expect(payload.applied).toHaveLength(1);
	});

	test("rejects --at + --all combination", async () => {
		const docPath = await freshCopy("batch-mutex");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--all",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("unknown id in the batch errors atomically (no writes)", async () => {
		const docPath = await freshCopy("batch-unknown");
		const beforeText = await paragraphText(docPath, "p0");

		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--at",
			"tc99",
		);
		expect(result.exitCode).toBe(3); // NOT_FOUND
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "TRACKED_CHANGE_NOT_FOUND",
		});
		// p0 unchanged because the batch aborted before any apply.
		const afterText = await paragraphText(docPath, "p0");
		expect(afterText).toBe(beforeText);
	});

	test("reject --at also supports the multiple flag", async () => {
		const docPath = await freshCopy("batch-reject");
		const before = await listTracked(docPath);
		const idA = before[0]?.id;
		const idB = before[1]?.id;
		expect(idA).toBeDefined();
		expect(idB).toBeDefined();

		const result = await runCli(
			"track-changes",
			"reject",
			docPath,
			"--at",
			idA as string,
			"--at",
			idB as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { applied: unknown[] };
		expect(payload.applied).toHaveLength(2);
	});
});
