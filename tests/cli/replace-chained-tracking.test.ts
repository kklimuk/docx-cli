import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// Repro of the agent-feedback case: under track-changes ON, two consecutive
// replace calls in the same paragraph used to corrupt offsets — the second
// match's start was computed against a string that included the first
// replace's <w:ins>, so the splice landed mid-word inside the inserted run.
//
// The default (accepted) view fix: replace's offsets ignore the just-emitted
// <w:ins> and existing <w:del> wrappers, so chained edits stay safe.

const FIXTURE = "tests/fixtures/chained-tracked-edits.docx";
// Layout (built by scripts/make-chained-tracked-edits-fixture.ts):
//   p0: "Cost of living, anti-price-gouging, and housing reform."
//   p1: "Old plan: ship Tuesday."
//   track-changes: ON, no tracked changes recorded yet.

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

describe("docx replace — chained edits under tracking", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("chained-replace");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));
	});

	test("two replaces in the same paragraph keep offsets stable in accepted view", async () => {
		const first = await runCli(
			"replace",
			docPath,
			"Cost of living",
			"Affordability",
		);
		expect(first.exitCode).toBe(0);

		// The second pattern is a phrase to the right of the first edit.
		// In the buggy version, the second replace's offset would be computed
		// against a haystack that included the just-inserted "Affordability"
		// AND the still-present <w:del>"Cost of living", landing the splice
		// inside the <w:ins>. With the accepted-view fix, neither the
		// pre-existing <w:del> nor the new <w:ins> shifts subsequent offsets.
		const second = await runCli(
			"replace",
			docPath,
			"anti-price-gouging",
			"price control",
		);
		expect(second.exitCode).toBe(0);

		// Read the accepted view of just p0 — under accepted view the
		// <w:del>s are dropped and the <w:ins>s are inlined as plain text.
		const result = await runCli("read", docPath, "--from", "p0", "--to", "p0");
		expect(result.exitCode).toBe(0);
		const accepted = result.stdout
			.split("\n")
			.map((line) => line.replace(/\s*<!--\s*[a-z0-9]+\s*-->\s*$/, ""))
			.join("\n")
			.trim();
		expect(accepted).toBe("Affordability, price control, and housing reform.");
	});

	test("--current view (legacy behavior) sees the raw concatenation", async () => {
		// Use p1 of the fixture: "Old plan: ship Tuesday."
		await runCli("replace", docPath, "Old", "New");

		// In --current view, find sees both ins and del text, so the next
		// query against "Old" still matches the deleted run.
		const find = await runCli("find", docPath, "Old", "--current");
		const payload = find.parsed as {
			matches: Array<{
				blockId: string;
				trackedChanges?: Array<{ kind: string }>;
			}>;
		};
		expect(payload.matches).toHaveLength(1);
		expect(payload.matches[0]?.blockId).toBe("p1");
		expect(payload.matches[0]?.trackedChanges?.[0]?.kind).toBe("del");

		// The default (accepted) view, by contrast, no longer sees "Old".
		const findDefault = await runCli("find", docPath, "Old");
		const defaultPayload = findDefault.parsed as { matches: unknown[] };
		expect(defaultPayload.matches).toEqual([]);

		// And the accepted view of p1 reads cleanly.
		expect(await paragraphText(docPath, "p1")).toContain("New plan");
	});
});
