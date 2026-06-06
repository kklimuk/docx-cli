import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// tracked-changes.docx layout (paragraph p0):
//   "This is a text with " (0..20, plain)
//   "two exciting "        (inside <w:ins>)
//   "insertions."          (plain, after the ins)
//
// Visible text per view:
//   current   — "This is a text with two exciting insertions." (44)
//   accepted  — same as current (ins is kept)            (44)
//   baseline  — "This is a text with insertions."        (31)
const FIXTURE = "tests/fixtures/tracked-changes.docx";

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "doc.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

describe("docx find — views", () => {
	test("default (accepted) reports the view in the response", async () => {
		const docPath = await freshCopy("find-view-default");
		const result = await runCli("find", docPath, "insertions");
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number; end: number }>;
		};
		expect(payload.view).toBe("accepted");
		expect(payload.matches[0]?.start).toBe(33);
	});

	test("--current view sees the same text as accepted when only ins is present", async () => {
		const docPath = await freshCopy("find-view-current");
		const result = await runCli("find", docPath, "insertions", "--current");
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number }>;
		};
		expect(payload.view).toBe("current");
		expect(payload.matches[0]?.start).toBe(33);
	});

	test("--baseline drops the inserted text and shifts offsets accordingly", async () => {
		const docPath = await freshCopy("find-view-baseline");
		const result = await runCli("find", docPath, "insertions", "--baseline");
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number; end: number }>;
		};
		expect(payload.view).toBe("baseline");
		// Without the 13-char "two exciting " insertion, "insertions" starts at 20.
		expect(payload.matches[0]?.start).toBe(20);
		expect(payload.matches[0]?.end).toBe(30);
	});

	test("--baseline does not match text that lives only inside <w:ins>", async () => {
		const docPath = await freshCopy("find-view-baseline-miss");
		const result = await runCli("find", docPath, "exciting", "--baseline");
		const payload = result.parsed as { matches: unknown[] };
		expect(payload.matches).toEqual([]);
	});

	test("--current and --baseline together are rejected", async () => {
		const docPath = await freshCopy("find-view-mutex");
		const result = await runCli(
			"find",
			docPath,
			"insertions",
			"--current",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});
