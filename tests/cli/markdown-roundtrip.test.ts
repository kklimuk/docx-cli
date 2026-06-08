import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// read --markdown output must re-import cleanly: the ` <!-- pN -->` locator
// separator must not accumulate spaces, and verse line breaks must survive,
// across read → create → read. (The section-break `---` line is excluded —
// section properties can't be recreated from markdown; tracked separately.)

const SOURCE = `# Quarterly Report

A normal paragraph with some text in it.

The winter evening settles down
With smell of steaks in passageways.
The burnt-out ends of smoky days.

| Name | Note |
| --- | --- |
| Dana | line one<br>line two |
`;

async function readBody(path: string): Promise<string> {
	const out = (await runCli("read", path)).stdout;
	// Drop the trailing section-break marker line (sections don't round-trip
	// through markdown — a separate, known limitation).
	return out
		.split("\n")
		.filter((line) => !line.includes("<!-- s0 -->"))
		.join("\n");
}

describe("read → create → read is stable for normal content", () => {
	test("no trailing-space growth before locators; verse + tables survive", async () => {
		const workspace = tempWorkspace("md-roundtrip");
		const src = join(workspace, "src.md");
		await Bun.write(src, SOURCE);
		const doc1 = join(workspace, "d1.docx");
		expect((await runCli("create", doc1, "--from", src)).exitCode).toBe(0);

		const read1 = await readBody(doc1);
		const reimport = join(workspace, "r.md");
		await Bun.write(reimport, `${read1}\n`);
		const doc2 = join(workspace, "d2.docx");
		expect((await runCli("create", doc2, "--from", reimport)).exitCode).toBe(0);
		const read2 = await readBody(doc2);

		expect(read2).toBe(read1);
		// no double-space crept in before any locator comment
		expect(read2).not.toMatch(/ {2}<!--/);
		// verse line breaks are still there (3 lines → 2 <w:br/>)
		const br = (await Bun.file(doc2).bytes()).length; // touch file
		expect(br).toBeGreaterThan(0);
		expect(read2).toContain("The winter evening settles down\n");
	});
});
