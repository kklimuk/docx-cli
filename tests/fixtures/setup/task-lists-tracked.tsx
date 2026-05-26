import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/task-lists-tracked.docx via the public CLI — tracked
 * checkbox toggles emitted by `edit --task` under track-changes. The shape
 * was validated empirically against Microsoft Word for Mac (probe in
 * /tmp/checkbox-track-probe/): an `<w:ins>` (new glyph) + `<w:del>` (old
 * glyph) pair INSIDE `<w:sdtContent>`, plus an in-place flip of the
 * `w14:checked` attribute. Our `checkboxToggle` TrackedChangeKind surfaces
 * the pair as a single tcN; this fixture exercises both directions:
 *
 *   - p2 "pay rent": tracked ☐ → ☒ (author marks complete)
 *   - p3 "call dentist": tracked ☒ → ☐ (author un-marks)
 *
 * Two test invariants this exercises end-to-end:
 *   1. The CLI can emit the canonical Word toggle shape (insert + track on +
 *      edit --task).
 *   2. The reader detects the result as a single checkboxToggle entry per
 *      paragraph (not two stray ins/del entries).
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/task-lists-tracked.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

// Build the base — three task items, all untracked.
await cli("create", out, "--force", "--text", "Tracked task lists");
await cli(
	"edit",
	out,
	"--at",
	"p0",
	"--text",
	"Tracked task lists",
	"--style",
	"Heading2",
);
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--task",
	"unchecked",
	"--text",
	"buy groceries",
);
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--task",
	"unchecked",
	"--text",
	"pay rent",
);
await cli(
	"insert",
	out,
	"--after",
	"p2",
	"--task",
	"checked",
	"--text",
	"call dentist",
);

// Turn tracking on, then toggle p2 and p3 — these are the changes we want
// surfaced as `checkboxToggle` tcN entries.
await cli("track-changes", out, "on");
process.env.DOCX_CLI_NOW ??= "2026-05-22T23:01:00Z";
await cli(
	"edit",
	out,
	"--at",
	"p2",
	"--task",
	"checked",
	"--author",
	"Kirill Klimuk",
);
await cli(
	"edit",
	out,
	"--at",
	"p3",
	"--task",
	"unchecked",
	"--author",
	"Kirill Klimuk",
);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
console.log("  tracked toggles: p2 (☐→☒), p3 (☒→☐)");
