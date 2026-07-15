import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/task-lists-tracked.docx via the public CLI — tracked
 * checkbox toggles emitted by `tasks check`/`uncheck` under track-changes. The shape
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
 *   1. The CLI can emit the canonical Word toggle shape (tasks add + track on +
 *      tasks check/uncheck).
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
	"tasks",
	"add",
	out,
	"--after",
	"p0",
	"--unchecked",
	"--text",
	"buy groceries",
);
await cli(
	"tasks",
	"add",
	out,
	"--after",
	"p1",
	"--unchecked",
	"--text",
	"pay rent",
);
await cli(
	"tasks",
	"add",
	out,
	"--after",
	"p2",
	"--checked",
	"--text",
	"call dentist",
);

// Turn tracking on, then toggle p2 and p3 — these are the changes we want
// surfaced as `checkboxToggle` tcN entries.
await cli("track-changes", out, "on");
process.env.DOCX_CLI_NOW ??= "2026-05-22T23:01:00Z";
await cli("tasks", "check", out, "--at", "p2", "--author", "Kirill Klimuk");
await cli("tasks", "uncheck", out, "--at", "p3", "--author", "Kirill Klimuk");

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
console.log("  tracked toggles: p2 (☐→☒), p3 (☒→☐)");
