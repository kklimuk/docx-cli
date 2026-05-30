import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/task-lists.docx via the public CLI — exercises the
 * full `insert --task` / `--list` / `--list-level` surface end-to-end. The
 * output is byte-equivalent (at the markdown-render level) to the canonical
 * fixture and validates that an agent can author every shape our reader
 * recognizes for the SDT task-list family.
 *
 * Mixed shapes in the fixture:
 *  - 3 top-level tasks (☐ / ☒ / ☐)
 *  - 1 plain bullet next to them (verifies the reader doesn't false-positive
 *    a non-task list paragraph as a task)
 *  - 1 nested task at level 1
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/task-lists.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Task list");
await cli(
	"edit",
	out,
	"--at",
	"p0",
	"--text",
	"Task list",
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
	"checked",
	"--text",
	"pay rent",
);
await cli(
	"insert",
	out,
	"--after",
	"p2",
	"--task",
	"unchecked",
	"--text",
	"call dentist",
);
await cli(
	"insert",
	out,
	"--after",
	"p3",
	"--list",
	"bullet",
	"--text",
	"regular reminder",
);
await cli(
	"insert",
	out,
	"--after",
	"p4",
	"--task",
	"checked",
	"--list-level",
	"1",
	"--text",
	"nested done item",
);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
