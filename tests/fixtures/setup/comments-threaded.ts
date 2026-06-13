import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + comment w:date so rebuilds are byte-deterministic
// (honored by `core/create::buildBlankPackage` and the comments lens via
// `resolveDate`). paraIds are pinned per-invocation below via
// DOCX_CLI_PARA_ID_SEED — the counter is per-process, so each CLI call gets a
// distinct base or the minted ids would collide.
process.env.DOCX_CLI_NOW ??= "2026-06-12T00:00:00Z";

/**
 * Build tests/fixtures/comments-threaded.docx — a CLI-authored comment thread
 * in the exact shape Word writes (verified against live Word round-trips and
 * the Word-authored comments-with-replies.docx):
 *
 *   p0: "The quick brown fox jumps over the lazy dog."
 *   c0: Alice  — root comment on "quick brown fox"
 *   c1: Bob    — reply to c0
 *   c2: Carol  — reply to c0
 *   c3: Dave   — reply addressed to c1; attaches to the thread ROOT c0
 *                (Word threads are single-level)
 *
 * Every reply carries its own commentRangeStart/End + commentReference in the
 * body (Word silently deletes unanchored replies on save — issue #1), all
 * paraIds sit inside MS-DOCX's valid range (< 0x80000000), and the thread
 * links live only in word/commentsExtended.xml.
 *
 * Built by dogfooding the CLI; doubles as the round-trip guard that
 * `comments add` + `comments reply` compose into a Word-valid thread.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/comments-threaded.docx");
const cliEntry = resolve(root, "src/index.ts");

let invocation = 0;
async function cli(...args: string[]): Promise<string> {
	const seed = (0x1000 * ++invocation).toString(16).padStart(8, "0");
	const result = await $`bun ${cliEntry} ${args}`
		.env({ ...process.env, DOCX_CLI_PARA_ID_SEED: seed })
		.quiet();
	return result.stdout.toString();
}

mkdirSync(dirname(out), { recursive: true });

await cli(
	"create",
	out,
	"--title",
	"Comments threaded fixture",
	"--author",
	"docx-cli",
	"--text",
	"The quick brown fox jumps over the lazy dog.",
	"--force",
);

await cli(
	"comments",
	"add",
	out,
	"--anchor",
	"quick brown fox",
	"--text",
	"Is this the right animal?",
	"--author",
	"Alice",
);
await cli(
	"comments",
	"reply",
	out,
	"--at",
	"c0",
	"--text",
	"Yes — it exercises every letter.",
	"--author",
	"Bob",
);
await cli(
	"comments",
	"reply",
	out,
	"--at",
	"c0",
	"--text",
	"Seconded.",
	"--author",
	"Carol",
);
await cli(
	"comments",
	"reply",
	out,
	"--at",
	"c1",
	"--text",
	"Agreed, attaching to the thread root like Word does.",
	"--author",
	"Dave",
);

const verifyJson = await cli("comments", "list", out);
const comments = JSON.parse(verifyJson) as Array<{
	id: string;
	author: string;
	parentId?: string;
}>;
const expected = [
	{ id: "c0", parent: undefined },
	{ id: "c1", parent: "c0" },
	{ id: "c2", parent: "c0" },
	{ id: "c3", parent: "c0" },
];
for (const { id, parent } of expected) {
	const found = comments.find((comment) => comment.id === id);
	if (!found || found.parentId !== parent) {
		throw new Error(
			`Fixture verification failed for ${id}: expected parentId ${parent}, got ${found?.parentId} (found: ${JSON.stringify(found)})`,
		);
	}
}

const bytes = await Bun.file(out).bytes();
console.log(`Wrote ${out} (${bytes.length} bytes)`);
console.log("Thread:");
for (const comment of comments) {
	const arrow = comment.parentId ? ` ↳ ${comment.parentId}` : "";
	console.log(`  ${comment.id} ${comment.author}${arrow}`);
}
