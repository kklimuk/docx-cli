import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/footnotes-mutations.docx — a doc that starts with no
 * footnotes/endnotes part and exercises every `footnotes` and `endnotes` verb
 * through the public CLI. This is the round-trip we want LibreOffice to chew
 * through: a doc whose footnotes.xml is created from scratch (separator +
 * continuationSeparator + real user entries) and whose body holds three
 * footnote references and one endnote reference.
 *
 * Final shape:
 *   p0: "Footnotes fixture"
 *   p1: "Anchored at end.[^fn1]"               (footnote authored, then `edit`)
 *   p2: "Anchored at start.[^fn2]Body."        (fn2 inserted at offset 0)
 *   p3: "Anchored mid-run.[^fn3]"              (fn3 inserted mid-run)
 *   p4: "Endnote here.[^en1]"                  (endnote authored)
 *
 *   footnotes.xml: separator, continuationSeparator, fn1 (edited), fn2, fn3
 *   endnotes.xml:  separator, continuationSeparator, en1
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/footnotes-mutations.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Footnotes fixture");

await cli("insert", out, "--after", "p0", "--text", "Anchored at end.");
await cli(
	"footnotes",
	"add",
	out,
	"--at",
	"p1",
	"--text",
	"Original body fn1.",
);
await cli(
	"footnotes",
	"edit",
	out,
	"--id",
	"fn1",
	"--text",
	"Edited body fn1.",
);

await cli("insert", out, "--after", "p1", "--text", "Anchored at start. Body.");
// Offset 0 = before "Anchored…" — first character.
await cli("footnotes", "add", out, "--at", "p2:0", "--text", "Body fn2.");

await cli("insert", out, "--after", "p2", "--text", "Anchored mid-run.");
// 8 = right after "Anchored", mid-run.
await cli("footnotes", "add", out, "--at", "p3:8", "--text", "Body fn3.");

await cli("insert", out, "--after", "p3", "--text", "Endnote here.");
await cli("endnotes", "add", out, "--at", "p4", "--text", "Body en1.");

// One tracked footnote so the integration suite covers the under-tracking
// emit path too: <w:ins> wraps both the reference run AND the body content,
// matching Word's empirical shape (see `/tmp/fn-probe/add.docx`).
await cli("insert", out, "--after", "p4", "--text", "Tracked anchor here.");
await cli("track-changes", out, "on");
await cli(
	"footnotes",
	"add",
	out,
	"--at",
	"p5",
	"--text",
	"Tracked body fn4.",
	"--author",
	"docx-cli",
);
await cli("track-changes", out, "off");

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
