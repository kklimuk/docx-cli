import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";
import { wordTextBoxParagraphXml } from "../../cli/helpers";

// Pin core.xml timestamps so rebuilds are byte-deterministic.
process.env.DOCX_CLI_NOW ??= "2026-09-04T00:00:00Z";

/**
 * Build tests/fixtures/text-boxes.docx — the issue #4 document: two flow
 * paragraphs plus a floating text box ("CONFIDENTIAL" stamp) written the way
 * Word writes every modern text box — an `<mc:AlternateContent>` pair holding
 * a `wps` shape under `<mc:Choice>` and a VML `<w:pict>` twin under
 * `<mc:Fallback>`, the SAME story in both.
 *
 * Dogfoods the surface end to end: the box enters through `raw insert` (the
 * only authoring path for a new box — its geometry is unmodeled), then every
 * ordinary verb addresses the story through `tbxN:pK`: `replace --all` reaches
 * the third "Acme" (the silent-under-application shape from the issue),
 * `edit --at tbx0:p0` rewrites the stamp, `comments add --at tbx0:p1:S-E`
 * is refused inside the box (Word drops it — the note lands on p0 instead). The save re-syncs the Fallback twin from the Choice
 * story, so the LibreOffice round-trip (which reads the wps Choice) and any
 * legacy reader (which takes the VML Fallback) see the same edited text.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/text-boxes.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

mkdirSync(dirname(out), { recursive: true });

await cli(
	"create",
	out,
	"--title",
	"Text boxes",
	"--author",
	"docx-cli",
	"--force",
	"--text",
	"Acme Corporation shall deliver the goods by Friday.",
);
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--text",
	"Payment is due to Acme within thirty days of delivery.",
);

// The floating stamp, anchored between the two flow paragraphs.
const fragment = resolve(dirname(out), "setup/.text-box-fragment.xml");
await Bun.write(
	fragment,
	wordTextBoxParagraphXml([
		"CONFIDENTIAL",
		"Confidential - Acme Corporation internal use only.",
	]),
);
await cli("raw", "insert", out, "--after", "p0", "--xml-file", fragment);
await $`rm -f ${fragment}`.quiet();

// Issue #4's shape: three occurrences on the page, one inside the box.
await cli("replace", out, "Acme", "Globex", "--all");
// Address the story directly.
await cli("edit", out, "--at", "tbx0:p0", "--text", "CONFIDENTIAL — DRAFT");
// Word keeps no comment INSIDE a text box (it drops it on save), so the
// review note goes on the anchor paragraph's neighbor instead.
await cli(
	"comments",
	"add",
	out,
	"--at",
	"p0:0-6",
	"--text",
	"Confirm the party name before sending (see the stamp).",
	"--author",
	"docx-cli",
);

const bytes = (await Bun.file(out).arrayBuffer()).byteLength;
console.log(`Wrote ${out} (${bytes} bytes)`);
