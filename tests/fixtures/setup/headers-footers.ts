import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps so rebuilds are byte-deterministic. The DATE field's
// cached run is empty (the viewer recomputes on open), so no live date leaks in.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/headers-footers.docx — a two-page report that exercises
 * the headers/footers (marginals) surface end to end, dogfooding the CLI:
 *
 *   - a DEFAULT two-zone header: "Acme Corporation" (left) + an auto DATE field
 *     (right), separated by a content-edge right tab — `docx headers set --text … --date`;
 *   - a FIRST-PAGE header "Annual Report 2026" (sets <w:titlePg/>) — `--first-page`;
 *   - a DEFAULT footer with centered "Page X of Y" (PAGE + NUMPAGES fields) —
 *     `docx footers set --page-number --of-pages`.
 *
 * Two physical pages (a page break splits the body) so the first-page header
 * differs from the default on page 2 and the page-count field is meaningful.
 * Built entirely through the surface verbs, so building it is an end-to-end smoke
 * test for `headers set` / `footers set` (default, first-page, two-zone, fields).
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/headers-footers.docx");
const cliEntry = resolve(root, "src/index.ts");

let lastP = 0;

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

async function appendText(
	text: string,
	opts: { style?: string } = {},
): Promise<void> {
	const args = ["insert", out, "--after", `p${lastP}`, "--text", text];
	if (opts.style) args.push("--style", opts.style);
	await cli(...args);
	lastP += 1;
}

async function appendPageBreak(): Promise<void> {
	await cli("insert", out, "--after", `p${lastP}`, "--page-break");
	lastP += 1;
}

mkdirSync(dirname(out), { recursive: true });

await cli(
	"create",
	out,
	"--title",
	"Acme Annual Report",
	"--author",
	"docx-cli",
	"--force",
);

await cli(
	"edit",
	out,
	"--at",
	"p0",
	"--text",
	"Acme Annual Report 2026",
	"--style",
	"Heading1",
);

await appendText(
	"This report opens on a cover page with its own first-page header, while every following page carries the standard running header (company name on the left, the date on the right) and a centered page-number footer.",
);
await appendPageBreak();
await appendText(
	"The body continues onto a second page so the running header, the first-page header, and the page-count footer are all visible and distinct in a render.",
);

// Default running header — two-zone: company name (left) + auto date (right).
await cli("headers", "set", out, "--text", "Acme Corporation", "--date");
// Different first page (sets <w:titlePg/> on the section).
await cli(
	"headers",
	"set",
	out,
	"--first-page",
	"--text",
	"Annual Report 2026",
);
// Centered "Page X of Y" footer (PAGE + NUMPAGES fields).
await cli("footers", "set", out, "--page-number", "--of-pages");

// Verify and report.
const verifyJson = await cli("read", out, "--ast");
const doc = JSON.parse(verifyJson) as {
	headers: Array<{ id: string; type: string; text: string }>;
	footers: Array<{ id: string; type: string; text: string }>;
};
console.log(`Wrote ${out}`);
console.log("Headers:");
for (const header of doc.headers) {
	console.log(`  ${header.id}: type=${header.type} text="${header.text}"`);
}
console.log("Footers:");
for (const footer of doc.footers) {
	console.log(`  ${footer.id}: type=${footer.type} text="${footer.text}"`);
}
