import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps so rebuilds are byte-deterministic.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/letter.docx — a modified-block BUSINESS LETTER, the
 * canonical document for paragraph SPACING + INDENTATION. It is the write-side
 * dogfood for Tier 1 items 3+4: every paragraph property is authored through the
 * new `insert`/`edit` flags (`--space-before`/`--space-after`/`--line-spacing`/
 * `--indent-left`/`--indent-right`/`--first-line`/`--hanging`), so BUILDING the
 * fixture exercises the feature end to end — the `word-formatted` pattern (which
 * is the run-formatting surface) applied to paragraph properties.
 *
 * Layout (modified-block):
 *   p0-p2  sender block — tight (space-after 0)
 *   p3     date — space-before/after
 *   p4-p6  recipient block — tight (space-after 0)
 *   p7     salutation — space-before/after
 *   p8-p10 body — traditional first-line indent + 1.15 line spacing
 *   p11    indented block quote of a clause — left+right indent, space-after
 *   p12    closing ("Sincerely,") — indented to ~center (modified block)
 *   p13    signature name — indented to match
 *   p14    enclosures — hanging indent
 * Then ONE range `edit` re-spaces the body to 1.5 (a realistic review pass) so
 * the build dogfoods `edit` (in place) AND `insert` (authoring), single-shot AND
 * range. No drawings/images — fully CLI-authorable, unlike resume-styling.docx.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/letter.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

mkdirSync(dirname(out), { recursive: true });

// Sender block — tight, no space between the lines.
await cli(
	"create",
	out,
	"--title",
	"Business letter",
	"--author",
	"docx-cli",
	"--text",
	"Aria Chen",
	"--force",
);
await cli(
	"insert",
	out,
	"--after",
	"p0",
	"--text",
	"500 Howard Street",
	"--space-after",
	"0",
);
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--text",
	"San Francisco, CA 94105",
	"--space-after",
	"0",
);

// Date — set off from the blocks above and below.
await cli(
	"insert",
	out,
	"--after",
	"p2",
	"--text",
	"June 23, 2026",
	"--space-before",
	"12",
	"--space-after",
	"12",
);

// Recipient block — tight.
await cli(
	"insert",
	out,
	"--after",
	"p3",
	"--text",
	"Ms. Dana Rivera",
	"--space-after",
	"0",
);
await cli(
	"insert",
	out,
	"--after",
	"p4",
	"--text",
	"General Counsel, Northwind Traders",
	"--space-after",
	"0",
);
await cli(
	"insert",
	out,
	"--after",
	"p5",
	"--text",
	"1200 Pike Street, Seattle, WA 98101",
	"--space-after",
	"0",
);

// Salutation.
await cli(
	"insert",
	out,
	"--after",
	"p6",
	"--text",
	"Dear Ms. Rivera,",
	"--space-before",
	"12",
	"--space-after",
	"12",
);

// Body — traditional first-line indent, 1.15 line spacing.
await cli(
	"insert",
	out,
	"--after",
	"p7",
	"--text",
	"Thank you for sending the revised Master Services Agreement. We have reviewed it with our team and are pleased with the overall direction.",
	"--first-line",
	"0.5in",
	"--line-spacing",
	"1.15",
);
await cli(
	"insert",
	out,
	"--after",
	"p8",
	"--text",
	"We have one remaining concern with the limitation-of-liability clause, which currently caps damages at the fees paid in the prior three months.",
	"--first-line",
	"0.5in",
	"--line-spacing",
	"1.15",
);
await cli(
	"insert",
	out,
	"--after",
	"p9",
	"--text",
	"We propose extending that window to twelve months, consistent with the rest of the agreement.",
	"--first-line",
	"0.5in",
	"--line-spacing",
	"1.15",
);

// Indented block quote of the clause under discussion.
await cli(
	"insert",
	out,
	"--after",
	"p10",
	"--text",
	'"In no event shall the Company\'s aggregate liability exceed the fees paid during the three (3) months preceding the claim."',
	"--indent-left",
	"0.5in",
	"--indent-right",
	"0.5in",
	"--space-after",
	"8",
);

// Closing + signature — indented to the page center (modified-block style),
// with room above the name for a wet signature.
await cli(
	"insert",
	out,
	"--after",
	"p11",
	"--text",
	"Sincerely,",
	"--indent-left",
	"3.5in",
	"--space-before",
	"12",
	"--space-after",
	"48",
);
await cli(
	"insert",
	out,
	"--after",
	"p12",
	"--text",
	"Aria Chen",
	"--indent-left",
	"3.5in",
);

// Enclosures — hanging indent so the wrapped list aligns under the first item.
await cli(
	"insert",
	out,
	"--after",
	"p13",
	"--text",
	"Enclosures: Master Services Agreement (redline); Statement of Work; Invoice #4471",
	"--hanging",
	"0.5in",
	"--space-before",
	"12",
);

// Dogfood `edit` (in place) AND the range path: a review pass widens the body to
// 1.5 line spacing in one call (overriding the 1.15 the body was authored with).
await cli("edit", out, "--at", "p8-p10", "--line-spacing", "1.5");

const verifyJson = await cli("read", out, "--ast");
const doc = JSON.parse(verifyJson) as {
	blocks: Array<{
		id: string;
		type: string;
		spacing?: Record<string, unknown>;
		indent?: Record<string, unknown>;
		runs?: Array<{ text?: string }>;
	}>;
};
console.log(`Wrote ${out}`);
for (const block of doc.blocks) {
	if (block.type !== "paragraph") continue;
	const text = (block.runs ?? []).map((run) => run.text ?? "").join("");
	const props = [
		block.spacing ? `spacing=${JSON.stringify(block.spacing)}` : "",
		block.indent ? `indent=${JSON.stringify(block.indent)}` : "",
	]
		.filter(Boolean)
		.join(" ");
	console.log(`  ${block.id}: "${text.slice(0, 32)}" ${props}`);
}
