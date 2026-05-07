import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/sections.docx — a multi-section doc that exercises
 * every supported section type with the surrounding context required to make
 * the behavior visible in a viewer.
 *
 * OOXML semantics reminder (ECMA-376 §17.6.22): the w:type element on a
 * w:sectPr describes where THIS section BEGINS, not where the next one starts.
 *   - continuous: this section begins on the same page as the prior section
 *   - nextPage:   this section begins on a new page
 *   - nextColumn: this section begins at the top of the next column
 *   - evenPage:   this section begins on the next even-numbered page (with
 *                 a blank odd page inserted if the natural next page is odd)
 *   - oddPage:    this section begins on the next odd-numbered page (with
 *                 a blank even page inserted if the natural next page is even)
 *
 *   s0: continuous,  1 col    — front matter (title + doc intro) in a normal
 *                               single-column layout. Without this boundary,
 *                               Section A's s1 properties would extend
 *                               backward and render the title in two columns.
 *   s1: continuous,  2 cols   — Section A: starts on same page as setup;
 *                               mid-page flip to 2-col layout.
 *   s2: nextColumn,  2 cols   — Section B: starts at the top of the next
 *                               available column (the leftover space in
 *                               page 1's right column is skipped).
 *   s3: continuous,  2 cols   — Section C: starts continuously after B in
 *                               the same column-flow context.
 *   s4: nextPage,    1 col    — Section D: starts on a fresh page; column
 *                               count drops to 1.
 *   s5: evenPage,    1 col    — Section E: starts on next even page. Since
 *                               D ends on page 4, the natural next is page 5
 *                               (odd), so a blank odd page 5 is inserted and
 *                               E begins on page 6.
 *   s6: oddPage,     1 col    — Section F: starts on next odd page. E ends
 *                               on page 7 (odd), natural next is page 8
 *                               (even), so a blank even page 8 is inserted
 *                               and F begins on page 9.
 *   s7: trailing,    2 cols + <w:sectPrChange> snapshot of prior 1-col state
 *                    (Section G). type=continuous, so G begins on the same
 *                    page F ends on (page 10): F's last paragraph in 1-col,
 *                    then a mid-page flip to 2-col for G's content.
 *
 * Real page count: 10 physical pages. Page 5 and page 8 are blank fillers
 * inserted by the viewer to satisfy evenPage / oddPage parity.
 *
 * Sections D, E, F each include an explicit page-break paragraph so each
 * spans two physical pages, giving evenPage and oddPage deterministic
 * starting parities. Without that padding, the natural pagination might
 * already line up with the requested parity and no blank filler would be
 * inserted.
 *
 * Built by dogfooding the CLI; doubles as an end-to-end smoke test for
 * `insert --section`, `insert --page-break`, `edit --at sN`, and
 * `track-changes`.
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/sections.docx");
const cliEntry = resolve(root, "src/index.ts");

let lastP = 0;

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

async function cliEnv(
	env: Record<string, string>,
	...args: string[]
): Promise<string> {
	const result = await $.env(env)`bun ${cliEntry} ${args}`.quiet();
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

async function appendSection(opts: {
	columns?: number;
	type?: string;
}): Promise<void> {
	const args = ["insert", out, "--after", `p${lastP}`, "--section"];
	if (opts.columns !== undefined) args.push("--columns", String(opts.columns));
	if (opts.type !== undefined) args.push("--type", opts.type);
	await cli(...args);
	lastP += 1;
}

mkdirSync(dirname(out), { recursive: true });

await cli(
	"create",
	out,
	"--title",
	"Sections fixture",
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
	"Sections fixture",
	"--style",
	"Heading1",
);

await appendText(
	"This document exercises every section type the CLI supports — continuous, nextColumn, nextPage, evenPage, oddPage — alongside a w:sectPrChange revision marker on the trailing section. Each section opens with a heading naming what to look for and includes enough body text and structural padding to make the behavior visible in Word or LibreOffice.",
);

// === Setup section: 1 column ======================================
// OOXML semantics: an inline w:pPr/w:sectPr defines the section ENDING at
// that paragraph, applying backwards to every paragraph since the prior
// boundary. Without this setup boundary, Section A's s0 (columns=2) would
// govern the title and doc-intro paragraphs above as well, rendering them
// in two columns. This explicit 1-column boundary keeps the front-matter
// in a normal single-column layout.
await appendSection({ columns: 1, type: "continuous" });

// === Section A: continuous, 2 columns ============================
await appendText("Section A — continuous, two columns (mid-page column flip)", {
	style: "Heading2",
});
await appendText(
	"This section uses type=continuous, which means Section A itself begins on the same page as the prior section (the front-matter setup section above). Because the column count flips from 1 to 2 across that boundary without a page break, the new column layout takes effect immediately on this same page. The visible signal is the mid-page transition you see right above this paragraph: single-column front-matter, then this two-column body.",
);
await appendText(
	"To make two-column flow visible, this section needs enough body text to span both columns of the layout. Word fills the left column down to the bottom margin, then resumes at the top of the right column. Without sufficient content, you would only see a single short column of text and miss the column-flow behavior entirely. The exact split point varies with font metrics, page size, and margin settings.",
);
await appendText(
	"Note: the w:type element describes where the CURRENT section begins, not where the next section starts (per ECMA-376 §17.6.22). It is easy to mis-read this — the same value can imply different things depending on which side of the boundary you have in mind. Throughout this fixture the body text is written from the current-section's perspective: continuous means this section continues on the prior page; nextPage means this section starts on a new page; etc.",
);
await appendSection({ columns: 2, type: "continuous" });

// === Section B: nextColumn, 2 columns =============================
// Section B's type=nextColumn forces B itself to start at the top of the
// next available column. Section A's content typically does not fill all
// columns of page 1, so the leftover white space at the bottom of page 1's
// right column is skipped and B's heading appears at the top of page 2's
// left column. With type=continuous instead, B would have started in that
// leftover white space.
await appendText(
	"Section B — nextColumn, two columns (jumps to top of next available column)",
	{ style: "Heading2" },
);
await appendText(
	"This section uses type=nextColumn, which means SECTION B itself begins at the top of the next available column. Section A's body did not completely fill the columns of page 1 (you can see the leftover white space at the bottom of page 1's right column above), so under nextColumn semantics that white space is skipped and B's heading lands at the top of the next fresh column — typically the left column of page 2. With type=continuous on B instead, this heading would have started in that leftover white space.",
);
await appendSection({ columns: 2, type: "nextColumn" });

// === Section C: continuous, 2 columns =============================
await appendText(
	"Section C — continuous, two columns (flows directly after B in the same column)",
	{ style: "Heading2" },
);
await appendText(
	"Section C uses type=continuous, so C begins on the same page as Section B, in the same column-flow context. C's heading therefore appears directly after B's body in the column, with no break between them. C is the visual contrast to B: continuous flows on, nextColumn jumps. Both share the same two-column layout.",
);
await appendText(
	"Section D below changes the column count back to one. Even though we will not explicitly mark D's break as nextPage on B's side, a column-count change typically forces a page break in most viewers regardless — the layout simply cannot reflow across a column-count boundary on the same page. Section D's own type is nextPage, so the page-break behavior is explicit anyway.",
);
await appendSection({ columns: 2, type: "continuous" });

// === Section D: nextPage, 1 column ================================
// Two paragraphs separated by an explicit page-break paragraph so D
// spans two physical pages. This controls the parity for the downstream
// evenPage / oddPage demonstrations.
await appendText(
	"Section D — nextPage, single column (starts on a fresh page)",
	{
		style: "Heading2",
	},
);
await appendText(
	"This section uses type=nextPage, the default and most familiar section break type. nextPage means SECTION D begins on a fresh page. We have also dropped the column count back to one, so D's body flows in a single column the full page width. The transition from C's two-column layout to D's single-column layout combines with the nextPage signal to give a clean fresh-page start.",
);
// Two-page padding to control parity for s4/s5.
await appendPageBreak();
await appendText(
	"This second paragraph of Section D lives on the next physical page because the empty paragraph above carries an inline page-break run (w:br w:type=page). The reason for this padding: Section D needs to reliably span two pages so the evenPage / oddPage parity-jumps downstream have a deterministic ending parity to react to. Without this control, the natural pagination might already line up with the requested parity and no blank filler page would be inserted.",
);
await appendSection({ columns: 1, type: "nextPage" });

// === Section E: evenPage, 1 column ================================
await appendText(
	"Section E — evenPage, single column (begins on next even page)",
	{
		style: "Heading2",
	},
);
await appendText(
	"This section uses type=evenPage. evenPage means SECTION E itself begins on the next even-numbered page. Section D ended on page 4 (an even page), so the natural next page is page 5 (odd). To honor evenPage, the viewer inserts a blank page 5 (odd) and lands E on page 6 (even). If you scrolled past a blank page just before reaching this heading, that is the inserted parity-filler. Use evenPage when chapter starts must always fall on a left-hand verso page in print layouts.",
);
await appendPageBreak();
await appendText(
	"This second paragraph of Section E lives on the next page so the section spans two physical pages, ending on page 7 (odd). The two-page span gives Section F downstream a deterministic starting context: F's oddPage will need to advance from this section's odd ending to the next odd page, which requires inserting another blank to skip past the natural even page in between.",
);
await appendSection({ columns: 1, type: "evenPage" });

// === Section F: oddPage, 1 column =================================
await appendText(
	"Section F — oddPage, single column (begins on next odd page)",
	{
		style: "Heading2",
	},
);
await appendText(
	"This section uses type=oddPage, the symmetric counterpart to evenPage. oddPage means SECTION F itself begins on the next odd-numbered page. Section E ended on page 7 (odd), so the natural next page is page 8 (even). To honor oddPage, the viewer inserts a blank page 8 (even) and lands F on page 9 (odd). If you scrolled past a blank page just before reaching this heading, that is the inserted parity-filler. Use oddPage when chapter starts must always fall on a right-hand recto page in print layouts.",
);
await appendPageBreak();
await appendText(
	"This second paragraph of Section F lives on the next page, ending Section F on page 10 (even). Section G that follows uses the trailing section's type=continuous, which means G begins on the same page F ends on — page 10 — with the column count flipping from 1 to 2 mid-page. That gives the same kind of mid-page column transition we saw in Section A at the start of the document. Together evenPage and oddPage let you ensure consistent recto / verso starts in printed documents.",
);
await appendSection({ columns: 1, type: "oddPage" });

// === Section G (trailing): 2 columns with a sectPrChange snapshot ===
await appendText(
	"Section G — trailing section, two columns (with a sectPrChange snapshot)",
	{ style: "Heading2" },
);
await appendText(
	"This is the trailing section, defined by the w:sectPr at the end of w:body rather than by an inline sectPr inside a paragraph. Every OOXML document must have exactly one trailing sectPr; it cannot be deleted. The trailing section uses type=continuous, so Section G begins on the same page Section F ends on. Combined with the column-count change from 1 to 2, that produces the mid-page 1-to-2 column flip you see at the top of this section.",
);
await appendText(
	"To exercise tracked revisions, we first set this trailing section to one column with track-changes off, then enabled tracking and edited it back to two columns. That edit recorded a w:sectPrChange snapshot inside the live sectPr. Running docx track-changes list on this fixture surfaces the revision with its prior and current properties, so agents can see what changed without inspecting raw XML.",
);
await appendText(
	"The sectPrChange marker has author Reviewer and a fixed date set via DOCX_CLI_NOW for reproducibility. To accept the revision and remove the snapshot while keeping the new two-column layout, run docx track-changes accept FILE --all. To reject and restore the prior single-column state, run docx track-changes reject FILE --all. Either operation makes the file structurally stable for downstream consumers.",
);

// Initialize the trailing sectPr to columns=1 (untracked baseline state).
await cli("edit", out, "--at", "s7", "--columns", "1", "--type", "continuous");
// Turn tracking on, edit to columns=2 — this leaves a w:sectPrChange capturing
// the prior 1-column state inside the live sectPr.
await cli("track-changes", out, "on");
await cliEnv(
	{ DOCX_AUTHOR: "Reviewer", DOCX_CLI_NOW: "2026-05-06T10:00:00Z" },
	"edit",
	out,
	"--at",
	"s7",
	"--columns",
	"2",
);
await cli("track-changes", out, "off");

// Verify and report
const verifyJson = await cli("read", out);
const doc = JSON.parse(verifyJson) as {
	blocks: Array<{
		id: string;
		type: string;
		columns?: number;
		sectionType?: string;
	}>;
};
const sections = doc.blocks.filter((b) => b.type === "sectionBreak");
console.log(`Wrote ${out}`);
console.log("Sections:");
for (const section of sections) {
	console.log(
		`  ${section.id}: columns=${section.columns ?? "(default)"} type=${section.sectionType ?? "(default)"}`,
	);
}
const trackedJson = await cli("track-changes", "list", out);
const tracked = JSON.parse(trackedJson) as Array<{
	id: string;
	kind: string;
	blockId: string;
	prior?: { columns?: number; sectionType?: string };
	current?: { columns?: number; sectionType?: string };
}>;
console.log(`Tracked changes: ${tracked.length}`);
for (const change of tracked) {
	console.log(
		`  ${change.id}: ${change.kind} on ${change.blockId} (prior=${JSON.stringify(change.prior ?? {})}, current=${JSON.stringify(change.current ?? {})})`,
	);
}
