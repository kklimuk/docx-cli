import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { $ } from "bun";

// Pin core.xml timestamps so rebuilds are byte-deterministic.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/style-defs.docx — a small report that exercises the whole
 * `styles` authoring surface end to end, dogfooding the CLI:
 *
 *   - `docx styles set --at Heading1` restyles the built-in heading definition
 *     (navy, 13pt, bold, 12pt space-before) so every heading paragraph updates;
 *   - `docx styles create Callout` mints a NEW paragraph style (red, bold,
 *     spacing) and `docx styles create KbdKey --type character` a character style
 *     (monospace + shaded background);
 *   - `docx edit --style` applies a created style to a body paragraph;
 *   - the doc carries GERMAN Word's built-in default ids (`Standard` /
 *     `Absatz-Standardschriftart` — the one hand-rolled step, since the CLI
 *     never authors a localized template), so every provisioned baseline
 *     (Heading1, Hyperlink, the created styles) must base itself on those and
 *     never mint a second `w:default="1"` `Normal` (issue #59);
 *   - `docx styles set-default-font` sets the document font (Garamond) — writing
 *     <w:docDefaults> AND repointing word/theme/theme1.xml's font scheme. The
 *     theme-following headings (Heading1/Title/Subtitle reference the theme
 *     major) adopt Garamond; the KbdKey character style that DELIBERATELY pins
 *     Consolas is preserved.
 *
 * Built entirely through the surface verbs, so building it is itself an
 * end-to-end smoke test for `styles set` / `create` / `set-default-font` +
 * `edit --style`, and the LibreOffice round-trip (CORE_FIXTURES) proves the
 * emitted styles.xml + repointed theme are render-valid (CT_Style ordering on
 * edited + freshly-minted definitions; the docDefaults + theme font swap).
 *
 * Coverage note (deliberate): `styles set` ALSO has a weak-agent scenario (the
 * `resume` task drives it), but `styles create` is covered by THIS fixture only —
 * no scenario drives it, by decision. Minting a brand-new named style is an
 * advanced/rare need; weak agents serve "make it look like X" via `styles set` or
 * direct formatting (the eliot-journal baseline confirmed this — the authoring
 * agent restyled built-in headings, never invented a style). The render-validity
 * axis is what matters for `create`, and the round-trip here covers it. Revisit
 * (fold a custom-style beat into an authoring scenario) if a real need appears.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/style-defs.docx");
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

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--title", "Style Definitions Demo", "--force");
await localizeDefaultStyles();

// A heading (materializes Heading1) + body paragraphs that will adopt the styles.
await cli(
	"edit",
	out,
	"--at",
	"p0",
	"--text",
	"Overview",
	"--style",
	"Heading1",
);
await appendText(
	"This paragraph stays in the body style so the restyled heading stands apart.",
);
await appendText("Key takeaway: editing a style updates everywhere at once.");
await appendText("Press Enter to continue.");
await appendText("Findings", { style: "Heading1" });
// A bare URL autolinks: provisions the Hyperlink character style, which must
// base itself on the German default character style (issue #59).
await cli(
	"insert",
	out,
	"--after",
	`p${lastP}`,
	"Reference: https://example.com",
);
lastP += 1;
await appendText("A second section heading, sharing the Heading 1 style.");

// Restyle the built-in Heading 1 definition — both headings update together.
// Also sets the East-Asian-script font (separate from --font: Word resolves
// w:eastAsia independently of w:ascii/w:hAnsi) so the round-trip exercises
// <w:rFonts w:eastAsia="…"> on a style definition.
await cli(
	"styles",
	"set",
	out,
	"--at",
	"Heading1",
	"--color",
	"1F3864",
	"--size",
	"13",
	"--bold",
	"--space-before",
	"12",
	"--font-east-asia",
	"SimHei",
);

// Mint a custom paragraph style and apply it to the "Key takeaway" line.
await cli(
	"styles",
	"create",
	out,
	"Callout",
	"--name",
	"Callout",
	"--color",
	"C00000",
	"--bold",
	"--space-after",
	"6",
);
await cli("edit", out, "--at", "p2", "--style", "Callout");

// Mint a custom CHARACTER style (monospace + shaded) — applied via a run style.
await cli(
	"styles",
	"create",
	out,
	"KbdKey",
	"--type",
	"character",
	"--font",
	"Consolas",
	"--shade",
	"EEEEEE",
);

// Set the document default font — writes <w:docDefaults> AND repoints the theme
// font scheme (the only honest way to set a doc-wide font). Without --all it
// preserves the styles that pin their own font (Heading1, KbdKey).
await cli("styles", "set-default-font", out, "Garamond");

// Explicit script fonts after the default change: Latin and Arabic stay distinct.
await appendText("Arabic sample: مرحبا بالعالم");
await cli(
	"edit",
	out,
	"--at",
	`p${lastP}`,
	"--font",
	"Arial",
	"--font-complex-script",
	"Amiri",
);
await cli(
	"styles",
	"set",
	out,
	"--at",
	"Heading1",
	"--font-complex-script",
	"Amiri",
);

/** Re-id the template's default paragraph style the way German Word writes it
 *  (`Standard`, display name still `Normal`) and add its default character
 *  style `Absatz-Standardschriftart`. */
async function localizeDefaultStyles(): Promise<void> {
	const pkg = await Pkg.open(out);
	const styles = (await pkg.readText("word/styles.xml"))
		.replace('w:styleId="Normal"', 'w:styleId="Standard"')
		.replace(
			"</w:styles>",
			'<w:style w:type="character" w:default="1" w:styleId="Absatz-Standardschriftart"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>',
		);
	pkg.writeText("word/styles.xml", styles);
	await pkg.save();
}

// Verify and report.
const used = await cli("styles", out, "--used");
console.log(`Wrote ${out}`);
console.log("Styles in use:");
console.log(
	used
		.trimEnd()
		.split("\n")
		.map((line) => `  ${line}`)
		.join("\n"),
);
const heading = await cli("styles", out, "--at", "Heading1", "--json");
console.log(`Heading1 now: ${heading.trim()}`);
