import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	get: () => import("./get"),
	insert: () => import("./insert"),
	replace: () => import("./replace"),
	edit: () => import("./edit"),
	part: () => import("./part"),
};

const HELP = `docx raw — read and patch exact OOXML (the LAST-RESORT escape hatch)

Prefer the modeled verbs (edit/insert/tables/sections/comments/…): they are
tracked, validated, and readable back. Reach for raw ONLY for constructs no
other verb covers — drop caps, TOC field codes, embedded objects, content
controls inside a paragraph, exotic section/part properties.

Usage:
  docx raw get FILE --at LOCATOR                     # exact XML of a target
  docx raw edit FILE --at LOCATOR --find S --with S  # patch it in ONE call
  docx raw replace FILE --at LOCATOR --xml '…'       # swap it wholesale
  docx raw insert FILE (--after|--before) LOCATOR --xml '…'   # splice new XML
  docx raw part list|get|add|replace|edit FILE …     # package parts, by --name

Examples:
  # The patch loop — read exact XML, change the one thing, put it back:
  docx raw get report.docx --at s0
  docx raw edit report.docx --at s0 --find '<w:cols' \\
      --with '<w:lnNumType w:countBy="1"/><w:cols'   # line numbers, one call

  # Splice a construct no verb models (a drop cap):
  docx raw insert report.docx --before p0 --xml '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/></w:pPr><w:r><w:t>D</w:t></w:r></w:p>'

  # An embedded object needs three steps — part, relationship, body XML:
  docx raw part add report.docx --name word/embeddings/object1.bin --from ole.bin \\
      --content-type application/vnd.openxmlformats-officedocument.oleObject
  docx raw insert report.docx --xml '<Relationship Type="…/oleObject" Target="embeddings/object1.bin"/>'
  docx raw insert report.docx --after p3 --xml '<w:p>…<o:OLEObject r:id="rId7"/>…</w:p>'

What --at accepts (raw get/edit/replace):
  pN | tN | sN        a paragraph / table / section, as printed by docx read
  t0:r0c0:p0          a block inside a table cell (nested chains work)
  rIdN | rels         one relationship / the whole relationships listing
  Package parts are NOT locators — address them by --name via "docx raw part".

Safety (every mutation, no exceptions):
  Fragments must be well-formed; roots must be addressable (<w:p>/<w:tbl>, one
  <w:sectPr> for an sN, <Relationship/> for the rels part); child order is
  checked against ECMA-376; unknown namespace prefixes reject (known ones
  auto-declare); dangling rIds/note ids reject; colliding drawing/bookmark ids
  re-mint; and the document is schema-validated against the bundled ECMA-376
  transitional XSDs — a change may not ADD schema errors (--no-validate skips
  only that last gate). NOTHING is written when any gate fails. Raw changes
  are never tracked (--track is rejected); under document tracking they apply
  untracked plus a [docx-cli] audit comment.

Output:
  Mutations print the affected locator(s) one per line (+ "note:" lines for
  warnings); errors print {code, error, hint?} + nonzero exit. Blocks written
  raw are marked in docx read with a "raw" token on their docx:p/docx:table
  note.

Run "docx raw <verb> --help" for each verb's flags and examples.
`;

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const loader = SUBCOMMANDS[verb];
	if (!loader) {
		return fail(
			"USAGE",
			`Unknown raw subcommand: ${verb}`,
			'Run "docx raw --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
