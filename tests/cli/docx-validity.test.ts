import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import type { XmlNode } from "@core/parser";
import { runCli, tempWorkspace } from "./harness";

// Catches the bug class that produced Word's "unreadable content" dialog: a part
// that uses an XML namespace prefix (`r:id` on a note-body `<w:hyperlink>`) the
// part never declares — malformed XML — and the related "dangling rId" class.
// Our LibreOffice-render + AST-roundtrip tests missed both because LibreOffice
// is permissive and the AST never re-parses namespaces. These checks are the
// guard. (`xmllint` is run too when present — it independently flags undeclared
// prefixes — but the structured checks below always run in CI.)

const XMLLINT = Bun.which("xmllint");
const PNG_PATH = join(
	import.meta.dir,
	"..",
	"fixtures",
	"assets",
	"sample.png",
);

describe("generated .docx is valid OOXML", () => {
	test("markdown with a reused footnote containing a hyperlink (the regression)", async () => {
		const docPath = join(tempWorkspace("validity-fn-link"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"First cite[^a] then again[^a].\n\n[^a]: see [AP](https://ap.example.com/x) and [NBC](https://nbc.example.com/y).",
		);
		await assertValidDocx(docPath);
	});

	test("body hyperlink + heading + footnote-with-link in one doc", async () => {
		const docPath = join(tempWorkspace("validity-mixed"), "out.docx");
		await runCli("create", docPath, "--text", "seed");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"# Title\n\nA [body link](https://example.com/a) and a note[^n].\n\n[^n]: with [a link](https://example.com/b).",
		);
		await assertValidDocx(docPath);
	});

	test("hyperlinks add (body relationship)", async () => {
		const docPath = join(tempWorkspace("validity-hyperlink"), "out.docx");
		await runCli("create", docPath, "--text", "Click this phrase here.");
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p0:6-17",
			"--url",
			"https://example.com",
		);
		await assertValidDocx(docPath);
	});

	test("footnotes add (no links — the plain note part stays valid)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "a note");
		await assertValidDocx(docPath);
	});

	// The note-body markdown path (`--markdown` on add/edit) mints hyperlink rels
	// and must route them into the NOTE part's own rels — the same class of bug
	// the regression above guards, but reached through a different command.
	test("footnotes add --markdown with a body hyperlink (note-part rels)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add-md"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--markdown",
			"See [AP](https://ap.example.com/x) for the long form.",
		);
		await assertValidDocx(docPath);
	});

	test("endnotes edit --markdown with a body hyperlink", async () => {
		const docPath = join(tempWorkspace("validity-en-edit-md"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"placeholder",
		);
		await runCli(
			"endnotes",
			"edit",
			docPath,
			"--at",
			"en1",
			"--markdown",
			"Revised with [a source](https://example.com/src).",
		);
		await assertValidDocx(docPath);
	});

	// A markdown image in a note body is dropped (note bodies are text + links);
	// if the strip ever regresses the image embeds with its media rel in
	// document.xml.rels while the <w:drawing r:embed> lands in footnotes.xml — a
	// dangling rId the relationship check below catches.
	test("footnotes add --markdown with an image strips it (no dangling media rel)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add-img"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--markdown",
			`Caption text ![pic](${PNG_PATH}) and more.`,
		);
		await assertValidDocx(docPath);
	});
});

/** Assert every XML part of `docPath` is namespace-well-formed and every
 *  relationship reference resolves. */
async function assertValidDocx(docPath: string): Promise<void> {
	const pkg = await Pkg.open(docPath);
	const xmlParts = pkg
		.listParts()
		.filter((name) => name.endsWith(".xml") || name.endsWith(".rels"));

	for (const name of xmlParts) {
		const text = await pkg.readText(name);

		// (1) Every namespace prefix used on a tag or attribute key is declared.
		// This is the exact invariant the footnote-hyperlink bug violated (`r:id`
		// used, `xmlns:r` never declared on `footnotes.xml`).
		const tree = await pkg.readPart(name);
		assertNamespacesDeclared(name, tree ?? []);

		// (2) Every `r:id` / `r:embed` / `r:link` resolves to a relationship in
		// THIS part's own rels (a dangling rId is the other corruption class).
		await assertRelationshipsResolve(pkg, name, tree ?? []);

		// (3) Belt-and-suspenders: libxml2's namespace-aware parser, when present,
		// independently rejects undeclared prefixes + any other malformedness.
		if (XMLLINT) await assertXmllintClean(name, text);
	}
}

function assertNamespacesDeclared(partName: string, tree: XmlNode[]): void {
	const declared = new Set<string>(["xml"]);
	const used = new Set<string>();
	const visit = (node: XmlNode): void => {
		const tagColon = node.tag.indexOf(":");
		if (tagColon > 0) used.add(node.tag.slice(0, tagColon));
		for (const key of Object.keys(node.attributes)) {
			if (key === "xmlns") continue;
			if (key.startsWith("xmlns:")) {
				declared.add(key.slice("xmlns:".length));
				continue;
			}
			const colon = key.indexOf(":");
			if (colon > 0) used.add(key.slice(0, colon));
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);

	for (const prefix of used) {
		expect(
			declared.has(prefix),
			`${partName}: namespace prefix "${prefix}:" is used but never declared (missing xmlns:${prefix}) — malformed XML, which Word reports as "unreadable content"`,
		).toBe(true);
	}
}

async function assertRelationshipsResolve(
	pkg: Pkg,
	partName: string,
	tree: XmlNode[],
): Promise<void> {
	const used = new Set<string>();
	const visit = (node: XmlNode): void => {
		for (const [key, value] of Object.entries(node.attributes)) {
			if (
				(key === "r:id" || key === "r:embed" || key === "r:link") &&
				typeof value === "string"
			) {
				used.add(value);
			}
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);
	if (used.size === 0) return;

	const relsName = relsPartNameFor(partName);
	const relsTree = await pkg.readPart(relsName);
	const declared = new Set<string>();
	for (const root of relsTree ?? []) {
		const visitRel = (node: XmlNode): void => {
			if (node.tag === "Relationship") {
				const id = node.attributes.Id;
				if (typeof id === "string") declared.add(id);
			}
			for (const child of node.children) visitRel(child);
		};
		visitRel(root);
	}

	for (const rId of used) {
		expect(
			declared.has(rId),
			`${partName}: references relationship "${rId}" but ${relsName} has no such <Relationship> (dangling rId — Word reports "unreadable content")`,
		).toBe(true);
	}
}

async function assertXmllintClean(
	partName: string,
	xml: string,
): Promise<void> {
	const proc = Bun.spawn(["xmllint", "--noout", "-"], {
		stdin: "pipe",
		stdout: "ignore",
		stderr: "pipe",
	});
	proc.stdin.write(xml);
	await proc.stdin.end();
	const [code, stderr] = await Promise.all([
		proc.exited,
		new Response(proc.stderr).text(),
	]);
	expect(code, `${partName} is not namespace-well-formed:\n${stderr}`).toBe(0);
}

function relsPartNameFor(partName: string): string {
	const slash = partName.lastIndexOf("/");
	const dir = partName.slice(0, slash);
	const base = partName.slice(slash + 1);
	return `${dir}/_rels/${base}.rels`;
}
