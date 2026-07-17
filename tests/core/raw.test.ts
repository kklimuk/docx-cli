import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Document } from "@core/ast/document";
import { XmlNode } from "@core/parser";
import {
	isRelationshipFragment,
	normalizePartName,
	Raw,
	RawError,
} from "@core/raw";
import {
	diffValidationIssues,
	validateWmlXml,
	validationXmlFor,
} from "@core/raw/validate";

let workspace: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-raw-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
});

async function openFixture(name: string, label: string): Promise<Document> {
	const target = join(workspace, label);
	await Bun.write(target, Bun.file(`tests/fixtures/${name}`));
	return Document.open(target);
}

function rejection(document: Document, xml: string): RawError {
	try {
		new Raw(document).prepareFragment(xml, "blocks");
	} catch (error) {
		if (error instanceof RawError) return error;
		throw error;
	}
	throw new Error(`expected RawError for: ${xml}`);
}

describe("gate 1 — well-formedness", () => {
	test("unclosed element rejects with line/col", async () => {
		const document = await openFixture("minimal.docx", "wf-unclosed.docx");
		const error = rejection(document, "<w:p><w:r><w:t>a</w:t>");
		expect(error.code).toBe("INVALID_XML");
		expect(error.message).toContain("not well-formed");
		expect(error.message).toContain("line 1");
	});

	test("stray < in text rejects instead of silently truncating", async () => {
		const document = await openFixture("minimal.docx", "wf-lt.docx");
		const error = rejection(document, "<w:p><w:r><w:t>a < b</w:t></w:r></w:p>");
		expect(error.message).toContain("not well-formed");
	});

	test.each([
		["<!DOCTYPE", "<!DOCTYPE x><w:p><w:r><w:t>a</w:t></w:r></w:p>"],
		["<!--", "<w:p><!-- note --><w:r><w:t>a</w:t></w:r></w:p>"],
		["<![CDATA[", "<w:p><w:r><w:t><![CDATA[a]]></w:t></w:r></w:p>"],
		["<?", '<?xml version="1.0"?><w:p><w:r><w:t>a</w:t></w:r></w:p>'],
	])("%s rejects (parser would drop or rewrite it)", async (needle, xml) => {
		const document = await openFixture(
			"minimal.docx",
			`wf-${needle.length}.docx`,
		);
		const error = rejection(document, xml);
		expect(error.message).toContain(needle);
	});
});

describe("gate 2 — addressable roots", () => {
	test("block-level w:sdt rejects with the invariant named", async () => {
		const document = await openFixture("minimal.docx", "roots-sdt.docx");
		const error = rejection(document, "<w:sdt><w:sdtContent/></w:sdt>");
		expect(error.message).toContain("not addressable");
		expect(error.hint).toContain("<w:p>");
	});

	test("w:sectPr root points at raw replace --at sN", async () => {
		const document = await openFixture("minimal.docx", "roots-sectpr.docx");
		const error = rejection(document, "<w:sectPr/>");
		expect(error.hint).toContain("sN");
	});

	test("sectPr context accepts exactly one w:sectPr and nothing else", async () => {
		const document = await openFixture("minimal.docx", "roots-sect-ok.docx");
		const raw = new Raw(document);
		const prepared = raw.prepareFragment(
			'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
			"sectPr",
		);
		expect(prepared.nodes).toHaveLength(1);
		expect(() =>
			raw.prepareFragment("<w:p><w:r><w:t>a</w:t></w:r></w:p>", "sectPr"),
		).toThrow(RawError);
	});

	test("bare text outside an element rejects", async () => {
		const document = await openFixture("minimal.docx", "roots-text.docx");
		const error = rejection(
			document,
			"hello <w:p><w:r><w:t>a</w:t></w:r></w:p>",
		);
		expect(error.message).toContain("text outside");
	});

	test("multiple p/tbl roots pass", async () => {
		const document = await openFixture("minimal.docx", "roots-multi.docx");
		const prepared = new Raw(document).prepareFragment(
			"<w:p><w:r><w:t>a</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
			"blocks",
		);
		expect(prepared.nodes.map((node) => node.tag)).toEqual(["w:p", "w:tbl"]);
	});
});

describe("gate 2.5 — inter-element whitespace", () => {
	test("pretty-printed fragment loses whitespace #text inside containers, keeps leaf text", async () => {
		const document = await openFixture("minimal.docx", "ws.docx");
		const prepared = new Raw(document).prepareFragment(
			'<w:p>\n\t<w:pPr>\n\t\t<w:jc w:val="center"/>\n\t</w:pPr>\n\t<w:r>\n\t\t<w:t xml:space="preserve">a b</w:t>\n\t</w:r>\n</w:p>',
			"blocks",
		);
		const paragraph = prepared.nodes[0];
		if (!paragraph) throw new Error("expected a root");
		const strayText = (node: XmlNode): boolean => {
			const hasElements = node.children.some((child) => !child.isText);
			if (hasElements && node.children.some((child) => child.isText))
				return true;
			return node.children.some((child) => !child.isText && strayText(child));
		};
		expect(strayText(paragraph)).toBe(false);
		expect(paragraph.findDescendant("w:t")?.collectText()).toBe("a b");
	});
});

describe("gate 3 — element order", () => {
	test("pPr after a run rejects", async () => {
		const document = await openFixture("minimal.docx", "order-ppr.docx");
		const error = rejection(
			document,
			'<w:p><w:r><w:t>x</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>',
		);
		expect(error.message).toContain("<w:pPr> must be the first child");
	});

	test("misordered rPr names the offending pair", async () => {
		const document = await openFixture("minimal.docx", "order-rpr.docx");
		const error = rejection(
			document,
			"<w:p><w:r><w:rPr><w:i/><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>",
		);
		expect(error.message).toContain("<w:b> must precede <w:i>");
	});

	test("misordered pPr children reject", async () => {
		const document = await openFixture("minimal.docx", "order-ppr2.docx");
		const error = rejection(
			document,
			'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
		);
		expect(error.message).toContain("<w:spacing> must precede <w:jc>");
	});

	test("correctly ordered fragment passes", async () => {
		const document = await openFixture("minimal.docx", "order-ok.docx");
		const prepared = new Raw(document).prepareFragment(
			'<w:p><w:pPr><w:spacing w:before="240"/><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>x</w:t></w:r></w:p>',
			"blocks",
		);
		expect(prepared.nodes).toHaveLength(1);
	});
});

describe("gate 4 — namespaces + marker", () => {
	test("known prefix auto-declares on the doc root", async () => {
		const document = await openFixture("minimal.docx", "ns-known.docx");
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		expect(root?.getAttribute("xmlns:v")).toBeUndefined();
		new Raw(document).prepareFragment(
			'<w:p><w:r><w:pict><v:rect style="width:10pt"/></w:pict></w:r></w:p>',
			"blocks",
		);
		expect(root?.getAttribute("xmlns:v")).toBe("urn:schemas-microsoft-com:vml");
	});

	test("unknown prefix rejects with a declare hint", async () => {
		const document = await openFixture("minimal.docx", "ns-unknown.docx");
		const error = rejection(
			document,
			"<w:p><w:r><foo:bar/><w:t>x</w:t></w:r></w:p>",
		);
		expect(error.message).toContain('"foo:"');
		expect(error.hint).toContain("xmlns:foo");
	});

	test("in-fragment declaration satisfies the check without touching the doc root", async () => {
		const document = await openFixture("minimal.docx", "ns-local.docx");
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		new Raw(document).prepareFragment(
			'<w:p><w:r><foo:bar xmlns:foo="urn:example"/><w:t>x</w:t></w:r></w:p>',
			"blocks",
		);
		expect(root?.getAttribute("xmlns:foo")).toBeUndefined();
	});

	test("marker stamps roots and registers dcx as mc:Ignorable", async () => {
		const document = await openFixture("minimal.docx", "ns-marker.docx");
		const prepared = new Raw(document).prepareFragment(
			"<w:p><w:r><w:t>x</w:t></w:r></w:p>",
			"blocks",
		);
		expect(prepared.nodes[0]?.getAttribute("dcx:raw")).toBe("1");
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		expect(root?.getAttribute("xmlns:dcx")).toBe("urn:docx-cli:raw");
		expect(root?.getAttribute("xmlns:mc")).toBeDefined();
		expect(root?.getAttribute("mc:Ignorable")?.split(/\s+/)).toContain("dcx");
	});
});

describe("gate 5 — reference/id integrity", () => {
	test("dangling rId rejects", async () => {
		const document = await openFixture("minimal.docx", "ref-rid.docx");
		const error = rejection(
			document,
			'<w:p><w:hyperlink r:id="rId999"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>',
		);
		expect(error.message).toContain('"rId999"');
		expect(error.hint).toContain("hyperlinks add");
	});

	test("missing footnote id rejects", async () => {
		const document = await openFixture("minimal.docx", "ref-fn.docx");
		const error = rejection(
			document,
			'<w:p><w:r><w:footnoteReference w:id="42"/></w:r></w:p>',
		);
		expect(error.message).toContain('footnote id "42"');
	});

	test("colliding wp:docPr ids re-mint inside the fragment", async () => {
		const document = await openFixture("images.docx", "ref-docpr.docx");
		const usedIds = new Set<string>();
		const collect = (node: XmlNode): void => {
			if (node.tag === "wp:docPr") {
				const id = node.getAttribute("id");
				if (id) usedIds.add(id);
			}
			for (const child of node.children) collect(child);
		};
		for (const root of document.documentTree) collect(root);
		const taken = [...usedIds][0];
		if (!taken) throw new Error("images.docx should carry a wp:docPr");
		const prepared = new Raw(document).prepareFragment(
			`<w:p><w:r><w:drawing><wp:inline><wp:docPr id="${taken}" name="x"/></wp:inline></w:drawing></w:r></w:p>`,
			"blocks",
		);
		const minted = prepared.nodes[0]
			?.findDescendant("wp:docPr")
			?.getAttribute("id");
		expect(minted).toBeDefined();
		expect(minted).not.toBe(taken);
		expect(usedIds.has(minted as string)).toBe(false);
	});

	test("colliding bookmark ids re-map as a linked pair", async () => {
		const document = await openFixture("minimal.docx", "ref-bookmark.docx");
		const body = XmlNode.findRoot(
			document.documentTree,
			"w:document",
		)?.findChild("w:body");
		body?.children.unshift(
			...XmlNode.parse(
				'<w:p><w:bookmarkStart w:id="0" w:name="existing"/><w:bookmarkEnd w:id="0"/></w:p>',
			),
		);
		const prepared = new Raw(document).prepareFragment(
			'<w:p><w:bookmarkStart w:id="0" w:name="mine"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>',
			"blocks",
		);
		const paragraph = prepared.nodes[0];
		const start = paragraph
			?.findDescendant("w:bookmarkStart")
			?.getAttribute("w:id");
		const end = paragraph
			?.findDescendant("w:bookmarkEnd")
			?.getAttribute("w:id");
		expect(start).toBeDefined();
		expect(start).not.toBe("0");
		expect(end).toBe(start);
	});

	test("unknown style and numId warn instead of rejecting", async () => {
		const document = await openFixture("minimal.docx", "ref-warn.docx");
		const prepared = new Raw(document).prepareFragment(
			'<w:p><w:pPr><w:pStyle w:val="NoSuchStyle"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
			"blocks",
		);
		expect(prepared.warnings.some((w) => w.includes("NoSuchStyle"))).toBe(true);
		expect(prepared.warnings.some((w) => w.includes('"77"'))).toBe(true);
	});
});

describe("gate 6 — XSD engine", () => {
	test("a real Word-authored document validates clean after MCE preprocessing", async () => {
		const document = await openFixture("academic-paper.docx", "xsd-clean.docx");
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		if (!root) throw new Error("no document root");
		expect(validateWmlXml(validationXmlFor(root))).toEqual([]);
	});

	test("a bogus element shows up as a fresh issue via the baseline diff", async () => {
		const document = await openFixture("academic-paper.docx", "xsd-bogus.docx");
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		if (!root) throw new Error("no document root");
		const baseline = validateWmlXml(validationXmlFor(root));
		root
			.findChild("w:body")
			?.children.unshift(...XmlNode.parse("<w:bogusElement/>"));
		const fresh = diffValidationIssues(
			baseline,
			validateWmlXml(validationXmlFor(root)),
		);
		expect(fresh).toHaveLength(1);
		expect(fresh[0]?.message).toContain("bogusElement");
	});

	test("our raw marker survives validation because dcx rides mc:Ignorable", async () => {
		const document = await openFixture("minimal.docx", "xsd-marker.docx");
		const raw = new Raw(document);
		const root = XmlNode.findRoot(document.documentTree, "w:document");
		if (!root) throw new Error("no document root");
		const baseline = validateWmlXml(validationXmlFor(root));
		const prepared = raw.prepareFragment(
			"<w:p><w:r><w:t>marked</w:t></w:r></w:p>",
			"blocks",
		);
		root.findChild("w:body")?.children.unshift(...prepared.nodes);
		const fresh = diffValidationIssues(
			baseline,
			validateWmlXml(validationXmlFor(root)),
		);
		expect(fresh).toEqual([]);
	});
});

describe("diffValidationIssues", () => {
	const issue = (message: string, line?: number) => ({ message, line });

	test("pre-existing issues cancel even when their line shifts", () => {
		expect(
			diffValidationIssues([issue("bad thing", 10)], [issue("bad thing", 25)]),
		).toEqual([]);
	});

	test("duplicate counts are respected (multiset, not set)", () => {
		const fresh = diffValidationIssues(
			[issue("bad thing")],
			[issue("bad thing"), issue("bad thing")],
		);
		expect(fresh).toHaveLength(1);
	});

	test("new issues surface", () => {
		expect(
			diffValidationIssues([issue("old")], [issue("old"), issue("new")]),
		).toEqual([issue("new", undefined)]);
	});
});

describe("Raw.serializeBlock", () => {
	test("returns the exact XML of a resolved block", async () => {
		const document = await openFixture("minimal.docx", "get.docx");
		const reference = document.body.resolveBlock("p0");
		const xml = new Raw(document).serializeBlock(reference);
		expect(xml.startsWith("<w:p")).toBe(true);
		expect(xml.endsWith("</w:p>") || xml.endsWith("/>")).toBe(true);
	});
});

describe("relationships — gates", () => {
	const HYPERLINK_TYPE =
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

	function relRejection(
		document: Document,
		xml: string,
		mode: { kind: "insert" } | { kind: "replace"; rId: string } = {
			kind: "insert",
		},
	): RawError {
		try {
			new Raw(document).prepareRelationships(xml, mode);
		} catch (error) {
			if (error instanceof RawError) return error;
			throw error;
		}
		throw new Error(`expected RawError for: ${xml}`);
	}

	test("isRelationshipFragment sniffs the root shape", () => {
		expect(isRelationshipFragment('<Relationship Id="rId9"/>')).toBe(true);
		expect(
			isRelationshipFragment('  <Relationship Type="t" Target="u"/>'),
		).toBe(true);
		expect(isRelationshipFragment("<w:p><w:r/></w:p>")).toBe(false);
		expect(isRelationshipFragment("<Relationships/>")).toBe(false);
	});

	test("mixing block and Relationship roots rejects", async () => {
		const document = await openFixture("minimal.docx", "rel-mix.docx");
		const error = relRejection(
			document,
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://x" TargetMode="External"/><w:p/>`,
		);
		expect(error.message).toContain("only <Relationship> roots");
	});

	test("children, unknown attributes, missing Type/Target, bad TargetMode reject", async () => {
		const document = await openFixture("minimal.docx", "rel-shape.docx");
		expect(
			relRejection(
				document,
				'<Relationship Type="t" Target="u" TargetMode="External"><x/></Relationship>',
			).message,
		).toContain("empty element");
		expect(
			relRejection(
				document,
				'<Relationship Type="t" Target="u" TargetMode="External" Bogus="1"/>',
			).message,
		).toContain('"Bogus"');
		expect(
			relRejection(document, '<Relationship Target="u" TargetMode="External"/>')
				.message,
		).toContain("Type");
		expect(
			relRejection(document, '<Relationship Type="t" TargetMode="External"/>')
				.message,
		).toContain("Target");
		expect(
			relRejection(
				document,
				'<Relationship Type="t" Target="u" TargetMode="Sideways"/>',
			).message,
		).toContain("TargetMode");
	});

	test("an internal Target must resolve to an existing part", async () => {
		const document = await openFixture("minimal.docx", "rel-target.docx");
		const error = relRejection(
			document,
			'<Relationship Type="t" Target="embeddings/missing.bin"/>',
		);
		expect(error.message).toContain("word/embeddings/missing.bin");
		expect(error.hint).toContain("External");
		// An existing part passes.
		const ok = new Raw(document).prepareRelationships(
			'<Relationship Type="t" Target="styles.xml"/>',
			{ kind: "insert" },
		);
		expect(ok.ids).toHaveLength(1);
	});

	test("insert: colliding and duplicate ids reject; omitted ids mint free rIdN", async () => {
		const document = await openFixture("minimal.docx", "rel-ids.docx");
		const existing = document.relationships.list()[0]?.id as string;
		expect(
			relRejection(
				document,
				`<Relationship Id="${existing}" Type="t" Target="https://x" TargetMode="External"/>`,
			).message,
		).toContain("already exists");
		expect(
			relRejection(
				document,
				'<Relationship Id="rId777" Type="t" Target="https://a" TargetMode="External"/><Relationship Id="rId777" Type="t" Target="https://b" TargetMode="External"/>',
			).message,
		).toContain("Duplicate");
		const raw = new Raw(document);
		const prepared = raw.prepareRelationships(
			'<Relationship Type="t" Target="https://a" TargetMode="External"/><Relationship Type="t" Target="https://b" TargetMode="External"/>',
			{ kind: "insert" },
		);
		expect(prepared.ids).toHaveLength(2);
		expect(new Set(prepared.ids).size).toBe(2);
		for (const id of prepared.ids) {
			expect(id).toMatch(/^rId\d+$/);
			expect(document.relationships.findByRid(id)).toBeUndefined();
		}
		raw.applyRelationships(prepared);
		for (const id of prepared.ids) {
			expect(document.relationships.findByRid(id)).toBeDefined();
		}
	});

	test("replace: a changed Id rejects, an omitted Id is stamped", async () => {
		const document = await openFixture("minimal.docx", "rel-replace.docx");
		const raw = new Raw(document);
		const inserted = raw.prepareRelationships(
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://old" TargetMode="External"/>`,
			{ kind: "insert" },
		);
		raw.applyRelationships(inserted);
		const rId = inserted.ids[0] as string;
		expect(
			relRejection(
				document,
				`<Relationship Id="rId999" Type="${HYPERLINK_TYPE}" Target="https://new" TargetMode="External"/>`,
				{ kind: "replace", rId },
			).message,
		).toContain("must keep");
		const replaced = raw.prepareRelationships(
			`<Relationship Type="${HYPERLINK_TYPE}" Target="https://new" TargetMode="External"/>`,
			{ kind: "replace", rId },
		);
		expect(replaced.ids).toEqual([rId]);
		raw.applyRelationships(replaced);
		expect(document.relationships.findByRid(rId)?.getAttribute("Target")).toBe(
			"https://new",
		);
	});

	test("a redundant rels-namespace xmlns is stripped, a foreign one rejects", async () => {
		const document = await openFixture("minimal.docx", "rel-xmlns.docx");
		const prepared = new Raw(document).prepareRelationships(
			'<Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships" Type="t" Target="https://x" TargetMode="External"/>',
			{ kind: "insert" },
		);
		expect(prepared.nodes[0]?.getAttribute("xmlns")).toBeUndefined();
		expect(
			relRejection(
				document,
				'<Relationship xmlns="urn:other" Type="t" Target="https://x" TargetMode="External"/>',
			).message,
		).toContain("xmlns");
	});

	test("serializeRelationships: one, all, and missing", async () => {
		const document = await openFixture("minimal.docx", "rel-serialize.docx");
		const raw = new Raw(document);
		const all = raw.serializeRelationships();
		expect(all).toContain("<Relationships");
		const first = document.relationships.list()[0]?.id as string;
		const one = raw.serializeRelationships(first);
		expect(one).toContain(`Id="${first}"`);
		expect(one?.startsWith("<Relationship ")).toBe(true);
		expect(raw.serializeRelationships("rId999")).toBeUndefined();
	});
});

describe("parts — gates", () => {
	function partRejection(fn: () => unknown): RawError {
		try {
			fn();
		} catch (error) {
			if (error instanceof RawError) return error;
			throw error;
		}
		throw new Error("expected RawError");
	}

	test("normalizePartName strips a leading slash and rejects traversal", () => {
		expect(normalizePartName("/word/media/a.png")).toBe("word/media/a.png");
		expect(
			partRejection(() => normalizePartName("word/../evil.xml")).code,
		).toBe("USAGE");
		expect(partRejection(() => normalizePartName("")).code).toBe("USAGE");
	});

	test("add: existing part, .rels names, and source/name mismatches reject", async () => {
		const document = await openFixture("minimal.docx", "part-add-guards.docx");
		const raw = new Raw(document);
		expect(
			partRejection(() =>
				raw.preparePartAdd(
					"word/styles.xml",
					{ xml: "<w:styles/>" },
					undefined,
				),
			).message,
		).toContain("already exists");
		expect(
			partRejection(() =>
				raw.preparePartAdd("word/_rels/x.xml.rels", { xml: "<a/>" }, undefined),
			).message,
		).toContain("rels");
		expect(
			partRejection(() =>
				raw.preparePartAdd("word/embeddings/a.bin", { xml: "<a/>" }, undefined),
			).message,
		).toContain("--from");
		expect(
			partRejection(() =>
				raw.preparePartAdd(
					"customXml/a.xml",
					{ bytes: new Uint8Array([1]) },
					undefined,
				),
			).message,
		).toContain("--xml");
	});

	test("add: unknown extension requires --content-type; the xml Default covers .xml", async () => {
		const document = await openFixture("minimal.docx", "part-add-ct.docx");
		const raw = new Raw(document);
		const error = partRejection(() =>
			raw.preparePartAdd(
				"word/embeddings/a.bin",
				{ bytes: new Uint8Array([1]) },
				undefined,
			),
		);
		expect(error.message).toContain("--content-type");
		expect(error.hint).toContain("oleObject");
		const viaDefault = raw.preparePartAdd(
			"customXml/item9.xml",
			{ xml: '<data xmlns="urn:t"/>' },
			undefined,
		);
		expect(viaDefault.contentType).toBeUndefined();
		const explicit = raw.preparePartAdd(
			"word/embeddings/a.bin",
			{ bytes: new Uint8Array([1]) },
			"application/vnd.openxmlformats-officedocument.oleObject",
		);
		expect(explicit.contentType).toBe(
			"application/vnd.openxmlformats-officedocument.oleObject",
		);
	});

	test("add: a leading XML declaration is tolerated and stored verbatim; multi-root rejects", async () => {
		const document = await openFixture("minimal.docx", "part-add-decl.docx");
		const raw = new Raw(document);
		const xml =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><data xmlns="urn:t"/>';
		const prepared = raw.preparePartAdd(
			"customXml/decl.xml",
			{ xml },
			undefined,
		);
		expect(prepared.content).toBe(xml);
		expect(prepared.root?.tag).toBe("data");
		expect(
			partRejection(() =>
				raw.preparePartAdd(
					"customXml/multi.xml",
					{ xml: "<a/><b/>" },
					undefined,
				),
			).message,
		).toContain("one root");
	});

	test("replace: body/notes/comments reject, view-backed parts route, root must match", async () => {
		const document = await openFixture("minimal.docx", "part-replace.docx");
		const raw = new Raw(document);
		await expect(
			raw.preparePartReplace("word/document.xml", "<w:document/>"),
		).rejects.toMatchObject({
			hint: expect.stringContaining("block locators"),
		});
		await expect(
			raw.preparePartReplace("word/footnotes.xml", "<w:footnotes/>"),
		).rejects.toMatchObject({
			hint: expect.stringContaining("pair with body reference marks"),
		});
		await expect(
			raw.preparePartReplace("word/comments.xml", "<w:comments/>"),
		).rejects.toMatchObject({
			hint: expect.stringContaining("docx comments"),
		});
		// styles.xml is view-backed now: a same-root replacement prepares fine.
		const stylesXml = await document.pkg.readText("word/styles.xml");
		const prepared = await raw.preparePartReplace("word/styles.xml", stylesXml);
		expect(prepared.route.kind).toBe("view");
		raw.applyPartAdd(
			raw.preparePartAdd(
				"word/embeddings/a.bin",
				{ bytes: new Uint8Array([1]) },
				"application/vnd.openxmlformats-officedocument.oleObject",
			),
		);
		await expect(
			raw.preparePartReplace("word/embeddings/a.bin", "<a/>"),
		).rejects.toMatchObject({ message: expect.stringContaining("binary") });
		await expect(
			raw.preparePartReplace("word/fontTable.xml", '<data xmlns="urn:t"/>'),
		).rejects.toMatchObject({
			message: expect.stringContaining("must keep the part's root <w:fonts>"),
		});
	});
});
