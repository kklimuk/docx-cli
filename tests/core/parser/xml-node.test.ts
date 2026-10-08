import { describe, expect, test } from "bun:test";
import { XmlNode } from "@core/parser";

describe("XmlNode", () => {
	test("parses element with attributes and text", () => {
		const tree = XmlNode.parse('<root attr="v"><child>hello</child></root>');
		const root = XmlNode.findRoot(tree, "root");
		expect(root).toBeDefined();
		expect(root?.getAttribute("attr")).toBe("v");
		const child = root?.findChild("child");
		expect(child?.collectText()).toBe("hello");
	});

	test("preserves namespaced tags and attributes", () => {
		const xml =
			'<w:r xmlns:w="http://w"><w:t xml:space="preserve">Hi </w:t></w:r>';
		const tree = XmlNode.parse(xml);
		const run = XmlNode.findRoot(tree, "w:r");
		expect(run?.getAttribute("xmlns:w")).toBe("http://w");
		const text = run?.findChild("w:t");
		expect(text?.getAttribute("xml:space")).toBe("preserve");
		expect(text?.collectText()).toBe("Hi ");
	});

	test("round-trips text content with entities", () => {
		const xml =
			"<root>Hostile: &lt;script&gt; &amp; &apos;quotes&apos; &quot;here&quot;</root>";
		const tree = XmlNode.parse(xml);
		const root = XmlNode.findRoot(tree, "root");
		expect(root?.collectText()).toBe("Hostile: <script> & 'quotes' \"here\"");
	});

	test("decodes numeric character references in text and attributes (#12)", () => {
		const xml =
			'<w:lvl><w:lvlText w:val="&#8226;"/><w:t>&#8226; &#x2013; &#x1F600; done</w:t></w:lvl>';
		const root = XmlNode.findRoot(XmlNode.parse(xml), "w:lvl");
		expect(root?.findChild("w:lvlText")?.getAttribute("w:val")).toBe("•");
		expect(root?.findChild("w:t")?.collectText()).toBe("• – 😀 done");
		const serialized = XmlNode.serialize([root as XmlNode]);
		expect(serialized).not.toContain("&amp;#");
		expect(serialized).toContain('w:val="•"');
	});

	test("an escaped reference stays literal text, decoded exactly once", () => {
		const xml = '<w:t a="&amp;#8226;">&amp;#8226; &amp;nbsp;</w:t>';
		const node = XmlNode.findRoot(XmlNode.parse(xml), "w:t");
		expect(node?.getAttribute("a")).toBe("&#8226;");
		expect(node?.collectText()).toBe("&#8226; &nbsp;");
		const serialized = XmlNode.serialize([node as XmlNode]);
		expect(serialized).toBe(
			'<w:t a="&amp;#8226;">&amp;#8226; &amp;nbsp;</w:t>',
		);
	});

	test("a decoded whitespace reference in an attribute is re-referenced on serialize", () => {
		const xml = '<wp:docPr descr="line one&#xA;line&#9;two&#13;"/>';
		const node = XmlNode.findRoot(XmlNode.parse(xml), "wp:docPr");
		expect(node?.getAttribute("descr")).toBe("line one\nline\ttwo\r");
		expect(XmlNode.serialize([node as XmlNode])).toBe(
			'<wp:docPr descr="line one&#10;line&#9;two&#13;"/>',
		);
	});

	test("only tab/newline/CR become references, and only in attributes", () => {
		const node = XmlNode.element("w:t", { a: "x\u000065;" }, [
			XmlNode.textNode("a\tb\nc\u000065;"),
		]);
		expect(XmlNode.serialize([node])).toBe(
			'<w:t a="x\u000065;">a\tb\nc\u000065;</w:t>',
		);
	});

	test("escapes special chars on serialize", () => {
		const node = XmlNode.element("w:t", {}, [XmlNode.textNode('< & > "')]);
		const xml = XmlNode.serialize([node]);
		expect(xml).toContain("&lt;");
		expect(xml).toContain("&amp;");
		expect(xml).toContain("&gt;");
		expect(xml).toContain("&quot;");
	});

	test("findChild returns first match only", () => {
		const tree = XmlNode.parse("<root><a>1</a><a>2</a><b>x</b></root>");
		const root = XmlNode.findRoot(tree, "root");
		expect(root?.findChild("a")?.collectText()).toBe("1");
		expect(root?.findChildren("a").map((node) => node.collectText())).toEqual([
			"1",
			"2",
		]);
	});

	test("findDescendant traverses nested children", () => {
		const tree = XmlNode.parse(
			"<root><a><b><target>deep</target></b></a></root>",
		);
		const root = XmlNode.findRoot(tree, "root");
		expect(root?.findDescendant("target")?.collectText()).toBe("deep");
	});

	test("setAttribute mutates and serializes back", () => {
		const tree = XmlNode.parse('<root attr="old"/>');
		const root = XmlNode.findRoot(tree, "root");
		root?.setAttribute("attr", "new");
		root?.setAttribute("added", "yes");
		const xml = XmlNode.serialize(tree);
		expect(xml).toContain('attr="new"');
		expect(xml).toContain('added="yes"');
	});

	test("serializeWithDeclaration prepends and dedupes", () => {
		const tree = XmlNode.parse('<?xml version="1.0" encoding="UTF-8"?><root/>');
		const xml = XmlNode.serializeWithDeclaration(tree);
		const declarations = xml.match(/<\?xml/g) ?? [];
		expect(declarations).toHaveLength(1);
		expect(xml).toContain("<root");
	});
});
