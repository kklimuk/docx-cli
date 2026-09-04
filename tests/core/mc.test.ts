import { describe, expect, test } from "bun:test";
import {
	alternateContentBranch,
	collectTextBoxContents,
	syncTextBoxFallbacks,
} from "@core/mc";
import { XmlNode } from "@core/parser";

const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="x" xmlns:v="y"';

function pair(choiceStory: string, fallbackStory: string): string {
	return (
		`<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp><wps:txbx><w:txbxContent>${choiceStory}</w:txbxContent></wps:txbx></wps:wsp></w:drawing></mc:Choice>` +
		`<mc:Fallback><w:pict><v:rect><v:textbox><w:txbxContent>${fallbackStory}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent>`
	);
}
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("alternateContentBranch", () => {
	test("takes the FIRST Choice, else the Fallback", () => {
		const [node] = XmlNode.parse(
			`<mc:AlternateContent ${NS}><mc:Choice Requires="a"><w:t>one</w:t></mc:Choice><mc:Choice Requires="b"><w:t>two</w:t></mc:Choice><mc:Fallback><w:t>fb</w:t></mc:Fallback></mc:AlternateContent>`,
		);
		if (!node) throw new Error("parse");
		expect(alternateContentBranch(node)?.collectText()).toBe("one");
		const [fallbackOnly] = XmlNode.parse(
			`<mc:AlternateContent ${NS}><mc:Fallback><w:t>fb</w:t></mc:Fallback></mc:AlternateContent>`,
		);
		if (!fallbackOnly) throw new Error("parse");
		expect(alternateContentBranch(fallbackOnly)?.collectText()).toBe("fb");
	});
});

describe("collectTextBoxContents", () => {
	test("returns the Choice story only, even when handed the wrapper itself", () => {
		const [node] = XmlNode.parse(
			`<w:r ${NS}>${pair(p("choice"), p("fallback"))}</w:r>`,
		);
		if (!node) throw new Error("parse");
		const stories = collectTextBoxContents(node);
		expect(stories.map((story) => story.collectText())).toEqual(["choice"]);
		const wrapper = node.children[0];
		if (!wrapper) throw new Error("wrapper");
		expect(collectTextBoxContents(wrapper).map((s) => s.collectText())).toEqual(
			["choice"],
		);
	});
});

describe("syncTextBoxFallbacks", () => {
	test("mirrors the Choice story onto the Fallback twin", () => {
		const tree = XmlNode.parse(
			`<w:p ${NS}><w:r>${pair(p("NEW"), p("OLD"))}</w:r></w:p>`,
		);
		expect(syncTextBoxFallbacks(tree)).toBe(1);
		const [root] = tree;
		const fallback = root?.findDescendant("mc:Fallback");
		expect(fallback?.collectText()).toBe("NEW");
	});

	test("nested boxes sync inner-first, so the outer Fallback carries the inner box's new text", () => {
		const inner = `<w:p><w:r>${pair(p("inner-NEW"), p("inner-OLD"))}</w:r></w:p>`;
		const tree = XmlNode.parse(
			`<w:p ${NS}><w:r>${pair(inner, p("outer-OLD"))}</w:r></w:p>`,
		);
		expect(syncTextBoxFallbacks(tree)).toBe(2);
		const [root] = tree;
		const outerFallback = root
			?.findChild("w:r")
			?.findChild("mc:AlternateContent")
			?.findChild("mc:Fallback");
		const text = outerFallback ? XmlNode.serialize([outerFallback]) : "";
		expect(text).toContain("inner-NEW");
		expect(text).not.toContain("inner-OLD");
		expect(text).not.toContain("outer-OLD");
	});

	test("an unpaired wrapper (story counts differ) is left alone", () => {
		const tree = XmlNode.parse(
			`<w:p ${NS}><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>${p("a")}</w:txbxContent></wps:txbx><wps:txbx><w:txbxContent>${p("b")}</w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent>${p("old")}</w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`,
		);
		expect(syncTextBoxFallbacks(tree)).toBe(0);
		expect(tree[0]?.findDescendant("mc:Fallback")?.collectText()).toBe("old");
	});
});
