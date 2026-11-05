import { describe, expect, test } from "bun:test";
import { auditCommentMarkers } from "@core/comments/marker-audit";
import { XmlNode } from "@core/parser";

// One function both REPORTS (validate, repair: false) and CURES (--fix,
// repair: true) — so the report's "fixable" flag can't drift from the cure.

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const ref = (id: string) =>
	`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;

function body(paragraphs: string): { name: string; root: XmlNode } {
	const root = XmlNode.parse(
		`<w:document ${W}><w:body>${paragraphs}</w:body></w:document>`,
	).find((node) => !node.isText) as XmlNode;
	return { name: "word/document.xml", root };
}

const xml = (story: { root: XmlNode }) => XmlNode.serialize([story.root]);

function audit(
	story: { name: string; root: XmlNode },
	ids: string[],
	repair: boolean,
	threadRoots?: Map<string, string>,
) {
	return auditCommentMarkers([story], new Set(ids), { repair, threadRoots });
}

describe("auditCommentMarkers", () => {
	test("a balanced comment and a bare-reference point comment are clean", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="0"/>${text("a")}<w:commentRangeEnd w:id="0"/>${ref("0")}${text("b")}${ref("1")}</w:p>`,
		);
		expect(audit(story, ["0", "1"], false)).toEqual([]);
	});

	test("a dangling commentReference is the one ERROR; the cure removes every marker of that id", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="9"/>${text("a")}<w:commentRangeEnd w:id="9"/>${ref("9")}</w:p>`,
		);
		const [finding] = audit(story, [], false);
		expect(finding?.severity).toBe("error");
		expect(finding?.message).toContain("Word refuses to open the file");
		expect(finding?.fix).toBeDefined();
		audit(story, [], true);
		expect(xml(story)).not.toContain('w:id="9"');
		expect(xml(story)).toContain("<w:t>a</w:t>");
		expect(xml(story)).not.toContain("CommentReference");
		expect(audit(story, [], false)).toEqual([]);
	});

	test("a reference run that also carries text keeps its text when the reference goes", () => {
		const story = body(
			`<w:p><w:r><w:t>keep</w:t><w:commentReference w:id="9"/></w:r></w:p>`,
		);
		audit(story, [], true);
		expect(xml(story)).toContain("<w:r><w:t>keep</w:t></w:r>");
	});

	test("a half-open range collapses to a point at the reference mark, as Word does", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="0"/>${text("a")}${ref("0")}${text("b")}</w:p>`,
		);
		const [finding] = audit(story, ["0"], false);
		expect(finding?.severity).toBe("warning");
		expect(finding?.message).toContain("collapses");
		audit(story, ["0"], true);
		expect(xml(story)).toContain(
			`${text("a")}<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>${ref("0")}`,
		);
		expect(audit(story, ["0"], false)).toEqual([]);
	});

	test("a range with no reference gets one after its end", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="0"/>${text("a")}<w:commentRangeEnd w:id="0"/>${text("b")}</w:p>`,
		);
		expect(audit(story, ["0"], false)[0]?.fix).toBeDefined();
		audit(story, ["0"], true);
		expect(xml(story)).toContain(
			`<w:commentRangeEnd w:id="0"/>${ref("0")}${text("b")}`,
		);
		expect(audit(story, ["0"], false)).toEqual([]);
	});

	test("a lone range start with no end or reference is closed into a point comment", () => {
		const story = body(
			`<w:p>${text("a")}<w:commentRangeStart w:id="0"/>${text("b")}</w:p>`,
		);
		audit(story, ["0"], true);
		expect(xml(story)).toContain(
			`<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>${ref("0")}${text("b")}`,
		);
		expect(audit(story, ["0"], false)).toEqual([]);
	});

	test("a block-level range with no reference is reported but not fixable", () => {
		const story = body(
			`<w:commentRangeStart w:id="0"/><w:p>${text("a")}</w:p><w:commentRangeEnd w:id="0"/>`,
		);
		const [finding] = audit(story, ["0"], false);
		expect(finding?.fix).toBeUndefined();
		expect(finding?.hint).toBeDefined();
		const before = xml(story);
		audit(story, ["0"], true);
		expect(xml(story)).toBe(before);
	});

	test("an orphan comment is reported, never deleted or invented an anchor", () => {
		const story = body(`<w:p>${text("a")}</w:p>`);
		const [finding] = audit(story, ["0"], false);
		expect(finding?.severity).toBe("warning");
		expect(finding?.message).toContain("silently deletes it");
		expect(finding?.fix).toBeUndefined();
		audit(story, ["0"], true);
		expect(xml(story)).not.toContain("comment");
	});

	test("an orphan REPLY under an anchored root is re-anchored in its thread", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="0"/>${text("a")}<w:commentRangeEnd w:id="0"/>${ref("0")}</w:p>`,
		);
		const threads = new Map([
			["1", "0"],
			["2", "0"],
		]);
		const findings = audit(story, ["0", "1", "2"], false, threads);
		expect(findings.map((finding) => finding.fix !== undefined)).toEqual([
			true,
			true,
		]);
		audit(story, ["0", "1", "2"], true, threads);
		const out = xml(story);
		// Word's thread layout: starts nest, each reply closes after the previous reference.
		expect(out).toContain(
			`<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/>`,
		);
		expect(out.indexOf('w:commentReference w:id="1"')).toBeLessThan(
			out.indexOf('w:commentReference w:id="2"'),
		);
		expect(audit(story, ["0", "1", "2"], false, threads)).toEqual([]);
	});

	test("a reply whose root only GAINS its reference in the same pass is fixable in one pass", () => {
		const story = body(
			`<w:p><w:commentRangeStart w:id="0"/>${text("a")}<w:commentRangeEnd w:id="0"/></w:p>`,
		);
		const threads = new Map([["1", "0"]]);
		const findings = audit(story, ["0", "1"], false, threads);
		expect(findings.every((finding) => finding.fix !== undefined)).toBe(true);
		audit(story, ["0", "1"], true, threads);
		expect(audit(story, ["0", "1"], false, threads)).toEqual([]);
	});

	test("markers duplicated under mc:AlternateContent count once and are cured in both copies", () => {
		const shape = (inner: string) =>
			`<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:txbxContent>${inner}</w:txbxContent></mc:Choice><mc:Fallback><w:txbxContent>${inner}</w:txbxContent></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
		const story = body(shape(`<w:p>${text("a")}${ref("9")}</w:p>`));
		expect(audit(story, [], false)).toHaveLength(1);
		audit(story, [], true);
		expect(xml(story)).not.toContain('w:id="9"');
	});
});
