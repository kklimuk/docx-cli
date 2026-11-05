import { describe, expect, test } from "bun:test";
import {
	auditThreadLinks,
	reconcileSidecars,
	type SidecarRoots,
	threadRootsOf,
} from "@core/ast/document/comments/reconcile-sidecars";
import { XmlNode } from "@core/parser";

// The rules mirror what live Word for Mac (16.114) does when it opens and
// re-saves a package: mint what's missing, prune what's stale, de-dupe
// durableIds, and never touch an existing entry's thread link or resolved bit.

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"`;

function comment(id: string, paraId?: string): string {
	const attribute = paraId ? ` w14:paraId="${paraId}"` : "";
	return `<w:comment w:id="${id}" w:author="A" w:date="2026-03-04T05:06:07Z"><w:p${attribute}><w:r><w:t>c${id}</w:t></w:r></w:p></w:comment>`;
}

function roots(parts: {
	comments: string;
	extended?: string;
	ids?: string;
	extensible?: string;
}): SidecarRoots {
	const root = (xml: string) =>
		XmlNode.parse(xml).find((node) => !node.isText) as XmlNode;
	return {
		comments: root(`<w:comments ${NS}>${parts.comments}</w:comments>`),
		...(parts.extended !== undefined
			? {
					extended: root(
						`<w15:commentsEx ${NS}>${parts.extended}</w15:commentsEx>`,
					),
				}
			: {}),
		...(parts.ids !== undefined
			? {
					ids: root(
						`<w16cid:commentsIds ${NS}>${parts.ids}</w16cid:commentsIds>`,
					),
				}
			: {}),
		...(parts.extensible !== undefined
			? {
					extensible: root(
						`<w16cex:commentsExtensible ${NS}>${parts.extensible}</w16cex:commentsExtensible>`,
					),
				}
			: {}),
	};
}

const serialize = (node: XmlNode | undefined) =>
	node ? XmlNode.serialize([node]) : "";

describe("reconcileSidecars", () => {
	test("a consistent set of parts is left alone and reports nothing", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="1A2B3C4D" w15:done="1"/>`,
			ids: `<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="5E6F7A8B"/>`,
			extensible: `<w16cex:commentExtensible w16cex:durableId="5E6F7A8B" w16cex:dateUtc="2026-03-04T05:06:07Z"/>`,
		});
		const before = [
			parts.comments,
			parts.extended,
			parts.ids,
			parts.extensible,
		].map(serialize);
		expect(reconcileSidecars(parts)).toEqual([]);
		expect(
			[parts.comments, parts.extended, parts.ids, parts.extensible].map(
				serialize,
			),
		).toEqual(before);
	});

	test("absent sidecars stay absent; a missing paraId is still minted", () => {
		const parts = roots({ comments: comment("0") });
		const findings = reconcileSidecars(parts);
		expect(parts.extended).toBeUndefined();
		expect(parts.ids).toBeUndefined();
		expect(serialize(parts.comments)).toMatch(/<w:p w14:paraId="[0-9A-F]{8}">/);
		expect(findings).toHaveLength(1);
		expect(findings[0]?.severity).toBe("silent");
	});

	test("adds a missing commentsExtended entry as open + top-level and reports it as data loss", () => {
		const parts = roots({ comments: comment("0", "1A2B3C4D"), extended: "" });
		const [finding] = reconcileSidecars(parts);
		expect(serialize(parts.extended)).toContain(
			'<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/>',
		);
		expect(finding?.severity).toBe("warning");
		expect(finding?.message).toContain("loses its thread");
		expect(finding?.message).toContain("reopens");
		expect(finding?.fix).toBeDefined();
	});

	test("a new commentsIds entry takes the comment's paraId as durableId (Word's choice)", () => {
		const parts = roots({ comments: comment("0", "1A2B3C4D"), ids: "" });
		reconcileSidecars(parts);
		expect(serialize(parts.ids)).toContain(
			'w16cid:paraId="1A2B3C4D" w16cid:durableId="1A2B3C4D"',
		);
	});

	test("a paraId outside the durableId range, or already taken, gets a freshly minted durableId", () => {
		const parts = roots({
			comments: comment("0", "7FFFFFFF") + comment("1", "1A2B3C4D"),
			ids: `<w16cid:commentId w16cid:paraId="7FFFFFFF" w16cid:durableId="1A2B3C4D"/>`,
		});
		reconcileSidecars(parts);
		const durableIds = [
			...serialize(parts.ids).matchAll(/w16cid:durableId="([0-9A-F]+)"/g),
		].map((match) => match[1]);
		expect(durableIds).toHaveLength(2);
		expect(new Set(durableIds).size).toBe(2);
		for (const durableId of durableIds) {
			expect(Number.parseInt(durableId ?? "", 16)).toBeLessThan(0x7fffffff);
		}
	});

	test("prunes stale entries in every present sidecar — silently", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/><w15:commentEx w15:paraId="0BADF00D" w15:done="1"/>`,
			ids: `<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="0BADF00D" w16cid:durableId="22222222"/>`,
			extensible: `<w16cex:commentExtensible w16cex:durableId="11111111"/><w16cex:commentExtensible w16cex:durableId="22222222"/>`,
		});
		const findings = reconcileSidecars(parts);
		expect(serialize(parts.extended)).not.toContain("0BADF00D");
		expect(serialize(parts.ids)).not.toContain("0BADF00D");
		expect(serialize(parts.extensible)).not.toContain("22222222");
		expect(findings).toHaveLength(3);
		expect(findings.every((finding) => finding.severity === "silent")).toBe(
			true,
		);
	});

	test("de-duplicates durableIds (keeping the first) and gives the re-minted one an extensible entry", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D") + comment("1", "2A2B3C4D"),
			ids: `<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="2A2B3C4D" w16cid:durableId="11111111"/>`,
			extensible: `<w16cex:commentExtensible w16cex:durableId="11111111" w16cex:dateUtc="2026-01-01T00:00:00Z"/>`,
		});
		const findings = reconcileSidecars(parts);
		const ids = serialize(parts.ids);
		expect(ids).toContain(
			'w16cid:paraId="1A2B3C4D" w16cid:durableId="11111111"',
		);
		const reminted = ids.match(
			/w16cid:paraId="2A2B3C4D" w16cid:durableId="([0-9A-F]+)"/,
		)?.[1];
		expect(reminted).toBeDefined();
		expect(reminted).not.toBe("11111111");
		expect(serialize(parts.extensible)).toContain(
			`w16cex:durableId="${reminted}" w16cex:dateUtc="2026-03-04T05:06:07Z"`,
		);
		expect(findings.every((finding) => finding.severity === "silent")).toBe(
			true,
		);
	});

	test("never drops or alters an existing paraIdParent or done — even a dangling one", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D") + comment("1", "2A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="2A2B3C4D" w15:paraIdParent="0BADF00D" w15:done="1"/>`,
		});
		reconcileSidecars(parts);
		expect(serialize(parts.extended)).toContain(
			'<w15:commentEx w15:paraId="2A2B3C4D" w15:paraIdParent="0BADF00D" w15:done="1"/>',
		);
	});

	test("matches keys case-insensitively, so a lowercase key isn't pruned and re-added", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="1a2b3c4d" w15:done="1"/>`,
		});
		expect(reconcileSidecars(parts)).toEqual([]);
		expect(serialize(parts.extended)).toContain('w15:done="1"');
	});

	test("keys a multi-paragraph comment by its LAST paragraph", () => {
		const parts = roots({
			comments: `<w:comment w:id="0" w:author="A"><w:p w14:paraId="0000027A"/><w:p w14:paraId="00000283"/></w:comment>`,
			extended: `<w15:commentEx w15:paraId="0000027A" w15:done="0"/>`,
		});
		reconcileSidecars(parts);
		const extended = serialize(parts.extended);
		expect(extended).toContain('w15:paraId="00000283"');
		expect(extended).not.toContain("0000027A");
	});

	test("re-keys an out-of-range paraId and carries its entries and reply links along", () => {
		const parts = roots({
			comments: comment("0", "E33D2225") + comment("1", "1A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="E33D2225" w15:done="1"/><w15:commentEx w15:paraId="1A2B3C4D" w15:paraIdParent="e33d2225" w15:done="0"/>`,
			ids: `<w16cid:commentId w16cid:paraId="E33D2225" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="22222222"/>`,
		});
		const findings = reconcileSidecars(parts);
		const fresh = serialize(parts.comments).match(
			/<w:comment w:id="0"[^>]*><w:p w14:paraId="([0-9A-F]{8})">/,
		)?.[1];
		expect(fresh).toBeDefined();
		expect(Number.parseInt(fresh ?? "", 16)).toBeLessThan(0x80000000);
		expect(serialize(parts.extended)).toBe(
			serialize(
				roots({
					comments: "",
					extended: `<w15:commentEx w15:paraId="${fresh}" w15:done="1"/><w15:commentEx w15:paraId="1A2B3C4D" w15:paraIdParent="${fresh}" w15:done="0"/>`,
				}).extended,
			),
		);
		expect(serialize(parts.ids)).toContain(
			`w16cid:paraId="${fresh}" w16cid:durableId="11111111"`,
		);
		expect(findings).toHaveLength(1);
		expect(findings[0]?.severity).toBe("warning");
		expect(findings[0]?.message).toContain("out of range");
		expect(findings[0]?.fix).toContain(`→ ${fresh}`);
	});

	test("an out-of-range paraId is re-keyed silently when no commentsExtended holds thread data", () => {
		const parts = roots({
			comments: `<w:comment w:id="0" w:author="A"><w:p w14:paraId="FFFFFFFF"/><w:p w14:paraId="8000ABCD"/></w:comment>`,
		});
		const findings = reconcileSidecars(parts);
		const paraIds = [
			...serialize(parts.comments).matchAll(/w14:paraId="([0-9A-F]{8})"/g),
		].map((match) => Number.parseInt(match[1] ?? "", 16));
		expect(paraIds).toHaveLength(2);
		for (const paraId of paraIds) expect(paraId).toBeLessThan(0x80000000);
		expect(findings.map((finding) => finding.severity)).toEqual(["silent"]);
	});

	test("is idempotent", () => {
		const parts = roots({
			comments: comment("0") + comment("1", "1A2B3C4D"),
			extended: "",
			ids: "",
			extensible: "",
		});
		expect(reconcileSidecars(parts).length).toBeGreaterThan(0);
		expect(reconcileSidecars(parts)).toEqual([]);
	});
});

describe("auditThreadLinks / threadRootsOf", () => {
	test("a paraIdParent naming no comment is a report-only warning", () => {
		const parts = roots({
			comments: comment("0", "1A2B3C4D") + comment("1", "2A2B3C4D"),
			extended: `<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/><w15:commentEx w15:paraId="2A2B3C4D" w15:paraIdParent="0BADF00D" w15:done="0"/>`,
		});
		const [finding] = auditThreadLinks(parts);
		expect(finding?.severity).toBe("warning");
		expect(finding?.message).toContain("c1's thread parent");
		expect(finding?.fix).toBeUndefined();
	});

	test("threadRootsOf walks reply-of-reply chains to the root and survives a cycle", () => {
		const parts = roots({
			comments:
				comment("0", "00000001") +
				comment("1", "00000002") +
				comment("2", "00000003") +
				comment("3", "00000004") +
				comment("4", "00000005"),
			extended: `<w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/><w15:commentEx w15:paraId="00000003" w15:paraIdParent="00000002"/><w15:commentEx w15:paraId="00000004" w15:paraIdParent="00000005"/><w15:commentEx w15:paraId="00000005" w15:paraIdParent="00000004"/>`,
		});
		const threadRoots = threadRootsOf(parts);
		expect(threadRoots.get("1")).toBe("0");
		expect(threadRoots.get("2")).toBe("0");
		expect(threadRoots.has("0")).toBe(false);
		// A 2-cycle terminates (the walk stops at the first repeat).
		expect(threadRoots.has("3") || threadRoots.has("4")).toBe(true);
	});
});
