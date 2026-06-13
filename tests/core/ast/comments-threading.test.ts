import { describe, expect, test } from "bun:test";
import { CommentsView } from "@core/ast/document/comments";

// Word keys <w15:commentEx> (and a reply's w15:paraIdParent) off a comment's
// LAST <w:p> — verified against the Word-authored
// tests/fixtures/comments-with-replies.docx, where the multi-paragraph
// comment registers in the sidecar by its final paragraph's paraId.

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;

const MULTI_PARAGRAPH_COMMENTS = `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="18" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="0000027A"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="00000283"><w:r><w:t>last</w:t></w:r></w:p></w:comment></w:comments>`;

describe("comment threading identity keys off the last paragraph", () => {
	test("paraIdFor returns the last paragraph's paraId", () => {
		const view = CommentsView.fromXml(MULTI_PARAGRAPH_COMMENTS);
		expect(view?.paraIdFor("18")).toBe("00000283");
	});

	test("ensureParaId reuses the last paragraph's existing paraId", () => {
		const view = CommentsView.fromXml(MULTI_PARAGRAPH_COMMENTS);
		expect(view?.ensureParaId("c18")).toBe("00000283");
	});

	test("a reply linked to the last paragraph resolves its parentId", () => {
		const comments = `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="18" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="0000027A"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="00000283"><w:r><w:t>last</w:t></w:r></w:p></w:comment><w:comment w:id="19" w:author="B" w:date="2020-01-02T00:00:00Z"><w:p w14:paraId="80CBDBBD"><w:r><w:t>OK</w:t></w:r></w:p></w:comment></w:comments>`;
		const extended = `<?xml version="1.0"?><w15:commentsEx ${NS}><w15:commentEx w15:paraId="00000283" w15:done="0"/><w15:commentEx w15:paraId="80CBDBBD" w15:paraIdParent="00000283" w15:done="0"/></w15:commentsEx>`;
		const view = CommentsView.fromXml(comments, extended);
		const reply = view
			?.toComments(new Map())
			.find((comment) => comment.id === "c19");
		expect(reply?.parentId).toBe("c18");
	});

	test("threadRootId resolves a chained reply to the thread root", () => {
		const comments = `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="0" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="AAAAAAA1"><w:r><w:t>root</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="B" w:date="2020-01-02T00:00:00Z"><w:p w14:paraId="AAAAAAA2"><w:r><w:t>reply</w:t></w:r></w:p></w:comment></w:comments>`;
		const extended = `<?xml version="1.0"?><w15:commentsEx ${NS}><w15:commentEx w15:paraId="AAAAAAA1" w15:done="0"/><w15:commentEx w15:paraId="AAAAAAA2" w15:paraIdParent="AAAAAAA1" w15:done="0"/></w15:commentsEx>`;
		const view = CommentsView.fromXml(comments, extended);
		expect(view?.threadRootId("c1")).toBe("0");
		expect(view?.threadRootId("c0")).toBe("0");
	});

	test("generateParaId stays inside Word's valid range (top bit clear)", async () => {
		const { generateParaId } = await import("@core/ast/document/comments");
		for (let attempt = 0; attempt < 256; attempt++) {
			const minted = Number.parseInt(generateParaId(), 16);
			expect(minted).toBeGreaterThan(0);
			expect(minted).toBeLessThan(0x80000000);
		}
	});

	test("descendantReplyIds walks transitive replies", () => {
		const comments = `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="0" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="AAAAAAA1"><w:r><w:t>root</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="B" w:date="2020-01-02T00:00:00Z"><w:p w14:paraId="AAAAAAA2"><w:r><w:t>reply</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="C" w:date="2020-01-03T00:00:00Z"><w:p w14:paraId="AAAAAAA3"><w:r><w:t>nested</w:t></w:r></w:p></w:comment><w:comment w:id="3" w:author="D" w:date="2020-01-04T00:00:00Z"><w:p w14:paraId="AAAAAAA4"><w:r><w:t>unrelated</w:t></w:r></w:p></w:comment></w:comments>`;
		const extended = `<?xml version="1.0"?><w15:commentsEx ${NS}><w15:commentEx w15:paraId="AAAAAAA1" w15:done="0"/><w15:commentEx w15:paraId="AAAAAAA2" w15:paraIdParent="AAAAAAA1" w15:done="0"/><w15:commentEx w15:paraId="AAAAAAA3" w15:paraIdParent="AAAAAAA2" w15:done="0"/><w15:commentEx w15:paraId="AAAAAAA4" w15:done="0"/></w15:commentsEx>`;
		const view = CommentsView.fromXml(comments, extended);
		expect(view?.descendantReplyIds("c0")?.sort()).toEqual(["1", "2"]);
		expect(view?.descendantReplyIds("c3")).toEqual([]);
	});
});
