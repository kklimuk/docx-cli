import { describe, expect, test } from "bun:test";
import { CommentsView } from "../../../src/core/ast/document/comments";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;

function multiParagraphComments(): string {
	return `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="18" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="0000027A"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="00000283"><w:r><w:t>last</w:t></w:r></w:p></w:comment></w:comments>`;
}

describe("comment threading identity keys off the last paragraph", () => {
	test("paraIdFor returns the last paragraph's paraId", () => {
		const view = CommentsView.fromXml(multiParagraphComments());
		expect(view?.paraIdFor("18")).toBe("00000283");
	});

	test("ensureParaId reuses the last paragraph's existing paraId", () => {
		const view = CommentsView.fromXml(multiParagraphComments());
		expect(view?.ensureParaId("c18")).toBe("00000283");
	});

	test("a reply linked to the last paragraph resolves its parentId", () => {
		const comments = `<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="18" w:author="A" w:date="2020-01-01T00:00:00Z"><w:p w14:paraId="0000027A"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="00000283"><w:r><w:t>last</w:t></w:r></w:p></w:comment><w:comment w:id="19" w:author="B" w:date="2020-01-02T00:00:00Z"><w:p w14:paraId="80CBDBBD" w14:paraIdParent="00000283"><w:r><w:t>OK</w:t></w:r></w:p></w:comment></w:comments>`;
		const extended = `<?xml version="1.0"?><w15:commentsEx ${NS}><w15:commentEx w15:paraId="00000283" w15:done="0"/><w15:commentEx w15:paraId="80CBDBBD" w15:paraIdParent="00000283" w15:done="0"/></w15:commentsEx>`;
		const view = CommentsView.fromXml(comments, extended);
		const list = view?.toComments(new Map());
		const reply = list?.find((comment) => comment.id === "c19");
		expect(reply?.parentId).toBe("c18");
	});
});
