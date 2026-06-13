import { w, w15 } from "../../jsx";
import { XmlNode } from "../../parser";
import type { Comment, CommentAnchor } from "../types";
import type { CommentReference } from "./body";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

const COMMENTS_PART_NAME = "word/comments.xml";
const COMMENTS_EXT_PART_NAME = "word/commentsExtended.xml";
const COMMENTS_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXT_RELATIONSHIP_TYPE =
	"http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXT_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

export class CommentsView {
	tree: XmlNode[];
	extendedTree?: XmlNode[];
	commentReferences: Map<string, CommentReference> = new Map();

	constructor(tree: XmlNode[] = [<CommentsRoot />], extendedTree?: XmlNode[]) {
		this.tree = tree;
		this.extendedTree = extendedTree;
	}

	/** Load this view (and its extended sidecar) from a package; returns
	 * undefined if the base comments part is absent. */
	static async fromPackage(pkg: Pkg): Promise<CommentsView | undefined> {
		const tree = await pkg.readPart(COMMENTS_PART_NAME);
		if (!tree) return undefined;
		const extendedTree = await pkg.readPart(COMMENTS_EXT_PART_NAME);
		return new CommentsView(tree, extendedTree);
	}

	/** Parse a view from raw XML; returns undefined if the base input is
	 * absent. The extended sidecar follows the base — if base is undefined,
	 * extended is ignored. */
	static fromXml(
		xml: string | undefined,
		extendedXml?: string | undefined,
	): CommentsView | undefined {
		if (!xml) return undefined;
		return new CommentsView(
			XmlNode.parse(xml),
			extendedXml ? XmlNode.parse(extendedXml) : undefined,
		);
	}

	/** Serialize this view's tree (and the extended sidecar if present) into
	 * the package's `word/comments.xml` and `word/commentsExtended.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(COMMENTS_PART_NAME, XmlNode.serialize(this.tree));
		if (this.extendedTree) {
			pkg.writeText(
				COMMENTS_EXT_PART_NAME,
				XmlNode.serialize(this.extendedTree),
			);
		}
	}

	/** Mint the comments relationship + content-type override on the
	 * containing package and return a fresh empty view. Idempotent on the
	 * relationship target. Called by `Document.ensureComments()`. */
	static register(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): CommentsView {
		if (!deps.relationships.hasTarget("comments.xml")) {
			deps.relationships.add(COMMENTS_RELATIONSHIP_TYPE, "comments.xml");
		}
		deps.contentTypes.registerPart(COMMENTS_PART_NAME, COMMENTS_CONTENT_TYPE);
		return new CommentsView();
	}

	/** Mint the commentsExtended relationship + content-type override on the
	 * containing package and materialize the extended tree on this view.
	 * Idempotent on both the relationship target and the in-memory tree.
	 * Called by `Document.ensureCommentsExtended()` (which calls
	 * `ensureComments()` first to guarantee the base part). */
	registerExtended(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): void {
		this.ensureExtended();
		if (!deps.relationships.hasTarget("commentsExtended.xml")) {
			deps.relationships.add(
				COMMENTS_EXT_RELATIONSHIP_TYPE,
				"commentsExtended.xml",
			);
		}
		deps.contentTypes.registerPart(
			COMMENTS_EXT_PART_NAME,
			COMMENTS_EXT_CONTENT_TYPE,
		);
	}

	listIds(): string[] {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) return [];
		const out: string[] = [];
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const id = child.getAttribute("w:id");
			if (id) out.push(id);
		}
		return out;
	}

	hasId(id: string): boolean {
		return this.findById(id) !== undefined;
	}

	findById(id: string): XmlNode | undefined {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) return undefined;
		return root.children.find(
			(child) => child.tag === "w:comment" && child.getAttribute("w:id") === id,
		);
	}

	/** Read a comment's threading `w14:paraId`, or undefined if the comment is
	 * missing or has no inner `<w:p>`. Word keys `<w15:commentEx>` and a
	 * reply's `w15:paraIdParent` off the comment's LAST paragraph, so
	 * threading must too. Accepts the `cN` or bare-`N` id form. */
	paraIdFor(commentId: string): string | undefined {
		const numericId = commentId.startsWith("c")
			? commentId.slice(1)
			: commentId;
		return lastParagraph(this.findById(numericId))?.getAttribute("w14:paraId");
	}

	/** Read a comment's `w14:paraId`, minting + persisting one (and the
	 * `xmlns:w14` declaration on the root) if absent. reply/resolve key off
	 * paraId, so this guarantees one exists. Returns undefined only when the
	 * comment itself is missing. Accepts the `cN` or bare-`N` id form. */
	ensureParaId(commentId: string): string | undefined {
		const numericId = commentId.startsWith("c")
			? commentId.slice(1)
			: commentId;
		const root = XmlNode.findRoot(this.tree, "w:comments");
		const paragraph = lastParagraph(this.findById(numericId));
		if (!root || !paragraph) return undefined;
		const existing = paragraph.getAttribute("w14:paraId");
		if (existing) return existing;
		const fresh = generateParaId();
		paragraph.setAttribute("w14:paraId", fresh);
		if (!root.attributes["xmlns:w14"]) {
			root.setAttribute("xmlns:w14", NS_W14);
		}
		return fresh;
	}

	/** Word reserves comment ids starting from 0. We allocate `max(existing
	 * id) + 1` so we never collide. Returns "0" when the part is empty (or only
	 * contains the root element with no comments). */
	nextId(): string {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) return "0";
		let highest = -1;
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const idAttribute = child.getAttribute("w:id");
			if (idAttribute == null) continue;
			const numeric = Number(idAttribute);
			if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
		}
		return String(highest + 1);
	}

	/** Resolve the thread root of `commentId` by walking `w15:paraIdParent`
	 * links upward; returns the input's numeric id when it has no parent.
	 * Word threads are single-level — its UI attaches a reply-to-a-reply to
	 * the thread ROOT, and it rewrites deeper chains to the root on open — so
	 * new replies chain to the root for round-trip stability. Accepts the
	 * `cN` or bare-`N` id form; returns the bare-`N` form. */
	threadRootId(commentId: string): string {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		const numericId = commentId.startsWith("c")
			? commentId.slice(1)
			: commentId;
		if (!root) return numericId;

		const extended = this.#readExtended();
		const idByParaId = new Map<string, string>();
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const childId = child.getAttribute("w:id");
			if (childId == null) continue;
			const paraId = lastParagraph(child)?.getAttribute("w14:paraId");
			if (paraId) idByParaId.set(paraId, childId);
		}

		let currentId = numericId;
		const seen = new Set([currentId]);
		for (;;) {
			const paraId = lastParagraph(this.findById(currentId))?.getAttribute(
				"w14:paraId",
			);
			const parentParaId = paraId
				? extended.get(paraId)?.parentParaId
				: undefined;
			const parentId = parentParaId ? idByParaId.get(parentParaId) : undefined;
			if (!parentId || seen.has(parentId)) return currentId;
			seen.add(parentId);
			currentId = parentId;
		}
	}

	/** Numeric ids of every transitive reply chained under `commentId` via
	 * `w15:paraIdParent`. Delete cascades through these so a removed parent
	 * never strands its replies' body markers or dangles their thread link.
	 * Accepts the `cN` or bare-`N` id form. */
	descendantReplyIds(commentId: string): string[] {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) return [];
		const numericId = commentId.startsWith("c")
			? commentId.slice(1)
			: commentId;
		const startParaId = lastParagraph(this.findById(numericId))?.getAttribute(
			"w14:paraId",
		);
		if (!startParaId) return [];

		const extended = this.#readExtended();
		const idByParaId = new Map<string, string>();
		const childParaIdsByParent = new Map<string, string[]>();
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const childId = child.getAttribute("w:id");
			if (childId == null) continue;
			const paraId = lastParagraph(child)?.getAttribute("w14:paraId");
			if (!paraId) continue;
			idByParaId.set(paraId, childId);
			const parentParaId = extended.get(paraId)?.parentParaId;
			if (!parentParaId) continue;
			const siblings = childParaIdsByParent.get(parentParaId) ?? [];
			siblings.push(paraId);
			childParaIdsByParent.set(parentParaId, siblings);
		}

		const descendants: string[] = [];
		const queue = [startParaId];
		while (queue.length > 0) {
			const paraId = queue.pop();
			if (!paraId) break;
			for (const childParaId of childParaIdsByParent.get(paraId) ?? []) {
				const childId = idByParaId.get(childParaId);
				if (childId) descendants.push(childId);
				queue.push(childParaId);
			}
		}
		return descendants;
	}

	/** Insert `comment` (a fresh `<w:comment>`) immediately after the last
	 * comment in the thread rooted at `parentNumericId` (the parent plus every
	 * transitive reply), so a new reply lands after the existing thread tail
	 * instead of after some other thread at the end of the part — matching how
	 * Word groups a thread's comments together. */
	insertReplyAfter(parentNumericId: string, comment: XmlNode): void {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) throw new Error("expected <w:comments> root");
		const threadIds = new Set([
			parentNumericId,
			...this.descendantReplyIds(parentNumericId),
		]);
		let insertIndex = root.children.length;
		for (const [index, child] of root.children.entries()) {
			if (child.tag !== "w:comment") continue;
			const id = child.getAttribute("w:id");
			if (id != null && threadIds.has(id)) insertIndex = index + 1;
		}
		root.children.splice(insertIndex, 0, comment);
	}

	/** Parse this part (plus the extended sidecar) into the AST `Comment[]` the
	 * reader exposes on `Body.comments`, and populate `commentReferences` for
	 * mutators. `anchors` are the span ranges the body walk computed from
	 * `<w:commentRangeStart>`/`<w:commentRangeEnd>` markers — the one input that
	 * can't be derived from the comments part alone, so the reader passes it in. */
	toComments(anchors: Map<string, CommentAnchor>): Comment[] {
		const root = XmlNode.findRoot(this.tree, "w:comments");
		if (!root) return [];

		const commentIdByParaId = new Map<string, string>();
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const numericId = child.getAttribute("w:id");
			if (numericId == null) continue;
			// Threading keys off the LAST paragraph (see `paraIdFor`), so index
			// by it to resolve a reply's `w15:paraIdParent` back to its parent.
			const paraId = lastParagraph(child)?.getAttribute("w14:paraId");
			if (paraId) commentIdByParaId.set(paraId, `c${numericId}`);
		}

		const extendedByParaId = this.#readExtended();
		const comments: Comment[] = [];

		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			const numericId = child.getAttribute("w:id");
			if (numericId == null) continue;
			const commentId = `c${numericId}`;
			const author = child.getAttribute("w:author") ?? "";
			const date = child.getAttribute("w:date") ?? "";
			const initials = child.getAttribute("w:initials");
			const text = child.collectText();
			const anchor = anchors.get(commentId) ?? {
				startBlockId: "",
				startOffset: 0,
				endBlockId: "",
				endOffset: 0,
			};

			const paraId = lastParagraph(child)?.getAttribute("w14:paraId");
			const meta = paraId ? (extendedByParaId.get(paraId) ?? {}) : {};
			const parentCommentId = meta.parentParaId
				? commentIdByParaId.get(meta.parentParaId)
				: undefined;

			comments.push({
				id: commentId,
				author,
				...(initials ? { initials } : {}),
				date,
				text,
				anchor,
				...(parentCommentId ? { parentId: parentCommentId } : {}),
				...(meta.resolved !== undefined ? { resolved: meta.resolved } : {}),
			});
			this.commentReferences.set(commentId, {
				node: child,
				parent: root.children,
			});
		}
		return comments;
	}

	/** Index `commentsExtended.xml` by paraId → { parentParaId, resolved }.
	 * `w15:done` carries the resolved bit; `w15:paraIdParent` links a reply to
	 * its thread parent. Empty map when the sidecar is absent. */
	#readExtended(): Map<string, { parentParaId?: string; resolved?: boolean }> {
		const out = new Map<
			string,
			{ parentParaId?: string; resolved?: boolean }
		>();
		if (!this.extendedTree) return out;
		const root = XmlNode.findRoot(this.extendedTree, "w15:commentsEx");
		if (!root) return out;

		for (const child of root.children) {
			if (child.tag !== "w15:commentEx") continue;
			const paragraphId = child.getAttribute("w15:paraId");
			if (!paragraphId) continue;
			const resolvedAttribute = child.getAttribute("w15:done");
			const parentParagraphId = child.getAttribute("w15:paraIdParent");
			const entry: { parentParaId?: string; resolved?: boolean } = {};
			if (resolvedAttribute === "1") entry.resolved = true;
			else if (resolvedAttribute === "0") entry.resolved = false;
			if (parentParagraphId) entry.parentParaId = parentParagraphId;
			out.set(paragraphId, entry);
		}
		return out;
	}

	/** Materialize the extended-comments tree if absent. Caller is responsible
	 * for minting the rel + registering the content-type via Document — this just
	 * allocates the in-memory tree. */
	ensureExtended(): XmlNode[] {
		if (this.extendedTree) return this.extendedTree;
		this.extendedTree = [<CommentsExRoot />];
		return this.extendedTree;
	}
}

function CommentsRoot(): XmlNode {
	return <w.comments {...{ "xmlns:w": NS_W, "xmlns:w14": NS_W14 }} />;
}

/** A comment's threading identity is its LAST `<w:p>` — that's the paragraph
 * Word keys `<w15:commentEx>` and a reply's `w15:paraIdParent` to. */
function lastParagraph(comment: XmlNode | undefined): XmlNode | undefined {
	return comment?.findChildren("w:p").at(-1);
}

/** Mint a fresh `w14:paraId` value. Word writes 8-char uppercase hex; we
 * match. Per MS-DOCX, a paraId MUST be greater than 0x00000000 and less
 * than 0x80000000 — Word silently re-keys out-of-range ids on open, which
 * orphans every sidecar entry (`w15:commentEx` resolved flags, reply
 * `w15:paraIdParent` links) keyed to the old value. Used both by
 * `CommentsView.ensureParaId` and by the comments lens for replies + audit
 * comments.
 *
 * `DOCX_CLI_PARA_ID_SEED` (8-char hex) switches to sequential minting from
 * that base — the byte-determinism seam for fixture rebuilds, mirroring
 * `DOCX_CLI_NOW`. The counter is per-process: a builder that invokes the CLI
 * multiple times must set a DISTINCT seed per invocation or ids collide. */
export function generateParaId(): string {
	const seed = Bun.env.DOCX_CLI_PARA_ID_SEED;
	if (seed) {
		const minted = (Number.parseInt(seed, 16) + seededParaIds++) % 0x80000000;
		return (minted || 1).toString(16).padStart(8, "0").toUpperCase();
	}
	const bytes = new Uint8Array(4);
	crypto.getRandomValues(bytes);
	bytes[0] = (bytes[0] ?? 0) & 0x7f;
	const hex = [...bytes]
		.map((byte) => byte.toString(16).padStart(2, "0"))
		.join("");
	if (hex === "00000000") return "00000001";
	return hex.toUpperCase();
}

let seededParaIds = 0;

function CommentsExRoot(): XmlNode {
	return (
		<w15.commentsEx
			{...{
				"xmlns:w15": NS_W15,
				"xmlns:mc":
					"http://schemas.openxmlformats.org/markup-compatibility/2006",
				"mc:Ignorable": "w15",
			}}
		/>
	);
}
