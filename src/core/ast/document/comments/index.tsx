import { w, w15 } from "../../../jsx";
import { ensureIgnorable, NS_MC } from "../../../mc";
import { XmlNode } from "../../../parser";
import type { Comment, CommentAnchor } from "../../types";
import type { CommentReference } from "../body";
import type { ContentTypesView } from "../content-types";
import type { Pkg } from "../package";
import type { RelationshipsView } from "../relationships";
import { generateParaId, mintDurableId } from "./para-id";
import {
	COMMENTS_EXT_PART_NAME,
	COMMENTS_EXTENSIBLE_PART_NAME,
	COMMENTS_IDS_PART_NAME,
	COMMENTS_PART_NAME,
	type CommentFinding,
	lastParagraph,
	reconcileSidecars,
	threadRootsOf,
	toDateUtc,
} from "./reconcile-sidecars";

// Word 2016+/2018+ sidecars keyed to each comment's `w14:paraId`
// (`commentsIds`) and to the `durableId` it mints there (`commentsExtensible`).
// We never create them — a package without them is consistent — but when a
// Word-saved input carries them, every comment we add or delete is mirrored,
// and every save reconciles them (`reconcileSidecars`).
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
	idsTree?: XmlNode[];
	extensibleTree?: XmlNode[];
	commentReferences: Map<string, CommentReference> = new Map();

	/** The XML each part was loaded from, so `writeTo` re-writes only a part
	 * whose tree actually changed (an unchanged sidecar keeps its original
	 * bytes instead of a re-serialization). Absent for a part we minted. */
	#sourceXml: Map<string, string>;

	constructor(
		tree: XmlNode[] = [<CommentsRoot />],
		extendedTree?: XmlNode[],
		sidecars: {
			idsTree?: XmlNode[];
			extensibleTree?: XmlNode[];
			sourceXml?: Map<string, string>;
		} = {},
	) {
		this.tree = tree;
		this.extendedTree = extendedTree;
		this.idsTree = sidecars.idsTree;
		this.extensibleTree = sidecars.extensibleTree;
		this.#sourceXml = sidecars.sourceXml ?? new Map();
		// `w14:paraId`/`w14:textId` (Word's threading keys, written on every comment
		// body) are valid only when `w14` rides the root's `mc:Ignorable`. A part
		// that came WITH the document — from Word, which often omits it here, or an
		// earlier docx-cli — may not declare it, so every `comments add`/`reply`
		// made `docx validate` flag a file the agent hadn't broken. Declaring it on
		// load (idempotent; an unused Ignorable prefix is harmless) heals any file
		// we save.
		const root = XmlNode.findRoot(tree, "w:comments");
		if (root) declareW14(root);
	}

	/** Load this view (and its extended/ids/extensible sidecars) from a
	 * package; returns undefined if the base comments part is absent. */
	static async fromPackage(pkg: Pkg): Promise<CommentsView | undefined> {
		if (!pkg.hasPart(COMMENTS_PART_NAME)) return undefined;
		const sourceXml = new Map<string, string>();
		const [tree, extendedTree, idsTree, extensibleTree] = await Promise.all(
			[
				COMMENTS_PART_NAME,
				COMMENTS_EXT_PART_NAME,
				COMMENTS_IDS_PART_NAME,
				COMMENTS_EXTENSIBLE_PART_NAME,
			].map(async (partName) => {
				if (!pkg.hasPart(partName)) return undefined;
				const xml = await pkg.readText(partName);
				sourceXml.set(partName, xml);
				return XmlNode.parse(xml);
			}),
		);
		return new CommentsView(tree, extendedTree, {
			idsTree,
			extensibleTree,
			sourceXml,
		});
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

	/** Reconcile the sidecars (`reconcileSidecars`), then serialize each part
	 * that is present AND changed into the package's `word/comments.xml`,
	 * `commentsExtended.xml`, `commentsIds.xml` and `commentsExtensible.xml`.
	 * Reconciling on every save is what makes ANY mutation cure a file an
	 * older producer (docx-cli ≤ 0.28 included) left inconsistent. */
	writeTo(pkg: Pkg): void {
		this.reconcileSidecars();
		const parts: [string, XmlNode[] | undefined][] = [
			[COMMENTS_PART_NAME, this.tree],
			[COMMENTS_EXT_PART_NAME, this.extendedTree],
			[COMMENTS_IDS_PART_NAME, this.idsTree],
			[COMMENTS_EXTENSIBLE_PART_NAME, this.extensibleTree],
		];
		for (const [partName, partTree] of parts) {
			if (!partTree) continue;
			const xml = XmlNode.serialize(partTree);
			if (this.#isUnchanged(partName, xml)) continue;
			pkg.writeText(partName, xml);
		}
	}

	/** Repair every sidecar inconsistency in place and return what was done
	 * (see `reconcileSidecars` for the rules). Idempotent — a second call
	 * returns `[]`. `writeTo` runs it on every save; `validate --fix` calls it
	 * first to report the repairs. */
	reconcileSidecars(): CommentFinding[] {
		const comments = XmlNode.findRoot(this.tree, "w:comments");
		if (!comments) return [];
		return reconcileSidecars({
			comments,
			extended: sidecarRoot(this.extendedTree, "w15:commentsEx"),
			ids: sidecarRoot(this.idsTree, "w16cid:commentsIds"),
			extensible: sidecarRoot(this.extensibleTree, "w16cex:commentsExtensible"),
		});
	}

	/** Reply id → thread-root id for every threaded comment (see
	 * `threadRootsOf`). */
	threadRoots(): Map<string, string> {
		const comments = XmlNode.findRoot(this.tree, "w:comments");
		if (!comments) return new Map();
		return threadRootsOf({
			comments,
			extended: sidecarRoot(this.extendedTree, "w15:commentsEx"),
		});
	}

	#isUnchanged(partName: string, xml: string): boolean {
		const source = this.#sourceXml.get(partName);
		if (source === undefined) return false;
		return XmlNode.serialize(XmlNode.parse(source)) === xml;
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

	/** Read a comment's `w14:paraId`, minting + persisting one if absent (the
	 * root's `w14` declaration is established on load). reply/resolve key off
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

		// `visited` guards a `paraIdParent` cycle (or a paraId shared by two
		// comments) in a malformed sidecar — without it the walk never ended and
		// `comments delete` hung, growing memory until killed.
		const descendants: string[] = [];
		const visited = new Set([startParaId]);
		const queue = [startParaId];
		while (queue.length > 0) {
			const paraId = queue.pop();
			if (!paraId) break;
			for (const childParaId of childParaIdsByParent.get(paraId) ?? []) {
				if (visited.has(childParaId)) continue;
				visited.add(childParaId);
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

	/** Register a freshly appended comment, keyed by its (last) paragraph's
	 * `paraId`, in every sidecar the package ALREADY carries — the shape Word
	 * writes: a `<w15:commentEx done="0">` (skipped for a `threaded` reply,
	 * whose lens places its own entry with `paraIdParent`), a
	 * `<w16cid:commentId>` with a freshly minted `durableId`, and a
	 * `<w16cex:commentExtensible>` carrying that `durableId` + `dateUtc`. An
	 * absent sidecar stays absent. */
	registerSidecarEntries(
		paraId: string,
		date: string,
		options: { threaded: boolean },
	): void {
		const extendedRoot = sidecarRoot(this.extendedTree, "w15:commentsEx");
		if (extendedRoot && !options.threaded) {
			extendedRoot.children.push(
				new XmlNode("w15:commentEx", {
					"w15:paraId": paraId,
					"w15:done": "0",
				}),
			);
		}
		const idsRoot = sidecarRoot(this.idsTree, "w16cid:commentsIds");
		if (!idsRoot) return;
		const durableId = mintDurableId(
			new Set(
				idsRoot.children.map(
					(child) =>
						child.getAttribute("w16cid:durableId")?.toUpperCase() ?? "",
				),
			),
		);
		idsRoot.children.push(
			new XmlNode("w16cid:commentId", {
				"w16cid:paraId": paraId,
				"w16cid:durableId": durableId,
			}),
		);
		const extensibleRoot = sidecarRoot(
			this.extensibleTree,
			"w16cex:commentsExtensible",
		);
		if (!extensibleRoot) return;
		const dateUtc = toDateUtc(date);
		extensibleRoot.children.push(
			new XmlNode("w16cex:commentExtensible", {
				"w16cex:durableId": durableId,
				...(dateUtc ? { "w16cex:dateUtc": dateUtc } : {}),
			}),
		);
	}

	/** Drop every sidecar entry keyed to a deleted comment's `paraId`: its
	 * `<w15:commentEx>`, its `<w16cid:commentId>`, and the
	 * `<w16cex:commentExtensible>` sharing that entry's `durableId`. */
	removeSidecarEntries(paraId: string): void {
		const extendedRoot = sidecarRoot(this.extendedTree, "w15:commentsEx");
		if (extendedRoot) {
			extendedRoot.children = extendedRoot.children.filter(
				(child) =>
					!(
						child.tag === "w15:commentEx" &&
						child.getAttribute("w15:paraId") === paraId
					),
			);
		}
		const idsRoot = sidecarRoot(this.idsTree, "w16cid:commentsIds");
		if (!idsRoot) return;
		const removedDurableIds = new Set<string>();
		idsRoot.children = idsRoot.children.filter((child) => {
			if (child.tag !== "w16cid:commentId") return true;
			if (child.getAttribute("w16cid:paraId") !== paraId) return true;
			const durableId = child.getAttribute("w16cid:durableId");
			if (durableId) removedDurableIds.add(durableId);
			return false;
		});
		const extensibleRoot = sidecarRoot(
			this.extensibleTree,
			"w16cex:commentsExtensible",
		);
		if (!extensibleRoot) return;
		extensibleRoot.children = extensibleRoot.children.filter(
			(child) =>
				!(
					child.tag === "w16cex:commentExtensible" &&
					removedDurableIds.has(child.getAttribute("w16cex:durableId") ?? "")
				),
		);
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
			// Break/tab-aware so a comment authored with `\n`/`\t` reads back
			// faithfully instead of joining words (the markdown renderer collapses
			// `\s+` for its single-line `[^cN]:` definition).
			const text = child.collectTextWithBreaks();
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
	 * allocates the in-memory tree. A part materialized mid-life (the first
	 * reply/resolve on a doc whose comments predate it — every `create`d doc)
	 * gets a `<w15:commentEx done="0">` for EVERY comment already there, the
	 * shape Word writes: once the part exists, a comment with no entry is the
	 * inconsistency `validate` reports (`comments add` → `add` → `reply` used
	 * to leave the second comment unregistered). Replies must therefore call
	 * this BEFORE appending their own body, which registers itself threaded. */
	ensureExtended(): XmlNode[] {
		if (this.extendedTree) return this.extendedTree;
		const extendedRoot = <CommentsExRoot />;
		for (const commentId of this.listIds()) {
			const paraId = this.ensureParaId(commentId);
			if (!paraId) continue;
			extendedRoot.children.push(
				new XmlNode("w15:commentEx", {
					"w15:paraId": paraId,
					"w15:done": "0",
				}),
			);
		}
		this.extendedTree = [extendedRoot];
		return this.extendedTree;
	}
}

/** Declare `xmlns:w14` and register `w14` in `mc:Ignorable` on the part root.
 *  Idempotent. */
function declareW14(root: XmlNode): void {
	if (!root.getAttribute("xmlns:w14")) root.setAttribute("xmlns:w14", NS_W14);
	ensureIgnorable(root, "w14");
}

function CommentsRoot(): XmlNode {
	// `w14:paraId`/`w14:textId` are extension attributes: per OOXML's
	// markup-compatibility rules they're only valid when their prefix rides
	// `mc:Ignorable` on the part root (Word declares exactly this on its own
	// comments.xml). Without it the part is schema-invalid.
	return (
		<w.comments
			{...{
				"xmlns:w": NS_W,
				"xmlns:w14": NS_W14,
				"xmlns:mc": NS_MC,
				"mc:Ignorable": "w14",
			}}
		/>
	);
}

function sidecarRoot(
	tree: XmlNode[] | undefined,
	rootTag: string,
): XmlNode | undefined {
	return tree ? XmlNode.findRoot(tree, rootTag) : undefined;
}

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
