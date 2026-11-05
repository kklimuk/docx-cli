import type { XmlNode } from "../parser";

/** Package-level consistency audit — the cross-part invariants no per-part
 *  schema can see, run by `docx validate` beside the XSD check. Every finding
 *  is something a schema-valid package can still get wrong and that Word
 *  repairs (or silently drops) on open:
 *
 *  - **Comment side-parts.** `commentsExtended.xml` (`w15:commentEx`) and
 *    `commentsIds.xml` (`w16cid:commentId`) key each comment by the
 *    `w14:paraId` of its LAST paragraph. When a side-part is present, a
 *    comment with no entry in it is a Word repair trigger (issue #33's
 *    "unreadable content"). An absent side-part is legal — older producers
 *    write none — so the check is per present part.
 *  - **Comment marker balance** across every story part (body, headers,
 *    footers, notes): a `commentRangeStart` needs its `commentRangeEnd` and a
 *    `commentReference`, and a reference must name a comment that exists.
 *
 *  Ids are compared as SETS: `mc:AlternateContent` writes the same text-box
 *  story under both Choice and Fallback, so a marker appearing twice is
 *  legal. Orphan side-part entries and comments with no body markers are not
 *  reported — Word drops those quietly, so flagging them would be noise.
 *
 *  The glossary document (`word/glossary/…`, building blocks) is a separate
 *  story set with its own comment parts, so it's excluded rather than mixed
 *  into the main document's id spaces. */
export function checkPackageConsistency(
	packageParts: PackagePart[],
): ConsistencyIssue[] {
	const parts = packageParts.filter(
		(part) => !part.name.startsWith("word/glossary/"),
	);
	const comments = collectComments(parts);
	return [
		...checkCommentSideParts(parts, comments),
		...checkCommentMarkers(parts, comments),
	];
}

const COMMENTS_ROOT = "w:comments";

/** The comment side-parts keyed by paraId: root tag → entry tag + key attr. */
const SIDE_PARTS = [
	{
		root: "w15:commentsEx",
		entry: "w15:commentEx",
		paraIdAttribute: "w15:paraId",
	},
	{
		root: "w16cid:commentsIds",
		entry: "w16cid:commentId",
		paraIdAttribute: "w16cid:paraId",
	},
] as const;

const STORY_ROOTS = new Set([
	"w:document",
	"w:hdr",
	"w:ftr",
	"w:footnotes",
	"w:endnotes",
]);

function collectComments(parts: PackagePart[]): CommentsSet {
	const commentsPart = parts.find((part) => part.root.tag === COMMENTS_ROOT);
	if (!commentsPart) return undefined;
	const entries: CommentEntry[] = [];
	for (const comment of commentsPart.root.findChildren("w:comment")) {
		const id = comment.getAttribute("w:id");
		if (id === undefined) continue;
		// A comment's threading identity is its LAST `<w:p>` (see CommentsView).
		const paraId = comment
			.findChildren("w:p")
			.at(-1)
			?.getAttribute("w14:paraId");
		entries.push({ id, ...(paraId ? { paraId } : {}) });
	}
	return { part: commentsPart.name, entries };
}

function checkCommentSideParts(
	parts: PackagePart[],
	comments: CommentsSet,
): ConsistencyIssue[] {
	if (!comments) return [];
	const issues: ConsistencyIssue[] = [];
	for (const sidePart of SIDE_PARTS) {
		const part = parts.find(
			(candidate) => candidate.root.tag === sidePart.root,
		);
		if (!part) continue;
		const keyed = new Set(
			part.root
				.findChildren(sidePart.entry)
				.map((entry) => entry.getAttribute(sidePart.paraIdAttribute)),
		);
		for (const comment of comments.entries) {
			if (comment.paraId && keyed.has(comment.paraId)) continue;
			issues.push({
				part: part.name,
				message: comment.paraId
					? `comment c${comment.id} (paraId ${comment.paraId}) has no entry in ${part.name}`
					: `comment c${comment.id} has no w14:paraId, so it has no entry in ${part.name}`,
			});
		}
	}
	return issues;
}

function checkCommentMarkers(
	parts: PackagePart[],
	comments: CommentsSet,
): ConsistencyIssue[] {
	const markers = {
		"w:commentRangeStart": new Map<string, string>(),
		"w:commentRangeEnd": new Map<string, string>(),
		"w:commentReference": new Map<string, string>(),
	};
	for (const part of parts) {
		if (!STORY_ROOTS.has(part.root.tag)) continue;
		collectMarkers(part.root, part.name, markers);
	}
	const starts = markers["w:commentRangeStart"];
	const ends = markers["w:commentRangeEnd"];
	const references = markers["w:commentReference"];
	const commentIds = new Set(comments?.entries.map((entry) => entry.id));
	const commentsPartName = comments?.part ?? "word/comments.xml";

	const issues: ConsistencyIssue[] = [];
	for (const [id, part] of starts) {
		if (!ends.has(id)) {
			issues.push({
				part,
				message: `commentRangeStart ${id} has no matching commentRangeEnd`,
			});
		}
	}
	for (const [id, part] of ends) {
		if (!starts.has(id)) {
			issues.push({
				part,
				message: `commentRangeEnd ${id} has no matching commentRangeStart`,
			});
		}
	}
	for (const [id, part] of new Map([...ends, ...starts])) {
		if (!references.has(id)) {
			issues.push({
				part,
				message: `comment range ${id} has no matching commentReference`,
			});
		}
	}
	for (const [id, part] of references) {
		if (!commentIds.has(id)) {
			issues.push({
				part,
				message: `commentReference ${id} points at no comment in ${commentsPartName}`,
			});
		}
	}
	return issues;
}

/** Record the first part each comment marker id appears in, per marker tag. */
function collectMarkers(
	node: XmlNode,
	partName: string,
	markers: Record<string, Map<string, string>>,
): void {
	for (const child of node.children) {
		if (child.isText) continue;
		const seen = markers[child.tag];
		if (seen) {
			const id = child.getAttribute("w:id");
			if (id !== undefined && !seen.has(id)) seen.set(id, partName);
			continue;
		}
		collectMarkers(child, partName, markers);
	}
}

export type PackagePart = { name: string; root: XmlNode };

export type ConsistencyIssue = { part: string; message: string };

type CommentEntry = { id: string; paraId?: string };

type CommentsSet = { part: string; entries: CommentEntry[] } | undefined;
