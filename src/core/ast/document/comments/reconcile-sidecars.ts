import { XmlNode } from "../../../parser";
import { isValidDurableId, mintDurableId, mintUniqueParaId } from "./para-id";

/** Bring Word's paraId-keyed comment sidecars back in line with
 *  `word/comments.xml`, mutating the roots in place — the same repairs Word
 *  for Mac (16.114) makes when it opens and re-saves the package:
 *
 *  - every comment's LAST paragraph gets a `w14:paraId` (minted unique);
 *  - in each sidecar the package HAS (an absent one stays absent): an entry
 *    is added for every comment without one, and entries whose comment is gone
 *    are pruned;
 *  - `commentsIds.xml` durableIds are de-duplicated (and out-of-range ones
 *    re-minted); a new entry's durableId is the comment's paraId when that is
 *    valid and free — Word's own choice;
 *  - a present `commentsExtensible.xml` follows `commentsIds.xml`'s durableIds.
 *
 *  An added `<w15:commentEx>` is `done="0"` with no parent — all Word can infer
 *  too, which is why a missing entry on a reply or a resolved comment is data
 *  loss (`validate` says so). An existing entry's `w15:paraIdParent`/`w15:done`
 *  is never touched. Returns one finding per repair (`[]` when the parts were
 *  already consistent, so a caller can tell "nothing to write"); `docx
 *  validate` runs this on a clone to REPORT, `--fix` and every save run it for
 *  real — one rule set, so a reported finding is always one the fix cures. */
export function reconcileSidecars(roots: SidecarRoots): CommentFinding[] {
	const findings: CommentFinding[] = [];
	const comments = keyComments(roots.comments, findings);
	if (roots.extended) reconcileExtended(roots.extended, comments, findings);
	if (roots.ids) {
		reconcileIds(roots.ids, comments, findings);
		if (roots.extensible) {
			reconcileExtensible(roots.extensible, roots.ids, comments, findings);
		}
	}
	return findings;
}

export const COMMENTS_PART_NAME = "word/comments.xml";
export const COMMENTS_EXT_PART_NAME = "word/commentsExtended.xml";
export const COMMENTS_IDS_PART_NAME = "word/commentsIds.xml";
export const COMMENTS_EXTENSIBLE_PART_NAME = "word/commentsExtensible.xml";

/** Report every reply whose `w15:paraIdParent` names no comment — a link
 *  Word silently drops on open, showing the reply as a top-level comment.
 *  There is no deterministic cure (which thread was meant is lost), and the
 *  reconcile never alters an existing `paraIdParent`, so this is report-only. */
export function auditThreadLinks(roots: SidecarRoots): CommentFinding[] {
	if (!roots.extended) return [];
	const commentIdByKey = new Map<string, string>();
	for (const comment of roots.comments.findChildren("w:comment")) {
		const id = comment.getAttribute("w:id");
		const paraId = lastParagraph(comment)?.getAttribute("w14:paraId");
		if (id !== undefined && paraId) commentIdByKey.set(key(paraId), id);
	}
	const findings: CommentFinding[] = [];
	for (const entry of roots.extended.findChildren("w15:commentEx")) {
		const commentId = commentIdByKey.get(key(entry.getAttribute("w15:paraId")));
		const parent = entry.getAttribute("w15:paraIdParent");
		if (commentId === undefined || !parent) continue;
		if (commentIdByKey.has(key(parent))) continue;
		findings.push({
			part: COMMENTS_EXT_PART_NAME,
			severity: "warning",
			message: `comment c${commentId}'s thread parent (w15:paraIdParent ${parent}) names no comment — Word drops the link and shows c${commentId} as a top-level comment`,
			hint: `if c${commentId} belongs under another comment, re-post it there with \`docx comments reply\``,
		});
	}
	return findings;
}

/** Map each reply's comment id to its thread ROOT's id by walking
 *  `w15:paraIdParent` links upward (a broken link or a cycle stops the walk;
 *  a comment with no parent is absent). The marker audit uses it to re-anchor
 *  a reply an old producer left without body markers. */
export function threadRootsOf(roots: {
	comments: XmlNode;
	extended?: XmlNode;
}): Map<string, string> {
	const out = new Map<string, string>();
	if (!roots.extended) return out;
	const commentIdByKey = new Map<string, string>();
	for (const comment of roots.comments.findChildren("w:comment")) {
		const id = comment.getAttribute("w:id");
		const paraId = lastParagraph(comment)?.getAttribute("w14:paraId");
		if (id !== undefined && paraId) commentIdByKey.set(key(paraId), id);
	}
	const parentByKey = new Map<string, string>();
	for (const entry of roots.extended.findChildren("w15:commentEx")) {
		const parent = entry.getAttribute("w15:paraIdParent");
		if (parent)
			parentByKey.set(key(entry.getAttribute("w15:paraId")), key(parent));
	}
	for (const [paraKey, id] of commentIdByKey) {
		let current = paraKey;
		const seen = new Set([current]);
		for (;;) {
			const parent = parentByKey.get(current);
			if (!parent || seen.has(parent) || !commentIdByKey.has(parent)) break;
			seen.add(parent);
			current = parent;
		}
		const rootId = commentIdByKey.get(current);
		if (rootId !== undefined && current !== paraKey) out.set(id, rootId);
	}
	return out;
}

/** Every `<w:comment>` with its threading key (its last paragraph's paraId),
 * minting the paraId where it is missing. A comment with no `<w:p>` at all
 * has nothing to key and is skipped. */
function keyComments(
	commentsRoot: XmlNode,
	findings: CommentFinding[],
): KeyedComment[] {
	const taken = new Set<string>();
	collectParaIds(commentsRoot, taken);
	const keyed: KeyedComment[] = [];
	for (const comment of commentsRoot.findChildren("w:comment")) {
		const id = comment.getAttribute("w:id");
		const paragraph = lastParagraph(comment);
		if (id === undefined || !paragraph) continue;
		const existing = paragraph.getAttribute("w14:paraId");
		const date = comment.getAttribute("w:date") ?? "";
		if (existing) {
			keyed.push({ id, paraId: existing, hadParaId: true, date });
			continue;
		}
		const fresh = mintUniqueParaId(taken);
		taken.add(fresh);
		paragraph.setAttribute("w14:paraId", fresh);
		keyed.push({ id, paraId: fresh, hadParaId: false, date });
		findings.push({
			part: COMMENTS_PART_NAME,
			severity: "silent",
			message: `comment c${id} has no w14:paraId`,
			fix: `minted w14:paraId ${fresh} for c${id}`,
		});
	}
	return keyed;
}

function collectParaIds(node: XmlNode, into: Set<string>): void {
	for (const child of node.children) {
		if (child.isText) continue;
		const paraId = child.getAttribute("w14:paraId");
		if (paraId) into.add(paraId.toUpperCase());
		collectParaIds(child, into);
	}
}

function reconcileExtended(
	root: XmlNode,
	comments: KeyedComment[],
	findings: CommentFinding[],
): void {
	const live = new Set(comments.map((comment) => key(comment.paraId)));
	const present = new Set<string>();
	root.children = root.children.filter((child) => {
		if (child.tag !== "w15:commentEx") return true;
		const paraId = child.getAttribute("w15:paraId");
		if (paraId && live.has(key(paraId))) {
			present.add(key(paraId));
			return true;
		}
		findings.push(staleEntry(COMMENTS_EXT_PART_NAME, paraId));
		return false;
	});
	for (const comment of comments) {
		if (present.has(key(comment.paraId))) continue;
		root.children.push(
			new XmlNode("w15:commentEx", {
				"w15:paraId": comment.paraId,
				"w15:done": "0",
			}),
		);
		findings.push({
			part: COMMENTS_EXT_PART_NAME,
			severity: "warning",
			message: `${missingEntryLead(comment, COMMENTS_EXT_PART_NAME)} — Word re-adds it as an open, top-level comment, so if c${comment.id} was a reply it loses its thread and if it was resolved it reopens`,
			fix: `added c${comment.id}'s entry (open, top-level)`,
		});
	}
}

function reconcileIds(
	root: XmlNode,
	comments: KeyedComment[],
	findings: CommentFinding[],
): void {
	const commentIdByKey = new Map(
		comments.map((comment) => [key(comment.paraId), comment.id]),
	);
	const present = new Set<string>();
	root.children = root.children.filter((child) => {
		if (child.tag !== "w16cid:commentId") return true;
		const paraId = child.getAttribute("w16cid:paraId");
		if (
			paraId &&
			commentIdByKey.has(key(paraId)) &&
			!present.has(key(paraId))
		) {
			present.add(key(paraId));
			return true;
		}
		findings.push(staleEntry(COMMENTS_IDS_PART_NAME, paraId));
		return false;
	});

	const durableIds = new Set<string>();
	for (const child of root.findChildren("w16cid:commentId")) {
		const durableId = child.getAttribute("w16cid:durableId");
		if (isValidDurableId(durableId) && !durableIds.has(key(durableId))) {
			durableIds.add(key(durableId));
			continue;
		}
		const fresh = mintDurableId(durableIdsIn(root, durableIds));
		durableIds.add(fresh);
		child.setAttribute("w16cid:durableId", fresh);
		const commentId = commentIdByKey.get(
			key(child.getAttribute("w16cid:paraId")),
		);
		findings.push({
			part: COMMENTS_IDS_PART_NAME,
			severity: "silent",
			message: `c${commentId}'s durableId ${durableId ?? "(none)"} is ${durableId && isValidDurableId(durableId) ? "a duplicate" : "invalid"}`,
			fix: `re-minted c${commentId}'s durableId ${durableId ?? "(none)"} → ${fresh}`,
		});
	}

	for (const comment of comments) {
		if (present.has(key(comment.paraId))) continue;
		const durableId =
			isValidDurableId(comment.paraId) && !durableIds.has(key(comment.paraId))
				? key(comment.paraId)
				: mintDurableId(durableIdsIn(root, durableIds));
		durableIds.add(durableId);
		root.children.push(
			new XmlNode("w16cid:commentId", {
				"w16cid:paraId": comment.paraId,
				"w16cid:durableId": durableId,
			}),
		);
		findings.push({
			part: COMMENTS_IDS_PART_NAME,
			severity: "warning",
			message: `${missingEntryLead(comment, COMMENTS_IDS_PART_NAME)} — Word mints one on open`,
			fix: `added c${comment.id}'s entry (durableId ${durableId})`,
		});
	}
}

function reconcileExtensible(
	root: XmlNode,
	idsRoot: XmlNode,
	comments: KeyedComment[],
	findings: CommentFinding[],
): void {
	const dateByKey = new Map(
		comments.map((comment) => [key(comment.paraId), comment.date]),
	);
	const commentIdByKey = new Map(
		comments.map((comment) => [key(comment.paraId), comment.id]),
	);
	const live = new Map<string, string>();
	for (const entry of idsRoot.findChildren("w16cid:commentId")) {
		const durableId = entry.getAttribute("w16cid:durableId");
		const paraId = entry.getAttribute("w16cid:paraId");
		if (durableId && paraId) live.set(key(durableId), key(paraId));
	}
	const present = new Set<string>();
	root.children = root.children.filter((child) => {
		if (child.tag !== "w16cex:commentExtensible") return true;
		const durableId = child.getAttribute("w16cex:durableId");
		if (durableId && live.has(key(durableId)) && !present.has(key(durableId))) {
			present.add(key(durableId));
			return true;
		}
		findings.push({
			part: COMMENTS_EXTENSIBLE_PART_NAME,
			severity: "silent",
			message: `${COMMENTS_EXTENSIBLE_PART_NAME} entry for durableId ${durableId ?? "(none)"} names no comment`,
			fix: `pruned the entry for durableId ${durableId ?? "(none)"}`,
		});
		return false;
	});
	for (const [durableId, paraKey] of live) {
		if (present.has(durableId)) continue;
		const dateUtc = toDateUtc(dateByKey.get(paraKey) ?? "");
		root.children.push(
			new XmlNode("w16cex:commentExtensible", {
				"w16cex:durableId": durableId,
				...(dateUtc ? { "w16cex:dateUtc": dateUtc } : {}),
			}),
		);
		findings.push({
			part: COMMENTS_EXTENSIBLE_PART_NAME,
			severity: "silent",
			message: `c${commentIdByKey.get(paraKey)} has no entry in ${COMMENTS_EXTENSIBLE_PART_NAME}`,
			fix: `added c${commentIdByKey.get(paraKey)}'s entry (durableId ${durableId})`,
		});
	}
}

/** Sidecar keys are hex — compare case-insensitively, or a lowercase key a
 * producer wrote would be pruned + re-added, losing its thread/resolved data. */
function key(value: string | undefined): string {
	return (value ?? "").toUpperCase();
}

function durableIdsIn(root: XmlNode, extra: Set<string>): Set<string> {
	const taken = new Set(extra);
	for (const child of root.findChildren("w16cid:commentId")) {
		const durableId = child.getAttribute("w16cid:durableId");
		if (durableId) taken.add(key(durableId));
	}
	return taken;
}

function missingEntryLead(comment: KeyedComment, partName: string): string {
	return comment.hadParaId
		? `comment c${comment.id} (paraId ${comment.paraId}) has no entry in ${partName}`
		: `comment c${comment.id} has no w14:paraId, so it has no entry in ${partName}`;
}

function staleEntry(
	partName: string,
	paraId: string | undefined,
): CommentFinding {
	return {
		part: partName,
		severity: "silent",
		message: `${partName} entry for paraId ${paraId ?? "(none)"} names no comment`,
		fix: `pruned the entry for paraId ${paraId ?? "(none)"} (no such comment)`,
	};
}

/** A comment's threading identity is its LAST `<w:p>` — that's the paragraph
 * Word keys `<w15:commentEx>` and a reply's `w15:paraIdParent` to. */
export function lastParagraph(
	comment: XmlNode | undefined,
): XmlNode | undefined {
	return comment?.findChildren("w:p").at(-1);
}

/** Word's `w16cex:dateUtc` is a whole-second UTC `xsd:dateTime`. Undefined
 * (attribute omitted — it's optional) when `date` doesn't parse. */
export function toDateUtc(date: string): string | undefined {
	if (!date) return undefined;
	const parsed = new Date(date);
	if (Number.isNaN(parsed.getTime())) return undefined;
	return parsed.toISOString().replace(/\.\d{3}Z$/, "Z");
}

export type SidecarRoots = {
	comments: XmlNode;
	extended?: XmlNode;
	ids?: XmlNode;
	extensible?: XmlNode;
};

/** One cross-part comment inconsistency, as `docx validate` reports it and
 *  `validate --fix` / reconcile-on-save cures it. `severity` follows Word's
 *  observed behavior: `error` — Word refuses to open the file; `warning` —
 *  Word opens it but repairs or silently drops something; `silent` — Word
 *  handles it harmlessly, so `validate` doesn't report it (the cure still
 *  runs). `fix` is what the cure does (present tense for a plan, the same
 *  text reads as the done-list after `--fix`); absent when no deterministic
 *  cure exists, and then `hint` may say what a person can do. */
export type CommentFinding = {
	part: string;
	severity: "error" | "warning" | "silent";
	message: string;
	fix?: string;
	hint?: string;
};

type KeyedComment = {
	id: string;
	paraId: string;
	hadParaId: boolean;
	date: string;
};
