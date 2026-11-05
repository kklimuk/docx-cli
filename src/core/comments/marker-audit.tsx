import {
	COMMENTS_PART_NAME,
	type CommentFinding,
} from "../ast/document/comments/reconcile-sidecars";
import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { addReplyCommentMarkers } from "./markers";

/** Audit the comment markers (`commentRangeStart`/`commentRangeEnd`/
 *  `commentReference`) across every story part against the comment ids that
 *  exist, and — with `repair: true` — apply the deterministic cure for each
 *  finding in place. ONE function for both, so `docx validate` (which runs it
 *  on a clone) reports a finding as fixable exactly when `validate --fix` will
 *  cure it. Severity follows what live Word (16.114) does:
 *
 *  - a `commentReference` to a comment that doesn't exist — the ONE state that
 *    makes Word refuse to open the file — is an `error`; cure: remove every
 *    marker of that id;
 *  - range markers of a nonexistent comment are stray (`warning`; removed);
 *  - a comment whose range is half-open (start without end, or end without
 *    start) but which has its reference is a `warning`; cure: what Word does on
 *    open — collapse the range to a point at the reference mark;
 *  - a comment with a range but no reference is a `warning` (Word has no mark
 *    to show it at); cure: close the range if needed and add the reference
 *    after it, when the range sits inside a paragraph;
 *  - a comment with no marker anywhere (an orphan) is a `warning` — Word
 *    silently deletes it on open. A REPLY orphan (docx-cli before 0.14.1 wrote
 *    replies without markers) whose thread root is anchored in the body is
 *    cured the way `comments reply` anchors one today: mirror the thread's
 *    markers (`addReplyCommentMarkers`). Any other orphan has NO cure — only a
 *    person can say where it belongs (or that it should go), and the audit
 *    never deletes a comment. `threadRoots` (reply id → thread-root id, from
 *    `threadRootsOf`) is what makes the reply case recognizable.
 *
 *  Ids compare as SETS per kind: `mc:AlternateContent` writes a text-box story
 *  under both Choice and Fallback, so every marker may legitimately appear
 *  twice — a cure applies to every occurrence. */
export function auditCommentMarkers(
	stories: StoryPart[],
	commentIds: Set<string>,
	options: { repair: boolean; threadRoots?: Map<string, string> },
): CommentFinding[] {
	const markers = collectMarkers(stories);
	const findings: CommentFinding[] = [];
	// Comments with a reference run in the body once the cures below apply —
	// the thread roots an orphan reply can be re-anchored under. Computed from
	// the PLANNED cures, so a report (repair: false) marks a reply fixable
	// exactly when one `--fix` pass anchors it.
	const referencedInBody = new Set<string>();
	for (const [id, found] of markers) {
		if (!commentIds.has(id)) {
			findings.push(auditMissingComment(id, found, options.repair));
			continue;
		}
		const finding = auditLiveComment(id, found, options.repair);
		if (finding) findings.push(finding);
		const bodyRun = found.reference.some(
			(occurrence) =>
				occurrence.part === BODY_PART_NAME && occurrence.parent.tag === "w:r",
		);
		const gainsBodyReference =
			found.reference.length === 0 &&
			finding?.fix !== undefined &&
			[...found.start, ...found.end].every(
				(occurrence) => occurrence.part === BODY_PART_NAME,
			);
		if (bodyRun || gainsBodyReference) referencedInBody.add(id);
	}
	const anchored = new Set(markers.keys());
	for (const id of commentIds) {
		if (markers.has(id)) continue;
		findings.push(
			auditOrphan(id, {
				stories,
				referencedInBody,
				anchored,
				commentIds,
				threadRoots: options.threadRoots ?? new Map(),
				repair: options.repair,
			}),
		);
	}
	return findings;
}

/** A comment with no marker anywhere. Fixable only as a reply whose thread
 * root has a reference run in the body: the reply's markers are mirrored
 * from the thread, after its already-anchored siblings (so replies cured in
 * one pass stay in comment order). */
function auditOrphan(
	id: string,
	context: {
		stories: StoryPart[];
		referencedInBody: Set<string>;
		anchored: Set<string>;
		commentIds: Set<string>;
		threadRoots: Map<string, string>;
		repair: boolean;
	},
): CommentFinding {
	const message = `comment c${id} has no anchor in the document (no commentRangeStart, commentRangeEnd or commentReference) — Word silently deletes it on open`;
	const rootId = context.threadRoots.get(id);
	const body = context.stories.find((story) => story.name === BODY_PART_NAME);
	const rootAnchoredInBody =
		rootId !== undefined &&
		rootId !== id &&
		context.referencedInBody.has(rootId);
	if (!rootId || !body || !rootAnchoredInBody) {
		return {
			part: COMMENTS_PART_NAME,
			severity: "warning",
			message,
			hint: `re-anchor its text with \`docx comments add\`, or remove it with \`docx comments delete FILE --at c${id}\``,
		};
	}
	if (context.repair) {
		const members = [
			rootId,
			...[...context.commentIds].filter(
				(other) =>
					other !== rootId &&
					context.threadRoots.get(other) === rootId &&
					context.anchored.has(other),
			),
		];
		addReplyCommentMarkers([body.root], members, id);
		context.anchored.add(id);
	}
	return {
		part: COMMENTS_PART_NAME,
		severity: "warning",
		message: `${message} (it is a reply in c${rootId}'s thread)`,
		fix: `anchored reply c${id} in c${rootId}'s thread (mirrored the thread's markers)`,
	};
}

const BODY_PART_NAME = "word/document.xml";

/** Every comment marker in the story parts, grouped by id in first-seen
 * (document) order. */
function collectMarkers(stories: StoryPart[]): Map<string, IdMarkers> {
	const markers = new Map<string, IdMarkers>();
	for (const story of stories) {
		walk(story.root, story.name, [], false, markers);
	}
	return markers;
}

function walk(
	node: XmlNode,
	part: string,
	ancestors: XmlNode[],
	inParagraph: boolean,
	markers: Map<string, IdMarkers>,
): void {
	for (const child of node.children) {
		if (child.isText) continue;
		const kind = MARKER_KINDS[child.tag];
		if (kind) {
			const id = child.getAttribute("w:id");
			if (id === undefined) continue;
			let entry = markers.get(id);
			if (!entry) {
				entry = { start: [], end: [], reference: [] };
				markers.set(id, entry);
			}
			entry[kind].push({
				node: child,
				parent: node,
				grandparent: ancestors.at(-1),
				part,
				inline: inParagraph && !node.tag.startsWith("m:"),
			});
			continue;
		}
		const childInParagraph =
			child.tag === "w:p"
				? true
				: BLOCK_CONTAINERS.has(child.tag)
					? false
					: inParagraph;
		walk(child, part, [...ancestors, node], childInParagraph, markers);
	}
}

/** Story containers whose children are BLOCKS — a marker directly inside one
 * (or inside a block-level `w:sdtContent`/`w:customXml` under one) is not in a
 * paragraph, so no reference run can be placed beside it. */
const BLOCK_CONTAINERS = new Set([
	"w:body",
	"w:tc",
	"w:txbxContent",
	"w:hdr",
	"w:ftr",
	"w:footnote",
	"w:endnote",
	"w:docPartBody",
]);

const MARKER_KINDS: Record<string, keyof IdMarkers> = {
	"w:commentRangeStart": "start",
	"w:commentRangeEnd": "end",
	"w:commentReference": "reference",
};

function auditMissingComment(
	id: string,
	found: IdMarkers,
	repair: boolean,
): CommentFinding {
	if (repair) {
		for (const occurrence of [...found.start, ...found.end]) {
			removeNode(occurrence.parent, occurrence.node);
		}
		for (const occurrence of found.reference) removeReference(occurrence);
	}
	const reference = found.reference[0];
	if (reference) {
		return {
			part: reference.part,
			severity: "error",
			message: `commentReference ${id} points at no comment in ${COMMENTS_PART_NAME} — Word refuses to open the file`,
			fix: `removed every marker of the missing comment ${id}`,
		};
	}
	return {
		part: (found.start[0] ?? found.end[0])?.part ?? COMMENTS_PART_NAME,
		severity: "warning",
		message: `comment range ${id} names no comment in ${COMMENTS_PART_NAME} (stray range markers)`,
		fix: `removed the stray range markers of ${id}`,
	};
}

function auditLiveComment(
	id: string,
	found: IdMarkers,
	repair: boolean,
): CommentFinding | undefined {
	const hasStart = found.start.length > 0;
	const hasEnd = found.end.length > 0;
	const firstReference = found.reference[0];
	if (firstReference) {
		if (hasStart === hasEnd) return undefined;
		if (repair) collapseOntoReference(id, found);
		const present = hasStart ? "commentRangeStart" : "commentRangeEnd";
		const missing = hasStart ? "commentRangeEnd" : "commentRangeStart";
		const occurrence =
			(hasStart ? found.start[0] : found.end[0]) ?? firstReference;
		return {
			part: occurrence.part,
			severity: "warning",
			message: `${present} ${id} has no matching ${missing} — Word collapses comment c${id} to its reference mark`,
			fix: `collapsed c${id}'s range onto its reference mark (as Word does)`,
		};
	}

	// A range but no reference mark.
	const anchors = hasEnd ? found.end : found.start;
	const first = anchors[0];
	if (!first) return undefined;
	const fixable = anchors.every((occurrence) => occurrence.inline);
	if (repair && fixable) closeRangeWithReference(id, found);
	return {
		part: first.part,
		severity: "warning",
		message: `comment range ${id} has no matching commentReference${hasStart && hasEnd ? "" : ` (and no ${hasStart ? "commentRangeEnd" : "commentRangeStart"})`} — Word has no mark to show comment c${id} at and drops it`,
		...(fixable
			? {
					fix: `added c${id}'s missing commentReference${hasStart && hasEnd ? "" : " and closed its range"}`,
				}
			: {
					hint: `the range sits outside any paragraph; re-anchor c${id} with \`docx comments add\``,
				}),
	};
}

/** Word's repair for a half-open range: drop the lone marker and give the
 * comment a zero-width range right before its reference run. */
function collapseOntoReference(id: string, found: IdMarkers): void {
	for (const occurrence of [...found.start, ...found.end]) {
		removeNode(occurrence.parent, occurrence.node);
	}
	for (const reference of found.reference) {
		const { container, anchor } = referenceRunSlot(reference);
		const index = container.children.indexOf(anchor);
		if (index === -1) continue;
		container.children.splice(
			index,
			0,
			<w.commentRangeStart w-id={id} />,
			<w.commentRangeEnd w-id={id} />,
		);
	}
}

/** A range with no reference: close it if half-open (a point at the lone
 * marker) and add the reference run right after its end. */
function closeRangeWithReference(id: string, found: IdMarkers): void {
	const anchors = found.end.length > 0 ? found.end : found.start;
	for (const anchor of anchors) {
		const children = anchor.parent.children;
		const index = children.indexOf(anchor.node);
		if (index === -1) continue;
		const inserted: XmlNode[] = [];
		if (found.start.length === 0) {
			children.splice(index, 0, <w.commentRangeStart w-id={id} />);
		}
		if (found.end.length === 0) inserted.push(<w.commentRangeEnd w-id={id} />);
		inserted.push(<CommentReferenceRun id={id} />);
		const at = children.indexOf(anchor.node) + 1;
		children.splice(at, 0, ...inserted);
	}
}

function CommentReferenceRun({ id }: { id: string }): XmlNode {
	return (
		<w.r>
			<w.rPr>
				<w.rStyle w-val="CommentReference" />
			</w.rPr>
			<w.commentReference w-id={id} />
		</w.r>
	);
}

/** The run carrying a reference (or the bare element, if a producer put it
 * outside a run) and the child list that run sits in. */
function referenceRunSlot(reference: MarkerOccurrence): {
	container: XmlNode;
	anchor: XmlNode;
} {
	if (reference.parent.tag === "w:r" && reference.grandparent) {
		return { container: reference.grandparent, anchor: reference.parent };
	}
	return { container: reference.parent, anchor: reference.node };
}

/** Remove a `<w:commentReference>`, and its run when nothing but run
 * properties is left in it. */
function removeReference(reference: MarkerOccurrence): void {
	removeNode(reference.parent, reference.node);
	const { container, anchor } = referenceRunSlot(reference);
	if (anchor !== reference.parent) return;
	const remaining = anchor.children.filter(
		(child) => !child.isText && child.tag !== "w:rPr",
	);
	if (remaining.length === 0) removeNode(container, anchor);
}

function removeNode(parent: XmlNode, node: XmlNode): void {
	const index = parent.children.indexOf(node);
	if (index !== -1) parent.children.splice(index, 1);
}

/** A story part (body, header/footer, notes) by package part name. */
export type StoryPart = { name: string; root: XmlNode };

type MarkerOccurrence = {
	node: XmlNode;
	parent: XmlNode;
	grandparent?: XmlNode;
	part: string;
	inline: boolean;
};

type IdMarkers = {
	start: MarkerOccurrence[];
	end: MarkerOccurrence[];
	reference: MarkerOccurrence[];
};
