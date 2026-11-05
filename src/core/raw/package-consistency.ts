import {
	auditThreadLinks,
	type CommentFinding,
	reconcileSidecars,
	type SidecarRoots,
	threadRootsOf,
} from "../ast/document/comments/reconcile-sidecars";
import { auditCommentMarkers, type StoryPart } from "../comments/marker-audit";
import type { XmlNode } from "../parser";

/** Package-level consistency audit — the cross-part invariants no per-part
 *  schema can see, run by `docx validate` beside the XSD check. Every finding
 *  is something a schema-valid package can still get wrong, graded by what
 *  live Word (16.114) does with it:
 *
 *  - **Comment markers** across every story part (body, headers, footers,
 *    notes) — `auditCommentMarkers`: a `commentReference` naming no comment
 *    is the one ERROR (Word refuses to open the file); half-open ranges,
 *    reference-less ranges, stray ranges and orphan comments are warnings.
 *  - **Comment side-parts** — `reconcileSidecars` run on a CLONE: every
 *    comment keyed (by its LAST paragraph's `w14:paraId`) in each PRESENT
 *    `commentsExtended.xml`/`commentsIds.xml` (an absent one is legal). Word
 *    re-adds a missing entry on open, so it's a warning — worded as data loss
 *    for commentsExtended, where the re-added entry is open and top-level.
 *    Stale entries and duplicate durableIds are cured but not reported (Word
 *    handles them harmlessly).
 *  - **Thread links** — `auditThreadLinks`: a `w15:paraIdParent` naming no
 *    comment (warning; Word drops the thread).
 *
 *  Because the reports come from the very functions `validate --fix` (and
 *  every save) applies, a finding is marked fixable exactly when the fix cures
 *  it. The glossary document (`word/glossary/…`, building blocks) is a
 *  separate story set with its own comment parts, so it's excluded rather
 *  than mixed into the main document's id spaces. */
export function checkPackageConsistency(
	packageParts: PackagePart[],
): ConsistencyIssue[] {
	const parts = packageParts.filter(
		(part) => !part.name.startsWith("word/glossary/"),
	);
	const commentsRoot = rootTagged(parts, "w:comments");
	const commentIds = new Set(
		(commentsRoot?.findChildren("w:comment") ?? []).flatMap((comment) => {
			const id = comment.getAttribute("w:id");
			return id === undefined ? [] : [id];
		}),
	);
	const stories: StoryPart[] = parts.filter((part) =>
		STORY_ROOTS.has(part.root.tag),
	);
	const roots: SidecarRoots | undefined = commentsRoot && {
		comments: commentsRoot,
		extended: rootTagged(parts, "w15:commentsEx"),
		ids: rootTagged(parts, "w16cid:commentsIds"),
		extensible: rootTagged(parts, "w16cex:commentsExtensible"),
	};
	const findings: CommentFinding[] = auditCommentMarkers(stories, commentIds, {
		repair: false,
		threadRoots: roots ? threadRootsOf(roots) : undefined,
	});
	if (roots) {
		findings.push(...reconcileSidecars(cloneRoots(roots)));
		findings.push(...auditThreadLinks(roots));
	}
	return findings.flatMap((finding) =>
		finding.severity === "silent"
			? []
			: [
					{
						part: finding.part,
						message: finding.message,
						severity: finding.severity,
						fixable: finding.fix !== undefined,
						...(finding.hint ? { hint: finding.hint } : {}),
					},
				],
	);
}

const STORY_ROOTS = new Set([
	"w:document",
	"w:hdr",
	"w:ftr",
	"w:footnotes",
	"w:endnotes",
]);

function rootTagged(parts: PackagePart[], tag: string): XmlNode | undefined {
	return parts.find((part) => part.root.tag === tag)?.root;
}

/** The reconcile mutates in place; reporting must not. */
function cloneRoots(roots: SidecarRoots): SidecarRoots {
	return {
		comments: roots.comments.clone(),
		...(roots.extended ? { extended: roots.extended.clone() } : {}),
		...(roots.ids ? { ids: roots.ids.clone() } : {}),
		...(roots.extensible ? { extensible: roots.extensible.clone() } : {}),
	};
}

export type PackagePart = { name: string; root: XmlNode };

export type ConsistencyIssue = {
	part: string;
	message: string;
	severity: "error" | "warning";
	fixable: boolean;
	hint?: string;
};
