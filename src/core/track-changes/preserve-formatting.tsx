import { w } from "../jsx";
import { isRunBearingWrapper, XmlNode } from "../parser";
import { Del, Ins } from "./emit";
import type { TrackedMeta } from "./index";

/**
 * Word-level diff for `edit --text` formatting preservation.
 *
 * The flow: read each existing `<w:r>` as a (text, rPr) segment, tokenize
 * old and new text on word/whitespace boundaries, run an LCS-based diff to
 * align new tokens against old ones, then re-emit runs grouped by rPr.
 *
 * Untracked path: emit only kept + inserted tokens (in NEW order). Inserted
 * tokens inherit rPr from the nearest matched neighbor.
 *
 * Tracked path: emit kept tokens as plain runs, inserted tokens wrapped in
 * `<w:ins>`, deleted tokens wrapped in `<w:del>`. The result is the same
 * shape Word produces when an author edits a few words in a paragraph
 * mid-tracking — much closer to the agent's intent than today's
 * whole-paragraph del+ins.
 *
 * Limitations: rPr equality is structural (XML-string compare on cloned
 * nodes). Two equivalent rPr blocks that serialize differently (e.g. child
 * order) won't merge into a single run. In practice the project's emitter
 * produces canonical output, so this is fine.
 */

export type OldToken = { text: string; rPr: XmlNode | null };

/** Read the paragraph's accepted-view text as a list of word/whitespace
 *  tokens, each tagged with the rPr active at that token's start.
 *
 *  Two reasons this is more involved than a flat `<w:r>` walk:
 *
 *   1. Source docs frequently split words across `<w:r>` boundaries
 *      (e.g. "me" + "ssenger" in adjacent italic runs). Per-segment
 *      tokenization would shred those words into sub-tokens that can
 *      never align with the new text's clean tokens — every
 *      cross-boundary word would round-trip as a spurious del+ins pair.
 *
 *   2. Paragraphs in real editorial workflows often contain run-bearing
 *      wrappers from prior tracked edits or markup: `<w:ins>`,
 *      `<w:moveTo>` (visible in accepted view), `<w:hyperlink>`,
 *      `<w:fldSimple>`, `<w:smartTag>` (transparent). Their inner runs
 *      contribute to the visible text; the diff must see them. `<w:del>`
 *      and `<w:moveFrom>` are invisible in accepted view — skip those.
 *
 *  Approach: walk recursively, descending through visible-in-accepted-view
 *  wrappers. Concatenate every visible run's text once, build a per-
 *  character rPr index, then tokenize the full string and look up rPr at
 *  each token's first character. Lossy in the (rare) case of mid-word
 *  formatting transitions — agents who need that should use `--runs` JSON.
 *
 *  Caller (`applyFormattingPreservingEdit`) must strip all run-bearing
 *  wrappers from the rebuild — their content was just flattened into the
 *  diff. */
export function extractOldTokens(paragraph: XmlNode): OldToken[] {
	let fullText = "";
	const charToRpr: Array<XmlNode | null> = [];
	function walk(node: XmlNode): void {
		for (const child of node.children) {
			if (child.tag === "w:r") {
				const segText = collectRunText(child);
				if (segText.length === 0) continue;
				const rPr = child.findChild("w:rPr") ?? null;
				fullText += segText;
				for (let index = 0; index < segText.length; index++) {
					charToRpr.push(rPr);
				}
				continue;
			}
			if (isAcceptedViewVisibleWrapper(child.tag)) {
				walk(child);
			}
			// else: <w:pPr>, bookmarks, comment markers, <w:del>/<w:moveFrom>
			// — skipped. Tracked-deleted text is invisible in accepted view.
		}
	}
	walk(paragraph);
	const tokenTexts = tokenize(fullText);
	const tokens: OldToken[] = [];
	let offset = 0;
	for (const text of tokenTexts) {
		const rPrSource = charToRpr[offset] ?? null;
		tokens.push({
			text,
			rPr: rPrSource ? rPrSource.clone() : null,
		});
		offset += text.length;
	}
	return tokens;
}

/** A wrapper whose content is visible in the accepted view: tracked
 *  insertions and move destinations (`<w:ins>`, `<w:moveTo>`) plus the
 *  view-neutral wrappers (`<w:hyperlink>`, `<w:fldSimple>`,
 *  `<w:smartTag>`). `<w:del>` and `<w:moveFrom>` are NOT visible. */
function isAcceptedViewVisibleWrapper(tag: string): boolean {
	if (!isRunBearingWrapper(tag)) return false;
	return tag !== "w:del" && tag !== "w:moveFrom";
}

function collectRunText(run: XmlNode): string {
	let out = "";
	for (const child of run.children) {
		if (child.tag === "w:t") out += child.collectText();
	}
	return out;
}

/** Tokenize on whitespace boundaries: each token is either a run of
 *  non-whitespace ("words") or a run of whitespace. Empty strings drop. */
export function tokenize(text: string): string[] {
	if (text.length === 0) return [];
	return text.match(/\S+|\s+/g) ?? [];
}

export type DiffOp =
	| { kind: "keep"; old: OldToken }
	| { kind: "insert"; text: string }
	| { kind: "delete"; old: OldToken };

/** LCS-based token diff. Returns a sequence of operations describing how
 *  to transform `oldTokens` into `newTokens`. `keep` ops carry the old
 *  token (for rPr); `insert` ops carry the new token's text only.
 *
 *  Two post-processing passes run after the raw LCS:
 *
 *   1. Demote orphan whitespace keeps. The LCS finds matches on text
 *      alone, so a single space token in the old text matches a single
 *      space in the new text — even when the surrounding words are
 *      completely different. Heading rewrites ("Police Reform / Safer
 *      Streets" → "Law and Order") get fragmented into alternating
 *      del/ins/keep-space runs that defeat run-merging. The fix:
 *      whitespace-only keeps with at least one non-keep neighbor get
 *      demoted to delete + insert.
 *
 *   2. Consolidate edit groups. Within each maximal run of non-keep
 *      ops, reorder so all deletes precede all inserts (preserving
 *      relative order within each kind). This lets `groupAndEmit*`
 *      merge consecutive same-rPr deletes into a single `<w:del>`
 *      wrapper (likewise for inserts), instead of emitting fragmented
 *      alternating wrappers. Position-pairing is unchanged — it indexes
 *      by ordinal within the group, which is preserved. */
export function diffTokens(
	oldTokens: OldToken[],
	newTokens: string[],
): DiffOp[] {
	const m = oldTokens.length;
	const n = newTokens.length;
	const dp: number[][] = Array.from({ length: m + 1 }, () =>
		new Array<number>(n + 1).fill(0),
	);
	for (let i = 1; i <= m; i++) {
		for (let j = 1; j <= n; j++) {
			const oldText = oldTokens[i - 1]?.text;
			const newText = newTokens[j - 1];
			const left = dp[i - 1] ?? [];
			const here = dp[i] ?? [];
			if (oldText === newText) {
				here[j] = (left[j - 1] ?? 0) + 1;
			} else {
				here[j] = Math.max(left[j] ?? 0, here[j - 1] ?? 0);
			}
		}
	}
	const ops: DiffOp[] = [];
	let i = m;
	let j = n;
	while (i > 0 || j > 0) {
		const oldToken = i > 0 ? oldTokens[i - 1] : undefined;
		const newText = j > 0 ? newTokens[j - 1] : undefined;
		if (i > 0 && j > 0 && oldToken !== undefined && oldToken.text === newText) {
			ops.unshift({ kind: "keep", old: oldToken });
			i--;
			j--;
			continue;
		}
		const upScore = i > 0 ? (dp[i - 1]?.[j] ?? 0) : -1;
		const leftScore = j > 0 ? (dp[i]?.[j - 1] ?? 0) : -1;
		if (j > 0 && newText !== undefined && (i === 0 || leftScore >= upScore)) {
			ops.unshift({ kind: "insert", text: newText });
			j--;
			continue;
		}
		if (i > 0 && oldToken !== undefined) {
			ops.unshift({ kind: "delete", old: oldToken });
			i--;
			continue;
		}
		// Defensive: shouldn't happen with the conditions above.
		break;
	}
	return consolidateEditGroups(demoteOrphanWhitespaceKeeps(ops));
}

/** Whitespace-only keeps surrounded by non-keep neighbors get demoted to
 *  delete + insert. See `diffTokens`'s docstring for the rationale. */
function demoteOrphanWhitespaceKeeps(ops: DiffOp[]): DiffOp[] {
	const out: DiffOp[] = [];
	for (let index = 0; index < ops.length; index++) {
		const op = ops[index];
		if (!op) continue;
		if (op.kind !== "keep") {
			out.push(op);
			continue;
		}
		if (!isWhitespaceOnly(op.old.text)) {
			out.push(op);
			continue;
		}
		const prev = ops[index - 1];
		const next = ops[index + 1];
		const prevIsEdit = prev !== undefined && prev.kind !== "keep";
		const nextIsEdit = next !== undefined && next.kind !== "keep";
		if (!prevIsEdit && !nextIsEdit) {
			out.push(op);
			continue;
		}
		out.push({ kind: "delete", old: op.old });
		out.push({ kind: "insert", text: op.old.text });
	}
	return out;
}

/** Within each maximal run of non-keep ops, reorder so all deletes come
 *  first, then all inserts. Preserves relative order within each kind so
 *  position-pairing in `inheritedRprForInsert` still aligns. */
function consolidateEditGroups(ops: DiffOp[]): DiffOp[] {
	const out: DiffOp[] = [];
	let cursor = 0;
	while (cursor < ops.length) {
		const op = ops[cursor];
		if (!op) {
			cursor++;
			continue;
		}
		if (op.kind === "keep") {
			out.push(op);
			cursor++;
			continue;
		}
		const groupEnd = (() => {
			let end = cursor;
			while (end < ops.length) {
				const candidate = ops[end];
				if (!candidate || candidate.kind === "keep") break;
				end++;
			}
			return end;
		})();
		const deletes: DiffOp[] = [];
		const inserts: DiffOp[] = [];
		for (let index = cursor; index < groupEnd; index++) {
			const groupOp = ops[index];
			if (!groupOp) continue;
			if (groupOp.kind === "delete") deletes.push(groupOp);
			else if (groupOp.kind === "insert") inserts.push(groupOp);
		}
		out.push(...deletes, ...inserts);
		cursor = groupEnd;
	}
	return out;
}

function isWhitespaceOnly(text: string): boolean {
	return text.length > 0 && /^\s+$/.test(text);
}

/** For a token-level edit, build the paragraph children for an UNTRACKED
 *  edit: only kept + inserted tokens, in NEW order, grouped into runs by
 *  shared rPr. Inserted tokens inherit rPr per the rules in
 *  `inheritedRprForInsert`. */
export function buildUntrackedRuns(ops: DiffOp[]): XmlNode[] {
	const flow: Array<{ text: string; rPr: XmlNode | null }> = [];
	for (let index = 0; index < ops.length; index++) {
		const op = ops[index];
		if (!op) continue;
		if (op.kind === "delete") continue;
		if (op.kind === "keep") {
			flow.push({ text: op.old.text, rPr: cloneOrNull(op.old.rPr) });
			continue;
		}
		const inherited = inheritedRprForInsert(ops, index);
		flow.push({ text: op.text, rPr: cloneOrNull(inherited) });
	}
	return groupAndEmitPlainRuns(flow);
}

/** When inserting a token in NEW order, inherit rPr by position-pairing
 *  within the surrounding "edit group" (the maximal run of insert/delete
 *  ops uninterrupted by any keep). The agent's mental model when replacing
 *  text is "swap word K of the deleted phrase for word K of the new
 *  phrase," so we pair them positionally:
 *
 *   1. Find the edit group spanning this insert.
 *   2. Index the inserts and deletes in that group in document order.
 *   3. If this insert is the K-th, use the K-th delete's rPr. If there
 *      are fewer deletes than inserts, repeat the last delete's rPr
 *      (the new "tail" inherits from the deleted tail).
 *   4. If the group has no deletes (pure insertion), inherit from the
 *      adjacent keep — backward first, then forward.
 *   5. Null when neither side has a keep (entire paragraph is insertion). */
function inheritedRprForInsert(ops: DiffOp[], index: number): XmlNode | null {
	let groupStart = index;
	while (groupStart > 0) {
		const prev = ops[groupStart - 1];
		if (!prev || prev.kind === "keep") break;
		groupStart--;
	}
	let groupEnd = index;
	while (groupEnd < ops.length - 1) {
		const next = ops[groupEnd + 1];
		if (!next || next.kind === "keep") break;
		groupEnd++;
	}

	const deletes: Array<XmlNode | null> = [];
	let insertOrdinal = 0;
	let myInsertOrdinal = -1;
	for (let cursor = groupStart; cursor <= groupEnd; cursor++) {
		const op = ops[cursor];
		if (!op) continue;
		if (op.kind === "delete") {
			deletes.push(op.old.rPr);
			continue;
		}
		if (op.kind === "insert") {
			if (cursor === index) myInsertOrdinal = insertOrdinal;
			insertOrdinal++;
		}
	}

	if (deletes.length > 0 && myInsertOrdinal >= 0) {
		const paired =
			myInsertOrdinal < deletes.length
				? deletes[myInsertOrdinal]
				: deletes[deletes.length - 1];
		if (paired !== undefined) return paired;
	}

	// No deletes in this edit group — inherit from the adjacent kept
	// neighbor (backward preferred, then forward).
	for (let cursor = groupStart - 1; cursor >= 0; cursor--) {
		const op = ops[cursor];
		if (op?.kind === "keep") return op.old.rPr;
	}
	for (let cursor = groupEnd + 1; cursor < ops.length; cursor++) {
		const op = ops[cursor];
		if (op?.kind === "keep") return op.old.rPr;
	}
	return null;
}

/** For a token-level edit, build the paragraph children for a TRACKED
 *  edit: kept tokens emit as plain runs with their original rPr; inserted
 *  tokens wrap in `<w:ins>` (rPr inherited from neighbors); deleted tokens
 *  wrap in `<w:del>` (with their original rPr; `<w:t>` becomes
 *  `<w:delText>`). The order follows the diff ops. */
export function buildTrackedRuns(
	ops: DiffOp[],
	mintMeta: () => TrackedMeta,
): XmlNode[] {
	const grouping: Array<{
		kind: "keep" | "insert" | "delete";
		flow: Array<{ text: string; rPr: XmlNode | null }>;
	}> = [];
	for (let index = 0; index < ops.length; index++) {
		const op = ops[index];
		if (!op) continue;
		if (op.kind === "keep") {
			pushOpTo(grouping, "keep", {
				text: op.old.text,
				rPr: cloneOrNull(op.old.rPr),
			});
			continue;
		}
		if (op.kind === "insert") {
			const inherited = inheritedRprForInsert(ops, index);
			pushOpTo(grouping, "insert", {
				text: op.text,
				rPr: cloneOrNull(inherited),
			});
			continue;
		}
		// delete: keeps the original rPr
		pushOpTo(grouping, "delete", {
			text: op.old.text,
			rPr: cloneOrNull(op.old.rPr),
		});
	}

	const out: XmlNode[] = [];
	for (const group of grouping) {
		const runs =
			group.kind === "delete"
				? groupAndEmitDeletedRuns(group.flow)
				: groupAndEmitPlainRuns(group.flow);
		if (runs.length === 0) continue;
		if (group.kind === "keep") {
			out.push(...runs);
		} else if (group.kind === "insert") {
			out.push(<Ins meta={mintMeta()}>{runs}</Ins>);
		} else {
			out.push(<Del meta={mintMeta()}>{runs}</Del>);
		}
	}
	return out;
}

function pushOpTo(
	grouping: Array<{
		kind: "keep" | "insert" | "delete";
		flow: Array<{ text: string; rPr: XmlNode | null }>;
	}>,
	kind: "keep" | "insert" | "delete",
	entry: { text: string; rPr: XmlNode | null },
): void {
	const last = grouping[grouping.length - 1];
	if (last && last.kind === kind) {
		last.flow.push(entry);
		return;
	}
	grouping.push({ kind, flow: [entry] });
}

/** Group consecutive flow entries with the same rPr into a single
 *  `<w:r>` containing one `<w:t xml:space="preserve">` child whose text is
 *  the concatenation. Whitespace-only segments collapse with their
 *  neighbors when rPr matches. */
function groupAndEmitPlainRuns(
	flow: Array<{ text: string; rPr: XmlNode | null }>,
): XmlNode[] {
	const out: XmlNode[] = [];
	let current: { text: string; rPr: XmlNode | null } | null = null;
	for (const entry of flow) {
		if (entry.text.length === 0) continue;
		if (current && rPrEqual(current.rPr, entry.rPr)) {
			current.text += entry.text;
			continue;
		}
		if (current) out.push(emitPlainRun(current.text, current.rPr));
		current = { text: entry.text, rPr: entry.rPr };
	}
	if (current) out.push(emitPlainRun(current.text, current.rPr));
	return out;
}

function groupAndEmitDeletedRuns(
	flow: Array<{ text: string; rPr: XmlNode | null }>,
): XmlNode[] {
	const out: XmlNode[] = [];
	let current: { text: string; rPr: XmlNode | null } | null = null;
	for (const entry of flow) {
		if (entry.text.length === 0) continue;
		if (current && rPrEqual(current.rPr, entry.rPr)) {
			current.text += entry.text;
			continue;
		}
		if (current) out.push(emitDeletedRun(current.text, current.rPr));
		current = { text: entry.text, rPr: entry.rPr };
	}
	if (current) out.push(emitDeletedRun(current.text, current.rPr));
	return out;
}

function emitPlainRun(text: string, rPr: XmlNode | null): XmlNode {
	return (
		<w.r>
			{rPr}
			<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
		</w.r>
	) as XmlNode;
}

function emitDeletedRun(text: string, rPr: XmlNode | null): XmlNode {
	return (
		<w.r>
			{rPr}
			<w.delText {...{ "xml:space": "preserve" }}>{text}</w.delText>
		</w.r>
	) as XmlNode;
}

function cloneOrNull(node: XmlNode | null): XmlNode | null {
	return node ? node.clone() : null;
}

function rPrEqual(a: XmlNode | null, b: XmlNode | null): boolean {
	if (a === null && b === null) return true;
	if (a === null || b === null) return false;
	return XmlNode.serialize([a]) === XmlNode.serialize([b]);
}
