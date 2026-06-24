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
 *  character rPr index, then tokenize the full string and tag each token with
 *  the rPr covering the MOST of its characters (`dominantRpr`). Using the bulk
 *  formatting (not the first character) keeps a token that glues differently-
 *  formatted characters — e.g. a plain `[` bracket fused to an underlined word in
 *  a fill-in-the-blank placeholder — from inheriting a stray edge char's format
 *  and coming out ragged. Still lossy for genuine mid-word format changes (a word
 *  half bold) — agents who need that should use `--runs` JSON.
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
		const rPrSource = dominantRpr(charToRpr, offset, text.length);
		tokens.push({
			text,
			rPr: rPrSource ? rPrSource.clone() : null,
		});
		offset += text.length;
	}
	return tokens;
}

/** The rPr covering the MOST characters of a token (ties → first seen), not just
 *  its first character. A token can glue characters from differently-formatted
 *  runs — e.g. a fill-in-the-blank placeholder `[Evaluating …]` where the `[`
 *  bracket is plain but the word is underlined, so the token `[Evaluating` spans
 *  a plain char and an underlined run. First-character lookup made that token
 *  inherit the bracket's (plain) rPr while its underlined neighbors inherited
 *  underline — so a whole-paragraph `--text` replace came out RAGGED (first word
 *  un-formatted, the rest formatted). Taking the bulk formatting fixes that.
 *
 *  Buckets by STRUCTURAL rPr (serialized), NOT by node reference: `extractOldTokens`
 *  glues a word split across several `<w:r>` (the "me"+"ssenger" case), and those
 *  sub-runs are SEPARATE XmlNodes even when format-identical. Reference-keyed
 *  counting would split identical underlined sub-runs into separate buckets and let
 *  one longer plain neighbor out-vote them, dropping the underline (the bug this
 *  function exists to fix). Structural keying counts all format-identical chars
 *  together. */
function dominantRpr(
	charToRpr: ReadonlyArray<XmlNode | null>,
	start: number,
	length: number,
): XmlNode | null {
	// key = serialized rPr ("" for plain/null); value keeps a first-seen
	// representative node + its running count. Insertion order is char order, so a
	// tie resolves to the first-seen formatting.
	const counts = new Map<string, { node: XmlNode | null; count: number }>();
	for (let index = start; index < start + length; index++) {
		const rPr = charToRpr[index] ?? null;
		const key = rPr ? XmlNode.serialize([rPr]) : "";
		const entry = counts.get(key);
		if (entry) entry.count += 1;
		else counts.set(key, { node: rPr, count: 1 });
	}
	let dominant: XmlNode | null = null;
	let best = -1;
	for (const { node, count } of counts.values()) {
		if (count > best) {
			best = count;
			dominant = node;
		}
	}
	return dominant;
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
		if (child.tag === "w:t") {
			out += child.collectText();
			continue;
		}
		// Represent in-run tabs/line-breaks as the same literal characters the new
		// text carries, so old and new tokenize identically. Otherwise a `<w:tab/>`
		// in its own run glues the surrounding words into one old token
		// ("University" + "Cambridge,") that can never match the new split tokens —
		// turning a verbatim re-type of a tabbed line into a spurious del+reinsert
		// under tracking. `runContentChildren` re-emits these from the same chars.
		if (child.tag === "w:tab") {
			out += "\t";
			continue;
		}
		// Only the default text-wrapping break is a newline; page/column breaks
		// aren't text the agent retypes, so leave them out (and out of the diff).
		if (child.tag === "w:br") {
			const type = child.getAttribute("w:type");
			if (type === undefined || type === "textWrapping") out += "\n";
		}
	}
	return out;
}

/** Tokenize on whitespace boundaries: each token is a run of non-whitespace
 *  ("words"), a run of TABS, or a run of non-tab whitespace (spaces/newlines).
 *  Empty strings drop.
 *
 *  Tabs get their OWN token class — never glued to adjacent spaces — so the tab
 *  ALWAYS aligns as a keep across old↔new. A template like `Title<tab><bold
 *  space>Month` reads as the old token `"\t "` under a naive `\s+`; the new
 *  `Title<tab>Date` has just `"\t"`, the two never match, the tab stops being a
 *  recognized boundary, and the last word before the tab (`Intern`, `Lead`) ends
 *  up positionally paired across the boundary and loses its bold. Splitting the
 *  tab out keeps `demoteOrphanWhitespaceKeeps`'s tab-boundary protection working
 *  and each tab-delimited segment formatted on its own side. */
export function tokenize(text: string): string[] {
	if (text.length === 0) return [];
	return text.match(/\S+|\t+|[^\S\t]+/g) ?? [];
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
		// A TAB keep is a structural/format boundary (left text | right-aligned
		// content), not a spuriously-matched space. Demoting it merges the regions
		// on either side into ONE edit group, so positional rPr-pairing bleeds the
		// left region's formatting across the tab — e.g. a bold org name → a bold
		// city ("Lincoln High School⇥Portland, OR" turning "Portland," bold). Keep
		// it so each tab-delimited segment keeps its own formatting.
		if (op.old.text.includes("\t")) {
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
export function buildUntrackedRuns(
	ops: DiffOp[],
	fallbackRpr: XmlNode | null = null,
): XmlNode[] {
	const flow: Array<{ text: string; rPr: XmlNode | null }> = [];
	for (let index = 0; index < ops.length; index++) {
		const op = ops[index];
		if (!op) continue;
		if (op.kind === "delete") continue;
		if (op.kind === "keep") {
			flow.push({ text: op.old.text, rPr: cloneOrNull(op.old.rPr) });
			continue;
		}
		const inherited = inheritedRprForInsert(ops, index, fallbackRpr);
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
function inheritedRprForInsert(
	ops: DiffOp[],
	index: number,
	fallbackRpr: XmlNode | null,
): XmlNode | null {
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

	// Two pairing tracks. A WORD insert pairs only against WORD deletes, so a real
	// word never inherits a stray whitespace token's rPr — the bold-space-after-tab
	// case (`Title<tab><bold space>Date` → new `Title<tab>Date`) where the first new
	// word after the tab ("Jun") would otherwise grab the deleted bold space's bold.
	// A WHITESPACE insert keeps the full pairing, so a space INSIDE a bold phrase
	// still inherits the surrounding bold (keeps `**A B C**` contiguous).
	const targetText = (() => {
		const op = ops[index];
		return op && op.kind === "insert" ? op.text : "";
	})();
	const targetIsWord = !isWhitespaceOnly(targetText);

	const allDeletes: Array<XmlNode | null> = [];
	const wordDeletes: Array<XmlNode | null> = [];
	let allInsertOrdinal = 0;
	let wordInsertOrdinal = 0;
	let myAllOrdinal = -1;
	let myWordOrdinal = -1;
	for (let cursor = groupStart; cursor <= groupEnd; cursor++) {
		const op = ops[cursor];
		if (!op) continue;
		if (op.kind === "delete") {
			allDeletes.push(op.old.rPr);
			if (!isWhitespaceOnly(op.old.text)) wordDeletes.push(op.old.rPr);
			continue;
		}
		if (op.kind === "insert") {
			const isWord = !isWhitespaceOnly(op.text);
			if (cursor === index) {
				myAllOrdinal = allInsertOrdinal;
				myWordOrdinal = isWord ? wordInsertOrdinal : -1;
			}
			allInsertOrdinal++;
			if (isWord) wordInsertOrdinal++;
		}
	}

	if (targetIsWord && wordDeletes.length > 0 && myWordOrdinal >= 0) {
		const paired =
			myWordOrdinal < wordDeletes.length
				? wordDeletes[myWordOrdinal]
				: wordDeletes[wordDeletes.length - 1];
		if (paired !== undefined) return paired;
	}
	if (!targetIsWord && allDeletes.length > 0 && myAllOrdinal >= 0) {
		const paired =
			myAllOrdinal < allDeletes.length
				? allDeletes[myAllOrdinal]
				: allDeletes[allDeletes.length - 1];
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
	// Nothing run-level to inherit from (the common case: filling an empty
	// paragraph/cell). Fall back to the paragraph-mark rPr — the formatting Word
	// applies when you type into an empty styled cell. Null when even that is
	// absent (truly unformatted paragraph).
	return fallbackRpr;
}

/** For a token-level edit, build the paragraph children for a TRACKED
 *  edit: kept tokens emit as plain runs with their original rPr; inserted
 *  tokens wrap in `<w:ins>` (rPr inherited from neighbors); deleted tokens
 *  wrap in `<w:del>` (with their original rPr; `<w:t>` becomes
 *  `<w:delText>`). The order follows the diff ops. */
export function buildTrackedRuns(
	ops: DiffOp[],
	mintMeta: () => TrackedMeta,
	fallbackRpr: XmlNode | null = null,
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
			const inherited = inheritedRprForInsert(ops, index, fallbackRpr);
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
	return <w.r>{[rPr, ...runContentChildren(text, false)]}</w.r>;
}

function emitDeletedRun(text: string, rPr: XmlNode | null): XmlNode {
	return <w.r>{[rPr, ...runContentChildren(text, true)]}</w.r>;
}

/** Split a token's text on tabs/newlines into `<w:t>`/`<w:delText>` segments
 *  plus `<w:tab/>` / `<w:br/>` elements, all destined for ONE `<w:r>` so the
 *  diff's per-token rPr (e.g. bold on a name, regular on the tabbed-over date)
 *  rides along. Without this, a tab in `--text` would land as a literal char
 *  inside `<w:t>` (Word ignores it) — which is why `--text` with a tab used to
 *  bypass this whole formatting-preserving path and flatten the run. Mirrors
 *  `textToRuns` in blocks.tsx, but keeps everything in the rPr-bearing run. */
function runContentChildren(text: string, deleted: boolean): XmlNode[] {
	if (!/[\n\r\t]/.test(text)) {
		return [textSegment(text, deleted)];
	}
	const out: XmlNode[] = [];
	for (const segment of text.split(/(\r\n|\r|\n|\t)/)) {
		if (segment === "") continue;
		if (segment === "\t") out.push(<w.tab />);
		else if (segment === "\n" || segment === "\r" || segment === "\r\n") {
			out.push(<w.br />);
		} else out.push(textSegment(segment, deleted));
	}
	return out;
}

function textSegment(text: string, deleted: boolean): XmlNode {
	return deleted ? (
		<w.delText {...{ "xml:space": "preserve" }}>{text}</w.delText>
	) : (
		<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
	);
}

function cloneOrNull(node: XmlNode | null): XmlNode | null {
	return node ? node.clone() : null;
}

/** The tracked-change / property-change markers that CT_ParaRPr permits at the
 *  front of a paragraph-mark `<w:rPr>` but that are INVALID inside a run's
 *  `<w:rPr>` (CT_RPr): all four revision markers plus the rPr-change snapshot.
 *  Strip every one when reusing the paragraph-mark rPr as a run rPr, or Word
 *  silently drops the whole (schema-invalid) run. */
const PARA_MARK_ONLY_RPR_CHILDREN = new Set([
	"w:ins",
	"w:del",
	"w:moveFrom",
	"w:moveTo",
	"w:rPrChange",
]);

/** The run formatting a paragraph declares on its paragraph mark
 *  (`<w:pPr><w:rPr>`), as a clean run-rPr: a clone with the paragraph-mark-only
 *  children stripped (the four `<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`
 *  revision markers and `<w:rPrChange>`, all invalid inside a `<w:r>`). Word
 *  uses this rPr for text typed into
 *  an otherwise-empty paragraph/cell, so it's the right fallback when there's no
 *  run-level neighbor to inherit from. Null when absent or empty after stripping. */
export function paragraphMarkRunRpr(paragraph: XmlNode): XmlNode | null {
	const rPr = paragraph.findChild("w:pPr")?.findChild("w:rPr");
	if (!rPr) return null;
	const clone = rPr.clone();
	clone.children = clone.children.filter(
		(child) => !PARA_MARK_ONLY_RPR_CHILDREN.has(child.tag),
	);
	if (clone.children.length === 0) return null;
	return clone;
}

function rPrEqual(a: XmlNode | null, b: XmlNode | null): boolean {
	if (a === null && b === null) return true;
	if (a === null || b === null) return false;
	return XmlNode.serialize([a]) === XmlNode.serialize([b]);
}
