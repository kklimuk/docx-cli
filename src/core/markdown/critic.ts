import type { Parent, Root, Text } from "mdast";
import type { Plugin } from "unified";
import { SKIP, visit } from "unist-util-visit";

/** CriticMarkup insertion `{++text++}` — represented as a phrasing-content
 * node we splice into the mdast tree, so the inline walker can wrap the value
 * in a `<w:ins>` (tracking on) or splat it as plain text (tracking off). */
export interface CriticInsert {
	type: "criticInsert";
	value: string;
}

/** CriticMarkup deletion `{--text--}` — mirror of {@link CriticInsert} for
 * `<w:del>`. With tracking off, the value is dropped entirely (it represents
 * removed content). */
export interface CriticDelete {
	type: "criticDelete";
	value: string;
}

declare module "mdast" {
	interface PhrasingContentMap {
		criticInsert: CriticInsert;
		criticDelete: CriticDelete;
	}
	interface RootContentMap {
		criticInsert: CriticInsert;
		criticDelete: CriticDelete;
	}
}

/** Match `{++…++}` or `{--…--}`. Lazy inner so `{++a++}{--b--}` two adjacent
 *  markers tokenize as two separate criticInsert/criticDelete nodes rather
 *  than one bloated insert with embedded marker text. */
const CRITIC_RE = /\{(\+\+|--)([\s\S]+?)\1\}/g;

/** Remark plugin: post-parse, walk every Text node and split it on CriticMarkup
 * markers, producing `criticInsert` / `criticDelete` phrasing siblings. Plain
 * text outside markers stays as `text`. Splitting happens AT text-node level —
 * a marker that straddles markdown formatting (`{++**bold**++}`) is NOT
 * recognized because the strong wrapper was already established at parse time;
 * to track formatted content, place the markers INSIDE the formatting
 * (`**{++bold++}**`). */
export const remarkCriticMarkup: Plugin<[], Root> = () => (tree) => {
	visit(tree, "text", (node, index, parent) => {
		if (!parent || index === undefined) return;
		if (!CRITIC_RE.test(node.value)) {
			CRITIC_RE.lastIndex = 0;
			return;
		}
		CRITIC_RE.lastIndex = 0;
		const split = splitCriticMarkup(node.value);
		if (split.length === 1 && split[0]?.type === "text") return;
		(parent as Parent).children.splice(index, 1, ...split);
		// Skip the nodes we just inserted; visit moves to the next sibling.
		return [SKIP, index + split.length];
	});
};

function splitCriticMarkup(
	value: string,
): Array<Text | CriticInsert | CriticDelete> {
	const out: Array<Text | CriticInsert | CriticDelete> = [];
	let cursor = 0;
	for (const match of value.matchAll(CRITIC_RE)) {
		const start = match.index;
		if (start === undefined) continue;
		if (start > cursor) {
			out.push({ type: "text", value: value.slice(cursor, start) });
		}
		const marker = match[1];
		const inner = match[2] ?? "";
		if (marker === "++") out.push({ type: "criticInsert", value: inner });
		else out.push({ type: "criticDelete", value: inner });
		cursor = start + match[0].length;
	}
	if (cursor < value.length) {
		out.push({ type: "text", value: value.slice(cursor) });
	}
	return out;
}
