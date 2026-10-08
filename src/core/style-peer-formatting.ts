import type { Document } from "./ast/document";
import { flattenParagraphs } from "./ast/text";
import type { Paragraph } from "./ast/types";
import { inheritRprChild, insertPprChildInOrder } from "./blocks";
import { isCodeBlockStyleId } from "./code-block/style";
import {
	paragraphStyleId,
	paragraphsIn,
	textRuns,
} from "./paragraph-inheritance";
import { XmlNode } from "./parser";

/** Give each new STYLED paragraph the direct formatting its style peers agree on
 * — what makes a new `# Experience` look like the Heading1s already on the page.
 * Real documents style their headings partly through the style and partly by
 * direct formatting stamped on each one: the résumé's Heading1s are centered by
 * a `<w:jc>` on every heading paragraph, not by the style, so a heading built
 * from the style alone rendered left-aligned beside centered siblings (haiku,
 * three runs). For every layout/look property the new paragraph doesn't set
 * itself, it takes the value a strict majority of the existing paragraphs with
 * the same `pStyle` carry; a property the peers don't agree on (or mostly
 * leave to the style) is left alone. No peers → nothing changes, and the style
 * alone decides.
 *
 * Paragraph side: alignment, spacing, indentation, borders, shading. Run side:
 * the emphasis/look properties headings carry directly (bold, italic, caps,
 * color, underline, character spacing) — face and size are the dominant-
 * formatting pass's job. List items are skipped (their indent IS the list
 * geometry; `continueList` gives a bullet beside a list that list's look) and
 * so are code blocks (their look is their style). */
export function applyStylePeerFormatting(
	document: Document,
	newBlocks: XmlNode[],
): void {
	const consensusByStyle = new Map<string, PeerConsensus | null>();
	for (const paragraph of paragraphsIn(newBlocks)) {
		const pPr = paragraph.findChild("w:pPr");
		const styleId = paragraphStyleId(paragraph);
		if (!pPr || !styleId || isCodeBlockStyleId(styleId)) continue;
		if (pPr.findChild("w:numPr")) continue;
		if (!consensusByStyle.has(styleId)) {
			consensusByStyle.set(styleId, peerConsensus(document, styleId));
		}
		const consensus = consensusByStyle.get(styleId);
		if (!consensus) continue;
		applyConsensus(paragraph, pPr, consensus);
	}
}

const PEER_PPR_TAGS = [
	"w:jc",
	"w:spacing",
	"w:ind",
	"w:pBdr",
	"w:shd",
] as const;

const PEER_RPR_TAGS = [
	"w:b",
	"w:i",
	"w:caps",
	"w:smallCaps",
	"w:color",
	"w:u",
	"w:spacing",
] as const;

type PeerConsensus = { pPr: XmlNode[]; rPr: XmlNode[] };

/** The pPr/rPr children a strict majority of the style's existing (text-
 * bearing, non-list) paragraphs carry verbatim, or null with no peers. */
function peerConsensus(
	document: Document,
	styleId: string,
): PeerConsensus | null {
	const peers: XmlNode[] = [];
	for (const paragraph of stylePeers(document, styleId)) {
		const node = document.body.blockReferences.get(paragraph.id)?.node;
		if (!node || node.findChild("w:pPr")?.findChild("w:numPr")) continue;
		peers.push(node);
	}
	if (peers.length === 0) return null;
	const pPrs = peers.map((peer) => peer.findChild("w:pPr"));
	const rPrs = peers.map(uniformRunProperties);
	return {
		pPr: majorityChildren(pPrs, PEER_PPR_TAGS, peers.length),
		rPr: majorityChildren(rPrs, PEER_RPR_TAGS, peers.length),
	};
}

/** The rPr children EVERY text run of a peer carries — the paragraph's look,
 * not one run's decoration. Reading only the first run made a bold lead-in
 * (`**1. Definitions.** The following…`, the shape of most clause styles) a
 * "consensus", so every run of a new paragraph in that style came out bold. */
function uniformRunProperties(paragraph: XmlNode): XmlNode | undefined {
	const [first, ...rest] = textRuns(paragraph);
	const rPr = first?.findChild("w:rPr");
	if (!rPr) return undefined;
	// Only the tags a consensus can carry are worth serializing — on body-style
	// peers that's thousands of runs.
	const candidates = rPr.children.filter((child) =>
		PEER_RPR_TAG_SET.has(child.tag),
	);
	if (candidates.length === 0) return undefined;
	const others = rest.map(
		(run) =>
			new Set(
				(run.findChild("w:rPr")?.children ?? [])
					.filter((child) => PEER_RPR_TAG_SET.has(child.tag))
					.map((child) => XmlNode.serialize([child])),
			),
	);
	const uniform = new XmlNode("w:rPr");
	uniform.children = candidates.filter((child) => {
		const signature = XmlNode.serialize([child]);
		return others.every((set) => set.has(signature));
	});
	return uniform;
}

const PEER_RPR_TAG_SET: ReadonlySet<string> = new Set(PEER_RPR_TAGS);

/** For each tag, the child (by exact serialization) that MORE than half of the
 * peers carry. A peer without the tag counts against every value, so a
 * property most peers leave to the style never gets stamped. */
function majorityChildren(
	containers: (XmlNode | undefined)[],
	tags: readonly string[],
	peerCount: number,
): XmlNode[] {
	const winners: XmlNode[] = [];
	for (const tag of tags) {
		const counts = new Map<string, { node: XmlNode; count: number }>();
		for (const container of containers) {
			const child = container?.findChild(tag);
			if (!child) continue;
			const key = XmlNode.serialize([child]);
			const entry = counts.get(key) ?? { node: child, count: 0 };
			entry.count += 1;
			counts.set(key, entry);
		}
		for (const { node, count } of counts.values()) {
			if (count * 2 > peerCount) winners.push(node);
		}
	}
	return winners;
}

/** Stamp the consensus onto one new paragraph without overriding anything it
 * sets itself (an explicit `--alignment` still wins). */
function applyConsensus(
	paragraph: XmlNode,
	pPr: XmlNode,
	consensus: PeerConsensus,
): void {
	for (const child of consensus.pPr) {
		if (pPr.findChild(child.tag)) continue;
		insertPprChildInOrder(pPr, child.clone());
	}
	if (consensus.rPr.length === 0) return;
	for (const run of textRuns(paragraph)) {
		let rPr = run.findChild("w:rPr");
		if (!rPr) {
			rPr = new XmlNode("w:rPr");
			run.children.unshift(rPr);
		}
		for (const child of consensus.rPr) inheritRprChild(rPr, child);
	}
}

/** The existing paragraphs in `styleId` that carry visible text — the
 * reference population both the peer-consensus and the dominant-formatting
 * passes match a new paragraph against. */
export function stylePeers(document: Document, styleId: string): Paragraph[] {
	return flattenParagraphs(document.body.blocks).filter(
		(paragraph) =>
			paragraph.style === styleId &&
			paragraph.runs.some(
				(run) => run.type === "text" && run.text.trim().length > 0,
			),
	);
}
