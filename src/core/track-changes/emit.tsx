import { type JsxChild, w } from "../jsx";
import type { XmlNode } from "../parser";
import type { TrackedMeta } from "./index";

export function Ins({
	meta,
	children,
}: {
	meta: TrackedMeta;
	children?: JsxChild;
}): XmlNode {
	return (
		<w.ins
			w-id={String(meta.revisionId)}
			w-author={meta.author}
			w-date={meta.date}
		>
			{children}
		</w.ins>
	);
}

export function Del({
	meta,
	children,
}: {
	meta: TrackedMeta;
	children?: JsxChild;
}): XmlNode {
	return (
		<w.del
			w-id={String(meta.revisionId)}
			w-author={meta.author}
			w-date={meta.date}
		>
			{children}
		</w.del>
	);
}

export function markParagraphMarkAs(
	paragraph: XmlNode,
	kind: "ins" | "del",
	meta: TrackedMeta,
): void {
	let pPr = paragraph.findChild("w:pPr");
	if (!pPr) {
		pPr = (<w.pPr />) as XmlNode;
		paragraph.children.unshift(pPr);
	}
	let rPr = pPr.findChild("w:rPr");
	if (!rPr) {
		rPr = (<w.rPr />) as XmlNode;
		pPr.children.push(rPr);
	}
	const marker = kind === "ins" ? <Ins meta={meta} /> : <Del meta={meta} />;
	rPr.children.push(marker);
}
