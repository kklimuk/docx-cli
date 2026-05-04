import { XmlNode } from "./xml-node";

export function runTextLength(run: XmlNode): number {
	let total = 0;
	for (const child of run.children) {
		if (child.tag === "w:t") total += child.collectText().length;
	}
	return total;
}

export function sliceRun(run: XmlNode, start: number, end: number): XmlNode {
	const sliced = new XmlNode("w:r", { ...run.attributes });
	let consumed = 0;
	for (const child of run.children) {
		if (child.tag === "w:t") {
			const text = child.collectText();
			const localStart = Math.max(0, start - consumed);
			const localEnd = Math.min(text.length, end - consumed);
			if (localStart < localEnd) {
				const slicedText = new XmlNode("w:t", { "xml:space": "preserve" });
				slicedText.children.push(
					XmlNode.textNode(text.slice(localStart, localEnd)),
				);
				sliced.children.push(slicedText);
			}
			consumed += text.length;
			continue;
		}
		sliced.children.push(child.clone());
	}
	return sliced;
}
