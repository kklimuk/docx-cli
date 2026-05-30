import type { Document } from "../ast/document";
import { iterateBlocks } from "../ast/document/body";
import { IMAGE_RELATIONSHIP_TYPE } from "../ast/document/relationships";
import { a, pic, w, wp } from "../jsx";
import type { XmlNode } from "../parser";
import type { ImageSource } from "./source";

/** Cross-cutting lens over the document's inline images: walks the document
 * tree for image runs, manages the media parts + relationships those runs
 * reference, and enriches AST image runs with content hashes. Constructed
 * at call sites with `new Images(document)`; holds only a back-reference. */
export class Images {
	constructor(private document: Document) {}

	/** Every inline-picture run in document order — the same order `read`
	 * assigns `imgN` ids (only drawings whose blip resolves to a known image
	 * relationship count, matching `readImageFromDrawing`). Callers index by
	 * ordinal (`imgN`), filter by `relationshipId`, or splice via `parent`. */
	list(): ImageRunHit[] {
		const hits: ImageRunHit[] = [];
		const relationships = this.document.relationships;
		function walk(nodes: XmlNode[]): void {
			for (const node of nodes) {
				if (node.tag === "w:r") {
					for (const child of node.children) {
						if (child.tag !== "w:drawing") continue;
						const relationshipId = drawingRelationshipId(child);
						if (
							relationshipId &&
							relationships.imagesByRelationshipId.has(relationshipId)
						) {
							hits.push({ run: node, parent: nodes, relationshipId });
						}
					}
				}
				if (node.children.length > 0) walk(node.children);
			}
		}
		walk(this.document.documentTree);
		return hits;
	}

	/** Write image bytes into `word/media/imageN.ext`, mint an image
	 * relationship, and register the extension's content-type default.
	 * Returns the rId + partName so the caller's `<Image>` blip can reference
	 * the new relationship. An *operation*, not a component — mutates package
	 * state (writes a part, pushes a Relationship, edits content-types). */
	add(source: ImageSource): { relationshipId: string; partName: string } {
		const partName = nextMediaPartName(
			this.document.pkg.listParts(),
			source.extension,
		);
		this.document.pkg.writeBytes(partName, source.bytes);

		const relationshipId = this.document.relationships.add(
			IMAGE_RELATIONSHIP_TYPE,
			partName.slice("word/".length),
		);
		this.document.contentTypes.registerExtension(
			source.extension,
			source.mimeType,
		);
		this.document.relationships.imagesByRelationshipId.set(relationshipId, {
			partName,
			contentType: source.mimeType,
		});
		return { relationshipId, partName };
	}

	/** Walk every `ImageRun` in the body, fetch its underlying media bytes,
	 * compute the sha256, and write it back into the run's `hash` field.
	 * Idempotent — runs that already have a hash are skipped, and identical
	 * media (shared part) is hashed once and reused. Used by `images list` /
	 * `images extract` / `read --ast` where the manifest needs content hashes. */
	async enrichHashes(): Promise<void> {
		const seen = new Set<string>();
		for (const block of iterateBlocks(this.document.body.blocks)) {
			if (block.type !== "paragraph") continue;
			for (const run of block.runs) {
				if (run.type !== "image" || run.hash) continue;
				if (seen.has(run.id)) continue;
				seen.add(run.id);
				const reference = this.document.body.imageById.get(run.id);
				if (!reference) continue;
				const bytes = await this.document.pkg.readBytes(reference.partName);
				run.hash = await sha256Hex(bytes);
			}
		}
	}
}

/** A run wrapping a single inline picture. `a:`/`pic:` namespaces are declared
 * inline on `a:graphic`/`pic:pic` (matching Word) so the subtree is valid even
 * in documents whose root doesn't declare them. */
export function Image({
	relationshipId,
	drawingId,
	widthEmu,
	heightEmu,
	alt,
}: {
	relationshipId: string;
	drawingId: number;
	widthEmu: number;
	heightEmu: number;
	alt?: string;
}): XmlNode {
	const name = `Picture ${drawingId}`;
	const description = alt ?? "";
	return (
		<w.r>
			<w.drawing>
				<wp.inline distT="0" distB="0" distL="0" distR="0">
					<wp.extent cx={widthEmu} cy={heightEmu} />
					<wp.effectExtent l="0" t="0" r="0" b="0" />
					<wp.docPr id={drawingId} name={name} descr={description} />
					<wp.cNvGraphicFramePr>
						<a.graphicFrameLocks
							{...{ "xmlns:a": A_NAMESPACE }}
							noChangeAspect="1"
						/>
					</wp.cNvGraphicFramePr>
					<a.graphic {...{ "xmlns:a": A_NAMESPACE }}>
						<a.graphicData uri={DRAWINGML_PICTURE_URI}>
							<pic.pic {...{ "xmlns:pic": PIC_NAMESPACE }}>
								<pic.nvPicPr>
									<pic.cNvPr id={drawingId} name={name} descr={description} />
									<pic.cNvPicPr />
								</pic.nvPicPr>
								<pic.blipFill>
									<a.blip {...{ "r:embed": relationshipId }} />
									<a.stretch>
										<a.fillRect />
									</a.stretch>
								</pic.blipFill>
								<pic.spPr>
									<a.xfrm>
										<a.off x="0" y="0" />
										<a.ext cx={widthEmu} cy={heightEmu} />
									</a.xfrm>
									<a.prstGeom prst="rect">
										<a.avLst />
									</a.prstGeom>
								</pic.spPr>
							</pic.pic>
						</a.graphicData>
					</a.graphic>
				</wp.inline>
			</w.drawing>
		</w.r>
	);
}

/** Drawing object ids (`wp:docPr`/`pic:cNvPr` @id) must be unique per document
 * or Word flags corruption. Scan existing ids and return max + 1. */
export function nextDrawingId(documentTree: XmlNode[]): number {
	let highest = 0;
	function walk(node: XmlNode): void {
		if (node.tag === "wp:docPr" || node.tag === "pic:cNvPr") {
			const id = Number(node.getAttribute("id") ?? "0");
			if (Number.isFinite(id) && id > highest) highest = id;
		}
		for (const child of node.children) walk(child);
	}
	for (const root of documentTree) walk(root);
	return highest + 1;
}

/** A drawing run located in the document tree: the `<w:r>`, the array it lives
 * in (for splicing), and the image relationship its blip references. */
export type ImageRunHit = {
	run: XmlNode;
	parent: XmlNode[];
	relationshipId: string;
};

function nextMediaPartName(parts: string[], extension: string): string {
	let highest = 0;
	for (const part of parts) {
		const match = part.match(/^word\/media\/image(\d+)\./);
		if (!match) continue;
		const index = Number(match[1]);
		if (Number.isFinite(index) && index > highest) highest = index;
	}
	return `word/media/image${highest + 1}.${extension}`;
}

function drawingRelationshipId(drawing: XmlNode): string | undefined {
	const blip = drawing.findDescendant("a:blip");
	if (!blip) return undefined;
	return (
		blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? undefined
	);
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
	const owned = new Uint8Array(bytes.byteLength);
	owned.set(bytes);
	const buffer = await crypto.subtle.digest("SHA-256", owned);
	const document = new Uint8Array(buffer);
	let hex = "";
	for (let index = 0; index < document.length; index++) {
		const byte = document[index];
		if (byte === undefined) continue;
		hex += byte.toString(16).padStart(2, "0");
	}
	return hex;
}

const DRAWINGML_PICTURE_URI =
	"http://schemas.openxmlformats.org/drawingml/2006/picture";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NAMESPACE =
	"http://schemas.openxmlformats.org/drawingml/2006/picture";
