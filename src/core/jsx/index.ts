import { type XmlAttributes, XmlNode } from "../parser";

// biome-ignore lint/suspicious/noExplicitAny: JSX needs to accept any prop shape
export type JsxProps = { [key: string]: any } | null;
export type JsxChild =
	| XmlNode
	| XmlNode[]
	| string
	| number
	| false
	| null
	| undefined;
export type TagFn = (props: JsxProps) => XmlNode;
export type FragmentFn = TagFn;

const FRAGMENT_TAG = "#fragment";

export const Fragment: FragmentFn = (props) => {
	const wrapper = new XmlNode(FRAGMENT_TAG);
	flatten(normalizeChildren(props?.children), wrapper.children);
	return wrapper;
};

export function normalizeChildren(children: unknown): JsxChild[] {
	if (children === undefined || children === null) return [];
	if (Array.isArray(children)) return children as JsxChild[];
	return [children as JsxChild];
}

export type Namespace<TagName extends string> = { [Tag in TagName]: TagFn };

export function namespace<TagName extends string>(
	prefix: string,
	tags: readonly TagName[],
): Namespace<TagName> {
	const result = {} as Record<string, TagFn>;
	for (const tagName of tags) {
		result[tagName] = makeTag(`${prefix}:${tagName}`);
	}
	return result as Namespace<TagName>;
}

const W_TAGS = [
	"document",
	"body",
	"p",
	"pPr",
	"pStyle",
	"jc",
	"numPr",
	"ilvl",
	"numId",
	"r",
	"rPr",
	"rStyle",
	"t",
	"b",
	"i",
	"u",
	"strike",
	"color",
	"highlight",
	"sz",
	"rFonts",
	"br",
	"tab",
	"drawing",
	"sectPr",
	"sectPrChange",
	"cols",
	"type",
	"tbl",
	"tblPr",
	"tblPrChange",
	"tblGrid",
	"tblGridChange",
	"gridCol",
	"tr",
	"trPr",
	"tc",
	"tcPr",
	"tcPrChange",
	"ins",
	"del",
	"cellIns",
	"cellDel",
	"delText",
	"commentRangeStart",
	"commentRangeEnd",
	"commentReference",
	"comments",
	"comment",
	"settings",
	"trackChanges",
	"hyperlink",
	"styles",
	"style",
	"name",
	"basedOn",
	"next",
	"qFormat",
	"keepNext",
	"keepLines",
	"spacing",
	"outlineLvl",
	"ind",
	"vertAlign",
	"pBdr",
	"bottom",
	"top",
	"left",
	"right",
	"insideH",
	"insideV",
	"tblBorders",
	"tcBorders",
	"tblW",
	"tcW",
	"tblLayout",
	"vMerge",
	"gridSpan",
	"numbering",
	"num",
	"abstractNum",
	"abstractNumId",
	"multiLevelType",
	"lvl",
	"start",
	"numFmt",
	"lvlText",
	"lvlJc",
	"contextualSpacing",
] as const;

const CP_TAGS = ["coreProperties", "lastModifiedBy"] as const;
const DC_TAGS = ["title", "creator", "subject", "description"] as const;
const DCTERMS_TAGS = ["created", "modified"] as const;
const W15_TAGS = ["commentsEx", "commentEx"] as const;

// Relationship + DrawingML + picture namespaces. Emit-side only, with no
// internal caller today — r/a/wp/pic get wired by S5 (image insertion: an
// inline drawing references its blip via r:embed and the a:/wp:/pic: graphic
// tree); w14 (paraId/textId) is reserved for paragraph-mark identity. `@public`
// keeps knip from flagging them as dead until those consumers land.
const R_TAGS = ["embed", "link", "id"] as const;
const A_TAGS = [
	"graphic",
	"graphicData",
	"blip",
	"xfrm",
	"off",
	"ext",
	"prstGeom",
	"avLst",
] as const;
const WP_TAGS = [
	"inline",
	"anchor",
	"extent",
	"effectExtent",
	"docPr",
	"cNvGraphicFramePr",
] as const;
const PIC_TAGS = [
	"pic",
	"nvPicPr",
	"cNvPr",
	"cNvPicPr",
	"blipFill",
	"spPr",
] as const;
const W14_TAGS = ["paraId", "textId"] as const;

export const w = namespace("w", W_TAGS);
export const cp = namespace("cp", CP_TAGS);
export const dc = namespace("dc", DC_TAGS);
export const dcterms = namespace("dcterms", DCTERMS_TAGS);
export const w15 = namespace("w15", W15_TAGS);
/** @public Emit-side image/drawing namespaces — wired by S5 (image insertion). */
export const r = namespace("r", R_TAGS);
/** @public */
export const a = namespace("a", A_TAGS);
/** @public */
export const wp = namespace("wp", WP_TAGS);
/** @public */
export const pic = namespace("pic", PIC_TAGS);
/** @public Emit-side paragraph-mark identity namespace (paraId/textId). */
export const w14 = namespace("w14", W14_TAGS);

function makeTag(qualifiedName: string): TagFn {
	return (props) => {
		const attributes: XmlAttributes = {};
		let childrenProp: unknown;
		if (props) {
			for (const [key, value] of Object.entries(props)) {
				if (key === "children") {
					childrenProp = value;
					continue;
				}
				if (value === false || value == null) continue;
				attributes[mapAttributeName(key)] = String(value);
			}
		}
		const childNodes: XmlNode[] = [];
		flatten(normalizeChildren(childrenProp), childNodes);
		return new XmlNode(qualifiedName, attributes, childNodes);
	};
}

function flatten(items: JsxChild[], out: XmlNode[]): void {
	for (const item of items) {
		if (item == null || item === false) continue;
		if (Array.isArray(item)) {
			flatten(item, out);
			continue;
		}
		if (item instanceof XmlNode) {
			if (item.tag === FRAGMENT_TAG) {
				for (const child of item.children) out.push(child);
				continue;
			}
			out.push(item);
			continue;
		}
		if (typeof item === "string" || typeof item === "number") {
			out.push(XmlNode.textNode(String(item)));
		}
	}
}

function mapAttributeName(name: string): string {
	const dashIndex = name.indexOf("-");
	if (dashIndex === -1) return name;
	return `${name.slice(0, dashIndex)}:${name.slice(dashIndex + 1)}`;
}
