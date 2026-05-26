import { readdir } from "node:fs/promises";
import { resolve } from "node:path";
import JSZip from "jszip";

const root = resolve(import.meta.dir, "../../..");
const fixturesDir = resolve(root, "tests/fixtures");

type Report = {
	name: string;
	bytes: number;
	parts: number;
	wordParts: string[];
	hasComments: boolean;
	hasCommentsExt: boolean;
	tables: number;
	insertions: number;
	deletions: number;
	drawings: number;
	charts: boolean;
	math: boolean;
	hyperlinks: number;
	lists: number;
	headers: number;
	footers: number;
	footnotes: boolean;
	endnotes: boolean;
	sdt: number;
	vml: boolean;
	custStyles: number;
	embeddedFonts: boolean;
	macroEnabled: boolean;
	docxStrict: boolean;
};

function count(s: string, needle: string): number {
	let n = 0;
	let idx = s.indexOf(needle);
	while (idx !== -1) {
		n++;
		idx = s.indexOf(needle, idx + needle.length);
	}
	return n;
}

async function inspect(path: string): Promise<Report> {
	const buf = await Bun.file(path).arrayBuffer();
	const zip = await JSZip.loadAsync(buf);
	const parts: string[] = [];
	zip.forEach((p) => {
		parts.push(p);
	});

	const wordParts = parts.filter((p) => p.startsWith("word/")).sort();

	const doc = zip.file("word/document.xml");
	const docXml = doc ? await doc.async("string") : "";

	const styles = zip.file("word/styles.xml");
	const stylesXml = styles ? await styles.async("string") : "";

	const charts = parts.some(
		(p) => p.includes("/charts/") && p.endsWith(".xml"),
	);
	const math = docXml.includes("<m:oMath") || docXml.includes("xmlns:m=");
	const sdt = count(docXml, "<w:sdt>") + count(docXml, "<w:sdt ");
	const vml = docXml.includes("<v:") || parts.some((p) => p.endsWith(".vml"));
	const embeddedFonts = parts.some((p) => p.startsWith("word/fonts/"));

	// Strict OOXML uses different namespace
	const docxStrict =
		docXml.includes("schemas.microsoft.com/office/2006/wordml/strict") ||
		docXml.includes("/strict/") ||
		(docXml.includes("xmlns:w=") &&
			!docXml.includes("wordprocessingml/2006/main"));

	return {
		name: path.split("/").pop() ?? path,
		bytes: buf.byteLength,
		parts: parts.length,
		wordParts,
		hasComments: parts.includes("word/comments.xml"),
		hasCommentsExt: parts.includes("word/commentsExtended.xml"),
		tables: count(docXml, "<w:tbl>"),
		insertions: count(docXml, "<w:ins "),
		deletions: count(docXml, "<w:del "),
		drawings: count(docXml, "<w:drawing>"),
		charts,
		math,
		hyperlinks: count(docXml, "<w:hyperlink"),
		lists: count(docXml, "<w:numId"),
		headers: parts.filter((p) => /^word\/header\d+\.xml$/.test(p)).length,
		footers: parts.filter((p) => /^word\/footer\d+\.xml$/.test(p)).length,
		footnotes: parts.includes("word/footnotes.xml"),
		endnotes: parts.includes("word/endnotes.xml"),
		sdt,
		vml,
		custStyles: count(stylesXml, "<w:style "),
		embeddedFonts,
		macroEnabled:
			parts.some((p) => p.endsWith(".bin") && p.includes("vbaProject")) ||
			path.endsWith(".docm"),
		docxStrict,
	};
}

const files = (await readdir(fixturesDir))
	.filter((f) => f.endsWith(".docx") || f.endsWith(".docm"))
	.sort();

const reports: Report[] = [];
for (const f of files) {
	reports.push(await inspect(resolve(fixturesDir, f)));
}

const headers = [
	"name",
	"bytes",
	"parts",
	"comments",
	"cmtExt",
	"tables",
	"ins",
	"del",
	"drawings",
	"charts",
	"math",
	"links",
	"lists",
	"headers",
	"footers",
	"footnotes",
	"endnotes",
	"sdt",
	"vml",
	"styles",
	"fonts",
	"strict",
];

for (const r of reports) {
	const summary: Record<string, unknown> = { name: r.name, bytes: r.bytes };
	if (r.hasComments) summary.comments = true;
	if (r.hasCommentsExt) summary.commentsExt = true;
	if (r.tables) summary.tables = r.tables;
	if (r.insertions) summary.insertions = r.insertions;
	if (r.deletions) summary.deletions = r.deletions;
	if (r.drawings) summary.drawings = r.drawings;
	if (r.charts) summary.charts = true;
	if (r.math) summary.math = true;
	if (r.hyperlinks) summary.hyperlinks = r.hyperlinks;
	if (r.lists) summary.lists = r.lists;
	if (r.headers) summary.headers = r.headers;
	if (r.footers) summary.footers = r.footers;
	if (r.footnotes) summary.footnotes = true;
	if (r.endnotes) summary.endnotes = true;
	if (r.sdt) summary.sdt = r.sdt;
	if (r.vml) summary.vml = true;
	if (r.embeddedFonts) summary.embeddedFonts = true;
	if (r.docxStrict) summary.strict = true;
	if (r.custStyles) summary.styles = r.custStyles;
	console.log(JSON.stringify(summary));
}
void headers;
