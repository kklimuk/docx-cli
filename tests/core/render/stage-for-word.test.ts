import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Pkg } from "../../../src/core/ast/document/package";
import { stageForWord } from "../../../src/core/render/engines/stage-for-word";

// Word for Mac blocks on open with "This document contains fields that may
// refer to other files. Do you want to update the fields…?" when a field is
// flagged dirty or settings ask to update fields on open — and `display alerts
// to none` doesn't suppress it, so an unattended render hung on the user's
// screen. The staged copy Word opens must carry neither.

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:dirty="true" w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr=" NUMPAGES " w:dirty="1"><w:r><w:t>4</w:t></w:r></w:fldSimple><w:r><w:t>dirty="true" stays as text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r></w:p></w:ftr>`;
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:updateFields w:val="true"/><w:defaultTabStop w:val="720"/></w:settings>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

async function buildDocx(parts: Record<string, string>): Promise<string> {
	const path = join(mkdtempSync(join(tmpdir(), "stage-for-word-")), "in.docx");
	const pkg = Pkg.empty(path);
	for (const [name, content] of Object.entries(parts))
		pkg.writeText(name, content);
	await pkg.save();
	return path;
}

function stagedPath(input: string): string {
	return input.replace(/in\.docx$/, "staged.docx");
}

describe("stageForWord", () => {
	test("strips update-on-open and dirty field flags from every field part", async () => {
		const input = await buildDocx({
			"word/document.xml": DOCUMENT,
			"word/footer1.xml": FOOTER,
			"word/settings.xml": SETTINGS,
			"word/styles.xml": STYLES,
		});
		const staged = stagedPath(input);
		await stageForWord(input, staged);

		const pkg = await Pkg.open(staged);
		const document = await pkg.readText("word/document.xml");
		expect(document).toContain('<w:fldChar w:fldCharType="begin"/>');
		expect(document).toContain('<w:fldSimple w:instr=" NUMPAGES ">');
		expect(document).not.toMatch(/<w:fld[^>]*dirty/);
		// Text that merely spells the attribute is content, not markup.
		expect(document).toContain('dirty="true" stays as text');
		expect(await pkg.readText("word/footer1.xml")).toContain(
			'<w:fldChar w:fldCharType="begin"/>',
		);
		const settings = await pkg.readText("word/settings.xml");
		expect(settings).not.toContain("updateFields");
		expect(settings).toContain('<w:zoom w:percent="100"/><w:defaultTabStop');
		expect(await pkg.readText("word/styles.xml")).toBe(STYLES);

		// The user's file is untouched.
		const original = await Pkg.open(input);
		expect(await original.readText("word/settings.xml")).toBe(SETTINGS);
	});

	test("a document without the flags is copied byte-for-byte", async () => {
		const input = await buildDocx({
			"word/document.xml": DOCUMENT.replace(/ w:dirty="[^"]*"/g, ""),
			"word/settings.xml": SETTINGS.replace(
				'<w:updateFields w:val="true"/>',
				"",
			),
		});
		const staged = stagedPath(input);
		await stageForWord(input, staged);
		expect(await Bun.file(staged).bytes()).toEqual(
			await Bun.file(input).bytes(),
		);
	});

	test("a file that isn't a zip is copied as-is for Word's own error path", async () => {
		const input = join(
			mkdtempSync(join(tmpdir(), "stage-for-word-")),
			"in.docx",
		);
		await Bun.write(input, "not a zip");
		const staged = stagedPath(input);
		await stageForWord(input, staged);
		expect(await Bun.file(staged).text()).toBe("not a zip");
	});
});
