import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/package";
import { runCli, tempWorkspace } from "./harness";

/**
 * Regression coverage for Bug A (parser dropped <w:t> siblings of inline
 * children) and Bug B (sliceRun cloned non-text inline markers into every
 * slice). The mnda.docx fixture has two <w:r> elements that put a <w:tab/>
 * and a <w:t> in the same run:
 *   <w:r>...<w:tab/><w:t xml:space="preserve">Expires </w:t></w:r>
 *   <w:r>...<w:tab/><w:t>Continues until terminated ...</w:t></w:r>
 */
describe("mnda fixture — tab + text in same <w:r>", () => {
	const FIXTURE = "tests/fixtures/mnda.docx";

	test("Bug A: AST exposes <w:t> text that sits after a <w:tab/> sibling", async () => {
		const result = await runCli("read", FIXTURE);
		const allText = collectAllText(result.parsed);
		expect(allText).toContain("Expires ");
		expect(allText).toContain(
			"Continues until terminated in accordance with the terms of the MNDA.",
		);
	});

	test("Bug A: find locates text after a tab in the same run", async () => {
		const result = await runCli("find", FIXTURE, "Continues until terminated");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ text: string }>;
		};
		expect(payload.totalMatches).toBeGreaterThanOrEqual(1);
		expect(payload.matches[0]?.text).toBe("Continues until terminated");
	});

	test("Bug B: track-changes-on replace of '1 year(s)' wraps only the year run; the adjacent tab+Expires run is untouched", async () => {
		const workspace = tempWorkspace("mnda-replace");
		const docPath = join(workspace, "mnda.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));

		await runCli("track-changes", docPath, "on");
		const replaceResult = await runCli(
			"replace",
			docPath,
			"1 year(s)",
			"2 year(s)",
			"--all",
			"--author",
			"Tester",
		);
		expect(replaceResult.exitCode).toBe(0);

		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");

		// The "Expires " run should still be a single <w:r> carrying the tab and
		// the original text, not split or duplicated by sliceRun.
		const expiresRuns = matchAll(
			documentXml,
			/<w:r\b[^>]*>(?:(?!<\/w:r>).)*?<w:tab\/>(?:(?!<\/w:r>).)*?<w:t[^>]*>Expires <\/w:t>(?:(?!<\/w:r>).)*?<\/w:r>/g,
		);
		expect(expiresRuns).toHaveLength(1);

		// The "Continues until terminated" run is the same shape — verify it
		// also survives untouched (the replace span doesn't reach this paragraph,
		// but XmlNode.serialize round-tripping should not damage it either).
		const continuesRuns = matchAll(
			documentXml,
			/<w:r\b[^>]*>(?:(?!<\/w:r>).)*?<w:tab\/>(?:(?!<\/w:r>).)*?<w:t[^>]*>Continues until terminated[^<]*<\/w:t>(?:(?!<\/w:r>).)*?<\/w:r>/g,
		);
		expect(continuesRuns).toHaveLength(1);

		// "1 year(s)" appears twice in the source. Each match should produce one
		// <w:del> wrapping the original run and one adjacent <w:ins> carrying
		// "2 year(s)".
		const delsWith1Year = matchAll(
			documentXml,
			/<w:del\b[^>]*w:author="Tester"[^>]*>(?:(?!<\/w:del>).)*?<w:delText[^>]*>1 year\(s\)<\/w:delText>(?:(?!<\/w:del>).)*?<\/w:del>/g,
		);
		const insWith2Year = matchAll(
			documentXml,
			/<w:ins\b[^>]*w:author="Tester"[^>]*>(?:(?!<\/w:ins>).)*?<w:t[^>]*>2 year\(s\)<\/w:t>(?:(?!<\/w:ins>).)*?<\/w:ins>/g,
		);
		expect(delsWith1Year).toHaveLength(2);
		expect(insWith2Year).toHaveLength(2);

		// And the original literal "1 year(s)" should no longer appear inside
		// any <w:t> (only inside <w:delText> within <w:del>).
		const liveOneYear = matchAll(documentXml, /<w:t[^>]*>1 year\(s\)<\/w:t>/g);
		expect(liveOneYear).toHaveLength(0);
	});
});

function collectAllText(parsed: unknown): string {
	const doc = parsed as {
		blocks: Array<{
			type: string;
			runs?: Array<{ type: string; text?: string }>;
			rows?: Array<{
				cells: Array<{
					blocks: Array<{
						type: string;
						runs?: Array<{ type: string; text?: string }>;
					}>;
				}>;
			}>;
		}>;
	};
	const out: string[] = [];
	function walk(
		blocks: Array<{
			type: string;
			runs?: Array<{ type: string; text?: string }>;
			rows?: Array<{
				cells: Array<{
					blocks: Array<{
						type: string;
						runs?: Array<{ type: string; text?: string }>;
					}>;
				}>;
			}>;
		}>,
	): void {
		for (const block of blocks) {
			if (block.runs) {
				for (const run of block.runs) {
					if (run.type === "text" && typeof run.text === "string") {
						out.push(run.text);
					}
				}
			}
			if (block.rows) {
				for (const row of block.rows) {
					for (const cell of row.cells) {
						walk(cell.blocks);
					}
				}
			}
		}
	}
	walk(doc.blocks);
	return out.join("");
}

function matchAll(haystack: string, pattern: RegExp): string[] {
	const out: string[] = [];
	for (const match of haystack.matchAll(pattern)) {
		out.push(match[0]);
	}
	return out;
}
