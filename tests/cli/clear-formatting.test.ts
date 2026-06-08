import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import {
	clearFormatting,
	resolveClearTags,
} from "../../src/core/edit/clear-formatting";
import { XmlNode } from "../../src/core/parser";
import { runCli, tempWorkspace } from "./harness";

// `find --highlight` + `edit --clear` is the highlight-removal workflow that
// took a weak model ~40 commands; it should now be find → clear.

const SOURCE =
	'Fill [the state]{highlight="yellow"} and [the county]{highlight="yellow"}; keep [this]{color="FF0000"} bold [word]{highlight="yellow"}.\n';

async function doc(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, SOURCE);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
	return path;
}

async function locators(path: string, ...flags: string[]): Promise<string[]> {
	const find = await runCli("find", path, ...flags, "--all");
	return (find.parsed as { matches: Array<{ locator: string }> }).matches.map(
		(match) => match.locator,
	);
}

describe("find formatting filters + edit --clear", () => {
	test("find --highlight yellow returns each highlighted span", async () => {
		const path = await doc("find-hl");
		const found = await locators(path, "--highlight", "yellow");
		expect(found.length).toBe(3); // the state / the county / word
	});

	test("find --highlight any matches any highlight color", async () => {
		const path = await doc("find-hl-any");
		expect((await locators(path, "--highlight", "any")).length).toBe(3);
	});

	test("find --color FF0000 returns the colored span", async () => {
		const path = await doc("find-color");
		const found = await locators(path, "--color", "FF0000");
		expect(found.length).toBe(1);
	});

	test("edit --clear highlight on found spans removes only highlight", async () => {
		const path = await doc("clear-hl");
		for (const loc of await locators(path, "--highlight", "any")) {
			expect(
				(await runCli("edit", path, "--at", loc, "--clear", "highlight"))
					.exitCode,
			).toBe(0);
		}
		// all highlights gone; text + the red color preserved
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(1);
		const read = await runCli("read", path);
		expect(read.stdout).toContain("Fill the state and the county");
	});

	test("edit --clear all on a whole paragraph strips all run formatting", async () => {
		const path = await doc("clear-all");
		expect(
			(await runCli("edit", path, "--at", "p0", "--clear", "all")).exitCode,
		).toBe(0);
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(0);
	});

	test("edit --clear with an unknown attribute is a usage error", async () => {
		const path = await doc("clear-bad");
		const result = await runCli("edit", path, "--at", "p0", "--clear", "bogus");
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

describe("clear-formatting preserves unmodeled rPr children (in-place invariant)", () => {
	test("--clear all strips formatting tags but keeps an unmodeled <w:lang>", () => {
		// The feature mutates rPr in place precisely so props we don't model
		// survive. `all` must remove highlight/bold but leave <w:lang> untouched.
		const [para] = XmlNode.parse(
			`<w:p><w:r><w:rPr><w:highlight w:val="yellow"/><w:b/><w:lang w:val="fr-FR"/></w:rPr><w:t>bonjour</w:t></w:r></w:p>`,
		);
		const tags = resolveClearTags(["all"]);
		expect(tags).not.toBeNull();
		clearFormatting(para as XmlNode, null, tags as Set<string>);
		const xml = XmlNode.serialize([para as XmlNode]);
		expect(xml).not.toContain("w:highlight");
		expect(xml).not.toContain("<w:b/>");
		expect(xml).toContain('w:lang w:val="fr-FR"'); // unmodeled prop survives
		expect(xml).toContain("bonjour");
	});
});
