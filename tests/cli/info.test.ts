import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli } from "./harness";

describe("docx info schema", () => {
	test("default outputs JSON Schema", async () => {
		const result = await runCli("info", "schema");
		expect(result.exitCode).toBe(0);
		const schema = result.parsed as { $defs: Record<string, unknown> };
		expect(Object.keys(schema.$defs)).toEqual(
			expect.arrayContaining([
				"Block",
				"Paragraph",
				"Run",
				"TextRun",
				"ImageRun",
				"Comment",
			]),
		);
	});

	test("--ts outputs the live types.ts source", async () => {
		const result = await runCli("info", "schema", "--ts");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("export type DocProperties");
		expect(result.stdout).toContain("export type Paragraph");
		expect(result.stdout).toContain("CommentAnchor");
	});
});

describe("docx info locators", () => {
	test("default outputs the grammar reference", async () => {
		const result = await runCli("info", "locators");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("LOCATOR GRAMMAR");
		expect(result.stdout).toContain("p3:5-20");
	});

	test("--json outputs structured reference", async () => {
		const result = await runCli("info", "locators", "--json");
		expect(result.exitCode).toBe(0);
		const reference = result.parsed as {
			blockLocators: Record<string, unknown>;
			spanLocators: { span: { syntax: string } };
		};
		expect(reference.spanLocators.span.syntax).toBe("pN:S-E");
	});
});

describe("docx info skill", () => {
	test("default emits SKILL.md frontmatter + body", async () => {
		const result = await runCli("info", "skill");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.startsWith("---\nname: docx-cli\n")).toBe(true);
		expect(result.stdout).toContain('description: "');
		expect(result.stdout).toContain("# docx-cli");
		expect(result.stdout).toContain("docx info locators");
	});

	test("--json exposes name, description, body", async () => {
		const result = await runCli("info", "skill", "--json");
		expect(result.exitCode).toBe(0);
		const skill = result.parsed as {
			name: string;
			description: string;
			body: string;
		};
		expect(skill.name).toBe("docx-cli");
		// Harnesses match the request against `description`; keep it under the
		// ~500-char activation recommendation.
		expect(skill.description.length).toBeLessThanOrEqual(500);
		expect(skill.body).toContain("Golden workflows");
	});

	test("committed skills/docx-cli/SKILL.md is in sync with the binary", async () => {
		// Regenerate with: docx info skill > skills/docx-cli/SKILL.md
		const result = await runCli("info", "skill");
		const committed = await Bun.file(
			join(import.meta.dir, "../../skills/docx-cli/SKILL.md"),
		).text();
		expect(result.stdout).toBe(committed);
	});

	test("manifest + npm descriptions match the binary's shortDescription", async () => {
		// The store tagline is single-sourced from src/cli/info/skill.ts. If this fails,
		// resync the manifests to `docx info skill --json | .shortDescription`.
		const result = await runCli("info", "skill", "--json");
		const { shortDescription } = result.parsed as { shortDescription: string };
		const root = join(import.meta.dir, "../..");
		const marketplace = (await Bun.file(
			join(root, ".claude-plugin/marketplace.json"),
		).json()) as { description: string; plugins: { description: string }[] };
		const claudePlugin = (await Bun.file(
			join(root, ".claude-plugin/plugin.json"),
		).json()) as { description: string };
		const codexPlugin = (await Bun.file(
			join(root, ".codex-plugin/plugin.json"),
		).json()) as { description: string };
		const packageJson = (await Bun.file(join(root, "package.json")).json()) as {
			description: string;
		};
		expect(marketplace.description).toBe(shortDescription);
		expect(marketplace.plugins[0]?.description).toBe(shortDescription);
		expect(claudePlugin.description).toBe(shortDescription);
		expect(codexPlugin.description).toBe(shortDescription);
		expect(packageJson.description).toBe(shortDescription);
	});

	test("codex plugin manifest version tracks package.json", async () => {
		// .codex-plugin/plugin.json is the one manifest that pins a version, with no other
		// sync — assert it equals package.json so a release bump can't leave Codex
		// advertising a stale version (the description drift test doesn't cover version).
		const root = join(import.meta.dir, "../..");
		const packageJson = (await Bun.file(join(root, "package.json")).json()) as {
			version: string;
		};
		const codexPlugin = (await Bun.file(
			join(root, ".codex-plugin/plugin.json"),
		).json()) as { version: string };
		expect(codexPlugin.version).toBe(packageJson.version);
	});
});

describe("docx info dispatcher", () => {
	test("no topic prints help with usage exit", async () => {
		const result = await runCli("info");
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("docx info");
		expect(result.stdout).toContain("schema");
		expect(result.stdout).toContain("locators");
		expect(result.stdout).toContain("skill");
	});

	test("unknown topic returns USAGE", async () => {
		const result = await runCli("info", "nope");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});
