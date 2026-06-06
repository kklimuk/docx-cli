import { describe, expect, test } from "bun:test";
import { runCli } from "./harness";

// The full command tree. Every command and sub-verb must answer `--help` with a
// usable screen — this is the regression guard for the help-drift bug class
// (an implemented flag with no docs, or docs for a flag that doesn't exist).
const COMMANDS: string[][] = [
	["create"],
	["read"],
	["edit"],
	["insert"],
	["delete"],
	["find"],
	["replace"],
	["wc"],
	["outline"],
	["render"],
	["info", "schema"],
	["info", "locators"],
	["comments", "add"],
	["comments", "reply"],
	["comments", "resolve"],
	["comments", "delete"],
	["comments", "list"],
	["footnotes", "add"],
	["footnotes", "edit"],
	["footnotes", "delete"],
	["footnotes", "list"],
	["endnotes", "add"],
	["endnotes", "edit"],
	["endnotes", "delete"],
	["endnotes", "list"],
	["images", "list"],
	["images", "extract"],
	["images", "replace"],
	["images", "delete"],
	["hyperlinks", "add"],
	["hyperlinks", "list"],
	["hyperlinks", "replace"],
	["hyperlinks", "delete"],
	["tables", "insert-row"],
	["tables", "delete-row"],
	["tables", "insert-column"],
	["tables", "delete-column"],
	["tables", "set-widths"],
	["tables", "merge"],
	["tables", "unmerge"],
	["tables", "borders"],
	["track-changes", "list"],
	["track-changes", "accept"],
	["track-changes", "reject"],
];

// Commands that take a locator advertise the unified `--at` (or the placement /
// slice variants) — none should still mention a removed addressing flag.
const REMOVED_ADDRESSING_FLAGS = ["--range ", "--id ", "--to cN", "--to ID"];

describe("help smoke", () => {
	for (const command of COMMANDS) {
		const label = command.join(" ");
		test(`docx ${label} --help`, async () => {
			const result = await runCli(...command, "--help");
			expect(result.exitCode).toBe(0);
			expect(result.stdout).toContain("Usage:");
			// The comprehensive-help pass gave every command an Output section
			// describing its success/error shape. The `info` reference printers
			// are the exception — they ARE the output, described inline.
			if (command[0] !== "info") {
				expect(result.stdout).toContain("Output:");
			}
			for (const removed of REMOVED_ADDRESSING_FLAGS) {
				expect(result.stdout).not.toContain(removed);
			}
		});
	}
});
