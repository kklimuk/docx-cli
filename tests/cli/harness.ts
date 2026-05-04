import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const BINARY = join(import.meta.dir, "..", "..", "src", "index.ts");

export type CliResult = {
	stdout: string;
	stderr: string;
	exitCode: number;
	parsed?: unknown;
};

export async function runCli(...args: string[]): Promise<CliResult> {
	const proc = Bun.spawn(["bun", BINARY, ...args], {
		stdout: "pipe",
		stderr: "pipe",
	});
	const stdout = await new Response(proc.stdout).text();
	const stderr = await new Response(proc.stderr).text();
	const exitCode = await proc.exited;
	const trimmed = stdout.trim();
	let parsed: unknown;
	if (trimmed.length > 0) {
		try {
			parsed = JSON.parse(trimmed);
		} catch {
			// not JSON — leave parsed undefined
		}
	}
	return { stdout, stderr, exitCode, parsed };
}

export function tempWorkspace(label: string): string {
	return mkdtempSync(join(tmpdir(), `docx-cli-${label}-`));
}
