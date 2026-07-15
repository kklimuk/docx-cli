import { diffArrays, structuredPatch } from "diff";

/** Context lines shown around each change — fixed (no CLI knob), git's default. */
const CONTEXT = 3;

/** Line-level unified diff between two texts, as structured hunks. Empty array ⇒
 *  the texts are identical. Thin wrapper over jsdiff's `structuredPatch`, mapped
 *  to our own shape so `--json` output and the string renderer share one source
 *  of truth (a jsdiff bump can't silently change either). */
export function diffHunks(oldText: string, newText: string): Hunk[] {
	const patch = structuredPatch("old", "new", oldText, newText, "", "", {
		context: CONTEXT,
	});
	return patch.hunks.map((hunk) => ({
		oldStart: hunk.oldStart,
		oldLines: hunk.oldLines,
		newStart: hunk.newStart,
		newLines: hunk.newLines,
		// jsdiff prefixes each content line with " "/"+"/"-"; a bare "\ No newline
		// at end of file" marker (starts with "\") is metadata, not content — drop it.
		lines: hunk.lines.filter((line) => line[0] !== "\\").map(toDiffLine),
	}));
}

/** Aggregate counts for the at-a-glance header (guards the large-diff-overwhelm
 *  failure — an agent gets the gist even when the body is long). */
export function diffStats(hunks: Hunk[]): DiffStats {
	let added = 0;
	let removed = 0;
	for (const hunk of hunks) {
		for (const line of hunk.lines) {
			if (line.kind === "insert") added++;
			else if (line.kind === "delete") removed++;
		}
	}
	return { hunks: hunks.length, added, removed };
}

/** Render hunks to unified-diff text: `--- oldLabel` / `+++ newLabel` headers
 *  then `@@ … @@` hunks. Empty string when there are no hunks. When `wordDiff`
 *  (default true), an equal-length delete/insert run (the common table-row /
 *  paragraph edit, where `read` puts the whole change on one long line) is
 *  paired by index and each pair refined into a single line with inline
 *  `[-removed-]{+added+}` markers; unequal runs (rows added/removed) stay plain
 *  `+`/`-`. */
export function renderUnified(
	hunks: Hunk[],
	options: { oldLabel?: string; newLabel?: string; wordDiff?: boolean } = {},
): string {
	if (hunks.length === 0) return "";
	const wordDiff = options.wordDiff ?? true;
	const out: string[] = [
		`--- ${options.oldLabel ?? "old"}`,
		`+++ ${options.newLabel ?? "new"}`,
	];
	for (const hunk of hunks) {
		out.push(
			`@@ -${hunk.oldStart},${hunk.oldLines} +${hunk.newStart},${hunk.newLines} @@`,
		);
		out.push(...renderHunkLines(hunk.lines, wordDiff));
	}
	return `${out.join("\n")}\n`;
}

function toDiffLine(line: string): DiffLine {
	const text = line.slice(1);
	if (line[0] === "+") return { kind: "insert", text };
	if (line[0] === "-") return { kind: "delete", text };
	return { kind: "context", text };
}

function renderHunkLines(lines: DiffLine[], wordDiff: boolean): string[] {
	const out: string[] = [];
	let index = 0;
	while (index < lines.length) {
		const line = lines[index];
		if (!line) break;
		if (line.kind === "context") {
			out.push(` ${line.text}`);
			index++;
			continue;
		}
		// Gather a maximal run of deletes followed by a maximal run of inserts, so
		// an isolated 1:1 change collapses to one refined line while a larger
		// delete-run/insert-run stays plain.
		const deletes: string[] = [];
		while (lines[index]?.kind === "delete") {
			// biome-ignore lint/style/noNonNullAssertion: guarded by the loop condition.
			deletes.push(lines[index]!.text);
			index++;
		}
		const inserts: string[] = [];
		while (lines[index]?.kind === "insert") {
			// biome-ignore lint/style/noNonNullAssertion: guarded by the loop condition.
			inserts.push(lines[index]!.text);
			index++;
		}
		// Equal-length delete/insert runs are positionally corresponding changes
		// (a form fill edits N adjacent rows in place; jsdiff emits all N deletes
		// then all N inserts). Pair them by index and word-refine each, so a block
		// of edited table rows reads as N refined lines, not N deletes + N inserts
		// the agent must re-pair by eye. Unequal runs (rows added/removed) can't be
		// safely paired — leave them plain.
		if (wordDiff && deletes.length > 0 && deletes.length === inserts.length) {
			for (let pair = 0; pair < deletes.length; pair++) {
				// biome-ignore lint/style/noNonNullAssertion: index < equal lengths.
				out.push(wordDiffLine(deletes[pair]!, inserts[pair]!));
			}
			continue;
		}
		for (const text of deletes) out.push(`-${text}`);
		for (const text of inserts) out.push(`+${text}`);
	}
	return out;
}

/** One changed line, refined to `[-removed-]{+added+}` inline markers (git
 *  `--word-diff` style). Unchanged tokens pass through verbatim. Tokenizes with
 *  HTML tags and `<!-- … -->` comments as ATOMIC units (via `diffArrays`), so a
 *  change straddling markup never splits a tag into `<[-/mark>` garbage — the
 *  read view is dense with `<span>`/`<mark>`/`<u>` runs and `docx:` annotations. */
function wordDiffLine(oldLine: string, newLine: string): string {
	let out = "";
	for (const part of diffArrays(tokenizeLine(oldLine), tokenizeLine(newLine))) {
		const value = part.value.join("");
		if (part.removed) out += `[-${value}-]`;
		else if (part.added) out += `{+${value}+}`;
		else out += value;
	}
	return out;
}

/** Split a line into atoms for word-level diffing: each `<!-- … -->` comment and
 *  each `<…>` tag is one token (never split mid-markup), then whitespace runs and
 *  word runs. The trailing `<` alternative catches a bare `<` that starts neither
 *  a comment nor a tag (e.g. `$x < y$`, `5 < 10`) so it's kept as its own token
 *  rather than dropped — every character must survive the tokenize→join round-trip
 *  or the refined line would silently lose it. */
function tokenizeLine(line: string): string[] {
	return line.match(/<!--[\s\S]*?-->|<[^>]*>|\s+|[^\s<]+|</g) ?? [];
}

export type DiffLine = { kind: "context" | "delete" | "insert"; text: string };
export type Hunk = {
	oldStart: number;
	oldLines: number;
	newStart: number;
	newLines: number;
	lines: DiffLine[];
};
export type DiffStats = { hunks: number; added: number; removed: number };
