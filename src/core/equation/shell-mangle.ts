import { mathSpans } from "../track-changes/equation-text";
import { formatEquationToken } from "./token";

/** Why `value` — new content that drops the `dropped` equations from a
 *  paragraph that held them — looks like a shell gutted it, or null.
 *
 *  An agent that writes `docx edit FILE --at p0 "Energy is $E=mc^2$ now"` in
 *  DOUBLE quotes hands the CLI `Energy is =mc^2$ now`: bash expanded `$E` to
 *  nothing (and would turn `$$` into its PID). The edit then reads as "delete
 *  the equation" and exits 0. Dropping an equation is legitimate, so the
 *  refusal needs the drop AND a telltale the shell leaves behind:
 *
 *  - a STRAY `$` — one that is neither part of a well-formed `$…$`/`$$…$$`
 *    span, nor a price (`$` + digit, which double-quoted bash can't produce:
 *    `$5` is a positional parameter and expands too), nor escaped (`\$`);
 *  - a dropped equation's token as bash would have expanded it (`=mc^2$`
 *    for `$E=mc^2$`, `12345x=112345` for `$$x=1$$`) found in the value.
 *
 *  `rejectShellMangledValue` (cli/parse-helpers.ts) is the precedent for
 *  prices (`"$300"` → `00`). This one only fires when equations are actually
 *  dropped, so a "costs $5" in an equation-free paragraph is never examined. */
export function shellMangledEquationLoss(
	value: string,
	dropped: { latex: string; display: boolean }[],
): ShellMangling | null {
	if (dropped.length === 0) return null;
	let repaired = value;
	let reason: string | null = null;
	for (const equation of dropped) {
		const token = formatEquationToken(equation.latex, "", equation.display);
		const remnant = bashExpansionPattern(token);
		const shown = remnant ? repaired.match(remnant)?.[0] : undefined;
		if (!remnant || shown === undefined) continue;
		reason ??= `it contains "${shown}" — what bash leaves of ${token} in double quotes`;
		// Put the token back where bash gutted it: the hint can then show the
		// exact line the agent meant to send.
		repaired = repaired.replace(remnant, () => token);
	}
	if (reason) return { reason, repaired };
	const stray = strayDollar(value);
	if (stray === null) return null;
	return {
		reason: `it has a stray "$" ("${value.slice(Math.max(0, stray - 8), stray + 8)}") — the half of a $…$ equation the shell didn't expand`,
	};
}

/** Why a value looks shell-mangled, and — when every gutted token was
 *  recognized — the value with them restored (what the agent meant). */
export type ShellMangling = { reason: string; repaired?: string };

/** The hint every equation-related refusal ends with: put `$`-bearing content
 *  in a batch FILE, which no shell expands, rather than an inline argument —
 *  and quote the heredoc if the file is written from the shell, since bash
 *  expands `$` inside an unquoted one. `line` is one example JSONL entry
 *  (JSON-encoded here, which also shows how a LaTeX backslash is escaped). */
export function batchFileHint(
	command: string,
	file: string,
	line: Record<string, string>,
): string {
	return `Put text containing $ in a batch FILE, not a shell argument — the shell can't expand JSON in a file: \`docx ${command} FILE --batch ${file}\` with the line ${JSON.stringify(line)}. Writing the file from the shell? Use a QUOTED heredoc (\`cat > ${file} <<'EOF'\` … \`EOF\`) — bash expands $ inside an unquoted one.`;
}

/** A regex matching what bash makes of `token` inside double quotes, or null
 *  when bash would leave it alone (or leave too little to recognize). */
function bashExpansionPattern(token: string): RegExp | null {
	let pattern = "";
	let literal = "";
	let changed = false;
	let index = 0;
	while (index < token.length) {
		const char = token[index] ?? "";
		const next = token[index + 1] ?? "";
		if (char === "\\" && (next === "\\" || next === "$" || next === "`")) {
			pattern += escapeRegex(next);
			literal += next;
			changed = true;
			index += 2;
			continue;
		}
		if (char !== "$") {
			pattern += escapeRegex(char);
			literal += char;
			index++;
			continue;
		}
		const rest = token.slice(index + 1);
		const expansion =
			/^(\$|[?#]|[A-Za-z_][A-Za-z0-9_]*|[0-9@*!-]|\{[^}]*\})/.exec(rest);
		if (!expansion) {
			pattern += "\\$";
			literal += "$";
			index++;
			continue;
		}
		const [name = ""] = expansion;
		changed = true;
		// `$$` is the shell's PID, `$?`/`$#` a number; the rest expand to an
		// unset variable's nothing (`$-` to a few flag letters).
		if (name === "$" || name === "?" || name === "#") pattern += "\\d+";
		else if (name === "-") pattern += "[A-Za-z]*";
		index += 1 + name.length;
	}
	if (!changed || literal.replace(/\$/g, "").trim().length < 2) return null;
	return new RegExp(pattern);
}

function escapeRegex(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** The index of the first `$` that is no math delimiter, price or escape. */
function strayDollar(value: string): number | null {
	const covered = new Set<number>();
	for (const span of mathSpans(value)) {
		for (let index = span.start; index < span.end; index++) covered.add(index);
	}
	for (let index = 0; index < value.length; index++) {
		if (value[index] !== "$" || covered.has(index)) continue;
		if (value[index - 1] === "\\") continue;
		if (/\d/.test(value[index + 1] ?? "")) continue;
		return index;
	}
	return null;
}
