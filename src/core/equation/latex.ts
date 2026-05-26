/** LaTeX text- and token-shaping helpers used by the OMML→LaTeX walker, and
 *  destined to be reused by the LaTeX→OMML emitter (Phase 2) — the
 *  Unicode↔command map and the special-char escape rules are inverted there
 *  but built from the same tables.
 *
 *  The maps are deliberately keyed by Unicode codepoint so we can hand-author
 *  them without worrying about parsing complications. We accept that some
 *  symbols have multiple LaTeX spellings (`\varepsilon` vs `\epsilon`) and
 *  pick whichever Pandoc-emit / Pandoc-read pair round-trips cleanest. */

/** Common Unicode → LaTeX command mapping for math text. Unmapped chars
 *  pass through as literal Unicode — modern Pandoc + unicode-math packages
 *  render those correctly, so this map is a quality upgrade for the common
 *  cases, not a correctness requirement. */
export const TEXT_LATEX_MAP: Record<string, string> = {
	// Greek lowercase
	α: "\\alpha",
	β: "\\beta",
	γ: "\\gamma",
	δ: "\\delta",
	ε: "\\varepsilon",
	ζ: "\\zeta",
	η: "\\eta",
	θ: "\\theta",
	ι: "\\iota",
	κ: "\\kappa",
	λ: "\\lambda",
	μ: "\\mu",
	ν: "\\nu",
	ξ: "\\xi",
	π: "\\pi",
	ρ: "\\rho",
	σ: "\\sigma",
	τ: "\\tau",
	υ: "\\upsilon",
	φ: "\\varphi",
	χ: "\\chi",
	ψ: "\\psi",
	ω: "\\omega",
	ϵ: "\\epsilon",
	ϕ: "\\phi",
	ϑ: "\\vartheta",
	ϱ: "\\varrho",
	ϖ: "\\varpi",
	// Greek uppercase (italicized ones get the corresponding command)
	Γ: "\\Gamma",
	Δ: "\\Delta",
	Θ: "\\Theta",
	Λ: "\\Lambda",
	Ξ: "\\Xi",
	Π: "\\Pi",
	Σ: "\\Sigma",
	Υ: "\\Upsilon",
	Φ: "\\Phi",
	Ψ: "\\Psi",
	Ω: "\\Omega",
	// Operators
	"×": "\\times",
	"÷": "\\div",
	"·": "\\cdot",
	"±": "\\pm",
	"∓": "\\mp",
	"∘": "\\circ",
	"∙": "\\bullet",
	"⋅": "\\cdot",
	"⊕": "\\oplus",
	"⊗": "\\otimes",
	"⊖": "\\ominus",
	"⊘": "\\oslash",
	"⊥": "\\perp",
	"∥": "\\parallel",
	// Relations
	"≤": "\\leq",
	"≥": "\\geq",
	"≠": "\\neq",
	"≈": "\\approx",
	"≡": "\\equiv",
	"≅": "\\cong",
	"∼": "\\sim",
	"∝": "\\propto",
	"≪": "\\ll",
	"≫": "\\gg",
	// Sets
	"∈": "\\in",
	"∉": "\\notin",
	"∋": "\\ni",
	"⊂": "\\subset",
	"⊃": "\\supset",
	"⊆": "\\subseteq",
	"⊇": "\\supseteq",
	"∪": "\\cup",
	"∩": "\\cap",
	"∅": "\\emptyset",
	// Calculus
	"∂": "\\partial",
	"∇": "\\nabla",
	"∞": "\\infty",
	// Logic
	"∧": "\\wedge",
	"∨": "\\vee",
	"¬": "\\neg",
	"∀": "\\forall",
	"∃": "\\exists",
	"⇒": "\\Rightarrow",
	"⇐": "\\Leftarrow",
	"⇔": "\\Leftrightarrow",
	"→": "\\to",
	"←": "\\leftarrow",
	"↔": "\\leftrightarrow",
	"↦": "\\mapsto",
	"↑": "\\uparrow",
	"↓": "\\downarrow",
	// Dots and ellipses
	"…": "\\ldots",
	"⋯": "\\cdots",
	"⋮": "\\vdots",
	"⋱": "\\ddots",
	// Standalone math letters that don't live in the systematic
	// Mathematical Alphanumeric Symbols block.
	ℓ: "\\ell",
	ℏ: "\\hbar",
	// (ℝ, ℕ, ℤ, ℂ, ℙ, ℍ, ℚ, ℜ, ℑ, and the script / fraktur letters live in
	// MATH_ALPHA_EXCEPTIONS instead — they decode systematically via the
	// `\mathbb{R}` / `\mathfrak{R}` path so `\mathbb{Y}` and `\mathbb{R}`
	// produce consistent output.)
	// Unicode punctuation normalization → ASCII LaTeX-source form.
	"−": "-", // U+2212 MINUS SIGN → ASCII hyphen-minus
	"⁻": "-", // U+207B SUPERSCRIPT MINUS (mhchem uses this in script positions)
	" ": " ", // U+00A0 NO-BREAK SPACE → regular space (avoids LaTeX munging)
};

/** Zero-width Unicode characters Word and Pandoc emit as placeholders inside
 *  otherwise-empty `<m:sub>` / `<m:sup>`. We strip them so an empty subscript
 *  contributes no LaTeX (rather than `_{}`).
 *
 *  Includes the Variation Selectors block (U+FE00–U+FE0F). VS15 / VS16
 *  appear in temml's output as a "use text variant" hint on math operators
 *  (e.g., `\oplus` emits `⊕` + VS15) but Word doesn't recognize them and
 *  renders the selector as a missing-glyph box. Stripping is harmless —
 *  Word picks the text variant by default for math runs. */
const ZERO_WIDTH_CHARS = new Set([
	"​", // zero width space
	"‌", // zero width non-joiner
	"‍", // zero width joiner
	"﻿", // zero width no-break space (BOM)
	...Array.from({ length: 16 }, (_, i) => String.fromCodePoint(0xfe00 + i)), // variation selectors VS1–VS16
]);

/** Predicate-form check used by the writer's leaf handlers — stripping these
 *  chars before they enter OMML keeps Word from showing missing-glyph boxes
 *  for the variation selectors. */
export function isMathTextNoise(ch: string): boolean {
	return ZERO_WIDTH_CHARS.has(ch);
}

/** LaTeX special characters that must be escaped when emitted from literal
 *  math text. We deliberately omit `&` (it's the eqArr column separator;
 *  literal `&` in math content is exceedingly rare and would have entered
 *  via `\text{...}` anyway). */
const LATEX_ESCAPE_RE = /([\\$%#_{}])/g;

/** Convert raw text from an `<m:t>` node into LaTeX-safe text: map common
 *  Unicode math chars to commands (with a trailing space to disambiguate
 *  from following letters: `\alpha x` parses as `\alpha` then `x`, while
 *  `\alphax` parses as a single undefined command), decode Mathematical
 *  Alphanumeric Symbols (U+1D400–U+1D7FF and friends) into `\mathbf{…}` /
 *  `\mathbb{…}` / etc., escape LaTeX-special chars, normalize Unicode
 *  punctuation (`−` → `-`), and strip zero-width placeholder characters.
 *
 *  Consecutive math-alpha chars of the same style are grouped into a
 *  single wrap (`𝐱𝐲𝐳` → `\mathbf{xyz}`, not `\mathbf{x}\mathbf{y}\mathbf{z}`). */
export function escapeAndMap(text: string): string {
	let out = "";
	// Iterate codepoints (math-alpha chars are outside the BMP — surrogate pairs
	// in UTF-16 — so we can't use a simple char loop without splitting them).
	const codepoints = [...text];
	let i = 0;
	while (i < codepoints.length) {
		const ch = codepoints[i] as string;
		if (ZERO_WIDTH_CHARS.has(ch)) {
			i++;
			continue;
		}
		const decoded = decodeMathAlpha(ch);
		if (decoded) {
			// Greedy: consume contiguous math-alpha chars of the same style.
			let letters = decoded.letter;
			let j = i + 1;
			while (j < codepoints.length) {
				const nextDecoded = decodeMathAlpha(codepoints[j] as string);
				if (!nextDecoded || nextDecoded.style !== decoded.style) break;
				letters += nextDecoded.letter;
				j++;
			}
			out += `\\${decoded.style}{${letters}}`;
			i = j;
			continue;
		}
		const mapped = TEXT_LATEX_MAP[ch];
		if (mapped) {
			// Commands need a trailing space to disambiguate; plain-char mappings
			// (Unicode → ASCII normalization) don't.
			out += mapped.startsWith("\\") ? `${mapped} ` : mapped;
			i++;
			continue;
		}
		out += ch.replace(LATEX_ESCAPE_RE, "\\$1");
		i++;
	}
	return out;
}

/** Decode a single Unicode codepoint from the Mathematical Alphanumeric
 *  Symbols block (U+1D400–U+1D7FF) into its `\mathbf` / `\mathit` / etc.
 *  style and base "letter" (which may be a multi-char LaTeX command like
 *  `\sigma` for Greek). Also covers the 24 "holes" in the block where
 *  Unicode points to pre-existing codepoints elsewhere (Planck h, the
 *  blackboard letters at U+2102 / U+210D / …, etc.).
 *
 *  Two contiguous sub-blocks are decoded:
 *   - Latin (U+1D400–U+1D6A3): 13 styles × 52 chars (A–Z then a–z)
 *   - Greek (U+1D6A8–U+1D7C9): 5 styles × 58 chars (Greek caps + nabla +
 *     small Greek + variant supplements). Only the bold-italic block is
 *     decoded with a `\boldsymbol` wrap; the other styles (bold-upright,
 *     plain-italic, sans variants) pass through as raw Unicode because
 *     they're rare and don't have widely-supported LaTeX names. */
export function decodeMathAlpha(
	ch: string,
): { style: string; letter: string } | undefined {
	const exception = MATH_ALPHA_EXCEPTIONS[ch];
	if (exception) return exception;
	const code = ch.codePointAt(0);
	if (code === undefined) return undefined;
	if (code >= 0x1d400 && code <= 0x1d6a3) {
		const offset = code - 0x1d400;
		const styleIndex = Math.floor(offset / 52);
		if (styleIndex >= MATH_ALPHA_STYLES.length) return undefined;
		const style = MATH_ALPHA_STYLES[styleIndex];
		if (!style) return undefined;
		const letterIndex = offset % 52;
		const letter =
			letterIndex < 26
				? String.fromCharCode(0x41 + letterIndex) // A–Z
				: String.fromCharCode(0x61 + letterIndex - 26); // a–z
		return { style, letter };
	}
	if (code >= 0x1d6a8 && code <= 0x1d6e1) {
		// Bold Greek (upright). temml uses this for `\boldsymbol{\Omega}`-style
		// caps (Greek upper-case is upright-by-default in many conventions, so
		// "bold upright" is the natural bolded form). Decode to `\boldsymbol`
		// so the wrap matches the user's likely intent.
		const offset = code - 0x1d6a8;
		const command = GREEK_BOLD_ITALIC_COMMANDS[offset];
		if (command) return { style: "boldsymbol", letter: command };
	}
	if (code >= 0x1d71c && code <= 0x1d755) {
		// Bold-italic Greek (the `\boldsymbol{\alpha}` case). 58 chars: 26 caps
		// + variant theta + nabla, then 25 smalls, then 7 supplements.
		const offset = code - 0x1d71c;
		const command = GREEK_BOLD_ITALIC_COMMANDS[offset];
		if (command) return { style: "boldsymbol", letter: command };
	}
	return undefined;
}

/** LaTeX command (with backslash) — or a literal letter for the few Greek
 *  capitals that lack a unique command (Alpha → A, Beta → B, etc.). 58
 *  entries in the bold-italic Greek block order. */
const GREEK_BOLD_ITALIC_COMMANDS: readonly string[] = [
	// Capitals (0–24)
	"A", // 0  Alpha
	"B", // 1  Beta
	"\\Gamma", // 2
	"\\Delta", // 3
	"E", // 4  Epsilon
	"Z", // 5  Zeta
	"H", // 6  Eta
	"\\Theta", // 7
	"I", // 8  Iota
	"K", // 9  Kappa
	"\\Lambda", // 10
	"M", // 11 Mu
	"N", // 12 Nu
	"\\Xi", // 13
	"O", // 14 Omicron
	"\\Pi", // 15
	"P", // 16 Rho
	"\\vartheta", // 17 ϴ variant Theta
	"\\Sigma", // 18
	"T", // 19 Tau
	"\\Upsilon", // 20
	"\\Phi", // 21
	"X", // 22 Chi
	"\\Psi", // 23
	"\\Omega", // 24
	"\\nabla", // 25
	// Smalls (26–50)
	"\\alpha", // 26
	"\\beta", // 27
	"\\gamma", // 28
	"\\delta", // 29
	"\\epsilon", // 30
	"\\zeta", // 31
	"\\eta", // 32
	"\\theta", // 33
	"\\iota", // 34
	"\\kappa", // 35
	"\\lambda", // 36
	"\\mu", // 37
	"\\nu", // 38
	"\\xi", // 39
	"o", // 40 omicron — no LaTeX command, just Latin o
	"\\pi", // 41
	"\\rho", // 42
	"\\varsigma", // 43 final sigma ς
	"\\sigma", // 44
	"\\tau", // 45
	"\\upsilon", // 46
	"\\phi", // 47
	"\\chi", // 48
	"\\psi", // 49
	"\\omega", // 50
	// Supplements (51–57)
	"\\partial", // 51
	"\\varepsilon", // 52
	"\\vartheta", // 53
	"\\varkappa", // 54
	"\\varphi", // 55
	"\\varrho", // 56
	"\\varpi", // 57
];

/** Latin-letter styles in U+1D400–U+1D6A3, in block order. 13 styles × 52
 *  letters each. Names use the `unicode-math` package's naming convention
 *  (`mathbfit` for bold italic, `mathbfsfit` for bold sans italic) — these
 *  are well-supported in modern LaTeX (`unicode-math` for XeLaTeX/LuaLaTeX)
 *  and degrade understandably in plain LaTeX (an unknown `\mathbfit` will
 *  produce a warning but render the argument). */
const MATH_ALPHA_STYLES = [
	"mathbf", // U+1D400
	"mathit", // U+1D434
	"boldsymbol", // U+1D468 — bold italic Latin; `\boldsymbol` is the widely
	// supported amsmath name (unicode-math uses `\mathbfit`, but `\boldsymbol`
	// works in classic LaTeX too via amsmath, which is loaded almost everywhere)
	"mathscr", // U+1D49C
	"mathbfscr", // U+1D4D0
	"mathfrak", // U+1D504
	"mathbb", // U+1D538
	"mathbffrak", // U+1D56C
	"mathsf", // U+1D5A0
	"mathbfsf", // U+1D5D4
	"mathsfit", // U+1D608
	"mathbfsfit", // U+1D63C
	"mathtt", // U+1D670
];

/** The 24 "holes" in the Mathematical Alphanumeric Symbols block where
 *  Unicode says "the letter you want lives at a separate, pre-existing
 *  codepoint" (Planck h, the blackboard / script / fraktur letters that
 *  appeared in Unicode before the systematic block existed). Producers
 *  emit the alternate codepoint, so we have to recognize them too.
 *
 *  Reference: Unicode Standard, Table 22-1 (Mathematical Alphanumeric
 *  Symbols Exceptions). Also documented in the `unicode-math` package's
 *  data tables. */
const MATH_ALPHA_EXCEPTIONS: Record<string, { style: string; letter: string }> =
	{
		// Italic h (Planck constant)
		ℎ: { style: "mathit", letter: "h" },
		// Script (`mathscr`) capitals
		ℬ: { style: "mathscr", letter: "B" },
		ℰ: { style: "mathscr", letter: "E" },
		ℱ: { style: "mathscr", letter: "F" },
		ℋ: { style: "mathscr", letter: "H" },
		ℐ: { style: "mathscr", letter: "I" },
		ℒ: { style: "mathscr", letter: "L" },
		ℳ: { style: "mathscr", letter: "M" },
		ℛ: { style: "mathscr", letter: "R" },
		ℯ: { style: "mathscr", letter: "e" },
		ℊ: { style: "mathscr", letter: "g" },
		ℴ: { style: "mathscr", letter: "o" },
		// Fraktur (`mathfrak`) capitals
		ℭ: { style: "mathfrak", letter: "C" },
		ℌ: { style: "mathfrak", letter: "H" },
		ℑ: { style: "mathfrak", letter: "I" },
		ℜ: { style: "mathfrak", letter: "R" },
		ℨ: { style: "mathfrak", letter: "Z" },
		// Blackboard bold (`mathbb`) capitals — these are also in TEXT_LATEX_MAP
		// under the old `\mathbb{R}`-as-special-token mapping, but we'd rather
		// decode them through the systematic path so the algorithm stays
		// internally consistent.
		ℂ: { style: "mathbb", letter: "C" },
		ℍ: { style: "mathbb", letter: "H" },
		ℕ: { style: "mathbb", letter: "N" },
		ℙ: { style: "mathbb", letter: "P" },
		ℚ: { style: "mathbb", letter: "Q" },
		ℝ: { style: "mathbb", letter: "R" },
		ℤ: { style: "mathbb", letter: "Z" },
	};

/** Match LaTeX's "single token" rule for super/subscripts. A token is:
 *   - a single character that's safe as a bare token (letter/digit/Greek/
 *     common punctuation like `+`, `-`, `*`, `'`)
 *   - a backslash-command optionally followed by a trailing space (`\alpha `)
 *   - a backslash-command followed by an optional `[…]` arg and exactly
 *     ONE balanced `{…}` arg (`\hat{x}`, `\sqrt[3]{x^{2k+1}}`, `\text{H}`) —
 *     uses brace-depth tracking so nested braces in the argument count
 *     correctly. Multi-arg commands like `\frac{a}{b}` and
 *     `\textcolor{red}{X}` are EXCLUDED: KaTeX-based renderers (Obsidian,
 *     VS Code preview) refuse to bind `^\textcolor{r}{x}` because they
 *     expect the script's argument to be a single atom — they parse
 *     `\textcolor` as the atom and then balk on the missing args. Wrapping
 *     multi-arg commands in `{…}` ourselves avoids the parse error.
 *  Everything else wraps in braces. */
export function convertGrouped(body: string): string {
	const trimmed = body.trimEnd();
	if (trimmed.length === 0) return "{}";
	// Single-char fast path: covers `x`, `2`, `α` (Greek), and the common
	// scripts-with-symbol shapes `H^+`, `x^-`, `f^*`, `y'` — LaTeX parses each
	// as a single token after the `^`/`_`, so the braces are visual noise.
	if (trimmed.length === 1 && /[\p{L}\p{N}+\-*'.,]/u.test(trimmed))
		return trimmed;
	if (/^\\[A-Za-z]+\s?$/.test(body)) return body.trimEnd();
	if (isSingleArgCommand(trimmed)) return trimmed;
	return `{${trimmed}}`;
}

/** True if `source` is exactly one backslash-command with an optional
 *  `[opt]` argument and one balanced `{…}` required argument — and nothing
 *  after it. Brace-depth aware: `\sqrt{x^{2k+1}}` and
 *  `\hat{\boldsymbol{\alpha}}` are recognized as single commands despite
 *  nested braces.
 *
 *  This is the standard LaTeX "single token" recognition that TeX,
 *  KaTeX, MathJax, and Pandoc all implement via depth-tracking linear
 *  scans — see [findMatchingBrace](#findMatchingBrace). */
export function isSingleArgCommand(source: string): boolean {
	const match = /^\\[A-Za-z]+/.exec(source);
	if (!match) return false;
	let cursor = match[0].length;
	if (source[cursor] === "[") {
		const closed = findMatchingBracket(source, cursor);
		if (closed < 0) return false;
		cursor = closed;
	}
	if (source[cursor] !== "{") return false;
	const closed = findMatchingBrace(source, cursor);
	return closed === source.length;
}

/** Return the index AFTER the `}` that matches the `{` at `openIndex`, or
 *  `-1` if unbalanced. Escape sequences `\{`, `\}`, and `\\` are skipped
 *  (their second character is a literal, not a group marker) — same rule
 *  TeX uses for tokenization. */
export function findMatchingBrace(source: string, openIndex: number): number {
	if (source[openIndex] !== "{") return -1;
	let depth = 1;
	let cursor = openIndex + 1;
	while (cursor < source.length && depth > 0) {
		const char = source[cursor];
		if (char === "\\") {
			const next = source[cursor + 1];
			if (next === "{" || next === "}" || next === "\\") {
				cursor += 2;
				continue;
			}
		}
		if (char === "{") depth++;
		else if (char === "}") depth--;
		cursor++;
	}
	return depth === 0 ? cursor : -1;
}

/** Bracket-pair counterpart to `findMatchingBrace` for `[…]` optional
 *  arguments (`\sqrt[3]{x}`). Symmetric implementation. */
export function findMatchingBracket(source: string, openIndex: number): number {
	if (source[openIndex] !== "[") return -1;
	let depth = 1;
	let cursor = openIndex + 1;
	while (cursor < source.length && depth > 0) {
		const char = source[cursor];
		if (char === "\\") {
			const next = source[cursor + 1];
			if (next === "[" || next === "]" || next === "\\") {
				cursor += 2;
				continue;
			}
		}
		if (char === "[") depth++;
		else if (char === "]") depth--;
		cursor++;
	}
	return depth === 0 ? cursor : -1;
}

/** Strip outer `\cmd{…}` wrappers throughout `source`, brace-depth aware.
 *  Used by the formatting bubble-up in [read.ts](./read.ts) to collapse
 *  nested `\boldsymbol{\boldsymbol{X}}` (which arises when bold-Greek
 *  codepoint decoding meets a bold-context outer wrap) into
 *  `\boldsymbol{X}`. Handles inner braces in `\cmd{X^{2}}` correctly. */
export function stripCommandWrap(source: string, command: string): string {
	const opening = `\\${command}{`;
	let out = "";
	let cursor = 0;
	while (cursor < source.length) {
		if (source.startsWith(opening, cursor)) {
			const braceOpen = cursor + opening.length - 1; // index of `{`
			const closed = findMatchingBrace(source, braceOpen);
			if (closed > 0) {
				out += source.slice(braceOpen + 1, closed - 1);
				cursor = closed;
				continue;
			}
		}
		out += source[cursor] ?? "";
		cursor++;
	}
	return out;
}

/** If `name` starts with a known LaTeX function-operator (lim, sin, cos, log,
 *  max, min, sup, inf, det, gcd, dim, etc.), replace that prefix with the
 *  `\name` command form. Preserves any sub/sup suffix the OMML attached. */
export function promoteOperatorName(name: string): string {
	const match = name.match(/^([A-Za-z]+)(.*)$/);
	if (!match) return name;
	const [, head, tail] = match;
	if (head && KNOWN_OPERATORS.has(head)) return `\\${head}${tail ?? ""}`;
	return name;
}

/** LaTeX function operators that get typeset upright (not italic). When the
 *  OMML reader sees these as a function name (`<m:fName>lim</m:fName>`), it
 *  promotes them to the `\<name>` command form. */
export const KNOWN_OPERATORS = new Set([
	"lim",
	"liminf",
	"limsup",
	"sin",
	"cos",
	"tan",
	"sec",
	"csc",
	"cot",
	"arcsin",
	"arccos",
	"arctan",
	"sinh",
	"cosh",
	"tanh",
	"coth",
	"log",
	"ln",
	"lg",
	"exp",
	"max",
	"min",
	"sup",
	"inf",
	"det",
	"gcd",
	"dim",
	"ker",
	"hom",
	"deg",
	"arg",
	"Pr",
]);
