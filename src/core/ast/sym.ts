/**
 * Decode a `<w:sym w:font="..." w:char="HHHH"/>` reference to a Unicode
 * character.
 *
 * The classic Word symbol fonts (Symbol, Wingdings, Wingdings 2, Webdings,
 * Zapf Dingbats) predate Unicode and store their glyphs at codepoints that
 * either reuse ASCII positions (Symbol — Greek/math at 0x21–0x7E) or live in
 * the Private Use Area (Wingdings & friends — 0xF020+). The displayed glyph
 * has a real Unicode home, but it's not at the same codepoint.
 *
 * We maintain best-effort tables for the most commonly used glyphs in each
 * font. Anything outside the table falls through to a literal codepoint
 * decode (`String.fromCodePoint`) so the AST stays usable for documents
 * that lean on symbol fonts we don't fully cover. Font names are matched
 * case-insensitively.
 */
export function decodeSym(font: string, charHex: string): string {
	const codepoint = Number.parseInt(charHex, 16);
	if (!Number.isFinite(codepoint)) return "";
	const fontKey = font.toLowerCase().trim();
	const table = TABLES.get(fontKey);
	if (table) {
		const mapped = table.get(codepoint);
		if (mapped !== undefined) return mapped;
	}
	// Wingdings/Webdings store glyphs at PUA codepoints (0xF020+) but Word
	// sometimes emits the ASCII alias (0x20+) instead — try the ASCII shadow
	// before falling through.
	if (table && codepoint >= 0xf000 && codepoint <= 0xf0ff) {
		const aliased = table.get(codepoint - 0xf000);
		if (aliased !== undefined) return aliased;
	}
	if (table && codepoint >= 0x20 && codepoint <= 0xff) {
		const aliased = table.get(codepoint + 0xf000);
		if (aliased !== undefined) return aliased;
	}
	try {
		return String.fromCodePoint(codepoint);
	} catch {
		return "";
	}
}

/** Adobe Symbol → Unicode. Source: Adobe Symbol Encoding mapping (the same
 * table Microsoft / Apple ship in their PostScript fonts). Covers ASCII
 * 0x21–0x7E plus the most common 0xA0+ extensions. */
const SYMBOL: ReadonlyArray<[number, string]> = [
	[0x21, "!"],
	[0x22, "∀"],
	[0x23, "#"],
	[0x24, "∃"],
	[0x25, "%"],
	[0x26, "&"],
	[0x27, "∋"],
	[0x28, "("],
	[0x29, ")"],
	[0x2a, "∗"],
	[0x2b, "+"],
	[0x2c, ","],
	[0x2d, "−"],
	[0x2e, "."],
	[0x2f, "/"],
	[0x30, "0"],
	[0x31, "1"],
	[0x32, "2"],
	[0x33, "3"],
	[0x34, "4"],
	[0x35, "5"],
	[0x36, "6"],
	[0x37, "7"],
	[0x38, "8"],
	[0x39, "9"],
	[0x3a, ":"],
	[0x3b, ";"],
	[0x3c, "<"],
	[0x3d, "="],
	[0x3e, ">"],
	[0x3f, "?"],
	[0x40, "≅"],
	[0x41, "Α"],
	[0x42, "Β"],
	[0x43, "Χ"],
	[0x44, "Δ"],
	[0x45, "Ε"],
	[0x46, "Φ"],
	[0x47, "Γ"],
	[0x48, "Η"],
	[0x49, "Ι"],
	[0x4a, "ϑ"],
	[0x4b, "Κ"],
	[0x4c, "Λ"],
	[0x4d, "Μ"],
	[0x4e, "Ν"],
	[0x4f, "Ο"],
	[0x50, "Π"],
	[0x51, "Θ"],
	[0x52, "Ρ"],
	[0x53, "Σ"],
	[0x54, "Τ"],
	[0x55, "Υ"],
	[0x56, "ς"],
	[0x57, "Ω"],
	[0x58, "Ξ"],
	[0x59, "Ψ"],
	[0x5a, "Ζ"],
	[0x5b, "["],
	[0x5c, "∴"],
	[0x5d, "]"],
	[0x5e, "⊥"],
	[0x5f, "_"],
	[0x60, "‾"],
	[0x61, "α"],
	[0x62, "β"],
	[0x63, "χ"],
	[0x64, "δ"],
	[0x65, "ε"],
	[0x66, "φ"],
	[0x67, "γ"],
	[0x68, "η"],
	[0x69, "ι"],
	[0x6a, "ϕ"],
	[0x6b, "κ"],
	[0x6c, "λ"],
	[0x6d, "μ"],
	[0x6e, "ν"],
	[0x6f, "ο"],
	[0x70, "π"],
	[0x71, "θ"],
	[0x72, "ρ"],
	[0x73, "σ"],
	[0x74, "τ"],
	[0x75, "υ"],
	[0x76, "ϖ"],
	[0x77, "ω"],
	[0x78, "ξ"],
	[0x79, "ψ"],
	[0x7a, "ζ"],
	[0x7b, "{"],
	[0x7c, "|"],
	[0x7d, "}"],
	[0x7e, "∼"],
	[0xa0, "€"],
	[0xa1, "ϒ"],
	[0xa2, "′"],
	[0xa3, "≤"],
	[0xa4, "⁄"],
	[0xa5, "∞"],
	[0xa6, "ƒ"],
	[0xa7, "♣"],
	[0xa8, "♦"],
	[0xa9, "♥"],
	[0xaa, "♠"],
	[0xab, "↔"],
	[0xac, "←"],
	[0xad, "↑"],
	[0xae, "→"],
	[0xaf, "↓"],
	[0xb0, "°"],
	[0xb1, "±"],
	[0xb2, "″"],
	[0xb3, "≥"],
	[0xb4, "×"],
	[0xb5, "∝"],
	[0xb6, "∂"],
	[0xb7, "•"],
	[0xb8, "÷"],
	[0xb9, "≠"],
	[0xba, "≡"],
	[0xbb, "≈"],
	[0xbc, "…"],
	[0xbd, "|"],
	[0xbe, "—"],
	[0xbf, "↵"],
	[0xc0, "ℵ"],
	[0xc1, "ℑ"],
	[0xc2, "ℜ"],
	[0xc3, "℘"],
	[0xc4, "⊗"],
	[0xc5, "⊕"],
	[0xc6, "∅"],
	[0xc7, "∩"],
	[0xc8, "∪"],
	[0xc9, "⊃"],
	[0xca, "⊇"],
	[0xcb, "⊄"],
	[0xcc, "⊂"],
	[0xcd, "⊆"],
	[0xce, "∈"],
	[0xcf, "∉"],
	[0xd0, "∠"],
	[0xd1, "∇"],
	[0xd2, "®"],
	[0xd3, "©"],
	[0xd4, "™"],
	[0xd5, "∏"],
	[0xd6, "√"],
	[0xd7, "·"],
	[0xd8, "¬"],
	[0xd9, "∧"],
	[0xda, "∨"],
	[0xdb, "⇔"],
	[0xdc, "⇐"],
	[0xdd, "⇑"],
	[0xde, "⇒"],
	[0xdf, "⇓"],
	[0xe0, "◊"],
	[0xe1, "⟨"],
	[0xe5, "∑"],
	[0xf1, "⟩"],
	[0xf2, "∫"],
];

/** Wingdings — Unicode equivalents for the most commonly used glyphs.
 * Codepoints are PUA (0xF020+); decodeSym also tries 0x20+ aliases. */
const WINGDINGS: ReadonlyArray<[number, string]> = [
	[0xf022, "✂"],
	[0xf027, "☎"],
	[0xf028, "✆"],
	[0xf029, "✉"],
	[0xf03d, "⌛"],
	[0xf04a, "☺"],
	[0xf04b, "😐"],
	[0xf04c, "☹"],
	[0xf04f, "☠"],
	[0xf050, "🏳"],
	[0xf051, "🏴"],
	[0xf058, "☜"],
	[0xf059, "☞"],
	[0xf05a, "☝"],
	[0xf05b, "☟"],
	[0xf06c, "❑"],
	[0xf06d, "❒"],
	[0xf06e, "▪"],
	[0xf06f, "□"],
	[0xf071, "◆"],
	[0xf073, "●"],
	[0xf074, "■"],
	[0xf075, "●"],
	[0xf076, "◆"],
	[0xf077, "❖"],
	[0xf0a7, "▪"],
	[0xf0a8, "▫"],
	[0xf0e0, "←"],
	[0xf0e1, "→"],
	[0xf0e2, "↑"],
	[0xf0e3, "↓"],
	[0xf0e4, "↖"],
	[0xf0e5, "↗"],
	[0xf0e6, "↘"],
	[0xf0e7, "↙"],
	[0xf0e8, "↔"],
	[0xf0e9, "↕"],
	[0xf0fa, "✓"],
	[0xf0fb, "✓"],
	[0xf0fc, "✓"],
	[0xf0fd, "✗"],
	[0xf0fe, "✗"],
];

/** Wingdings 2 — fewer common glyphs in document text. */
const WINGDINGS_2: ReadonlyArray<[number, string]> = [
	[0xf050, "▶"],
	[0xf051, "◀"],
	[0xf052, "▲"],
	[0xf053, "▼"],
	[0xf0d8, "❒"],
	[0xf0d9, "☐"],
	[0xf0fb, "☑"],
	[0xf0fc, "☒"],
	[0xf0fd, "✓"],
	[0xf0fe, "✗"],
	[0xf0ff, "✘"],
];

/** Webdings — web/computer-themed glyphs. */
const WEBDINGS: ReadonlyArray<[number, string]> = [
	[0xf021, "🕷"],
	[0xf024, "👁"],
	[0xf025, "👂"],
	[0xf032, "🌐"],
	[0xf039, "✉"],
	[0xf03c, "📅"],
	[0xf040, "🏠"],
	[0xf04a, "📞"],
	[0xf058, "🔍"],
	[0xf067, "🔒"],
	[0xf068, "🔓"],
	[0xf078, "✏"],
	[0xf0a4, "★"],
];

/** Zapf Dingbats — rich set of typographic ornaments and check/cross marks. */
const ZAPF_DINGBATS: ReadonlyArray<[number, string]> = [
	[0x21, "✁"],
	[0x22, "✂"],
	[0x23, "✃"],
	[0x24, "✄"],
	[0x2b, "✚"],
	[0x33, "✓"],
	[0x34, "✔"],
	[0x35, "✕"],
	[0x36, "✖"],
	[0x37, "✗"],
	[0x38, "✘"],
	[0x39, "✙"],
	[0x3a, "✚"],
	[0x3b, "✛"],
	[0x3c, "✜"],
	[0x4a, "❀"],
	[0x4b, "❁"],
	[0x4c, "❂"],
	[0x4d, "❃"],
	[0x4e, "❄"],
	[0x4f, "❅"],
	[0x6c, "●"],
	[0x6e, "■"],
	[0x71, "◆"],
];

const TABLES = new Map<string, Map<number, string>>([
	["symbol", new Map(SYMBOL)],
	["wingdings", new Map(WINGDINGS)],
	["wingdings 2", new Map(WINGDINGS_2)],
	["wingdings 3", new Map(WINGDINGS_2)],
	["webdings", new Map(WEBDINGS)],
	["zapfdingbats", new Map(ZAPF_DINGBATS)],
	["zapf dingbats", new Map(ZAPF_DINGBATS)],
	["itc zapf dingbats", new Map(ZAPF_DINGBATS)],
]);
