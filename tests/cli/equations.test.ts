import { describe, expect, test } from "bun:test";
import { copyFileSync } from "node:fs";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

/** Canonical equations fixture — see `tests/fixtures/setup/equations.ts`.
 *  Authored by the CLI's `insert --equation` pipeline (LaTeX → temml MathML →
 *  our OMML adapter), so every test here exercises the round-trip path end
 *  to end. */
const FIXTURE = join(import.meta.dir, "..", "fixtures", "equations.docx");

type EquationRun = {
	type: "equation";
	id: string;
	latex: string;
	text: string;
	display: boolean;
};

async function equationsOf(path: string): Promise<EquationRun[]> {
	const result = await runCli("read", path, "--ast");
	const doc = result.parsed as { blocks: AnyBlock[] };
	const out: EquationRun[] = [];
	collectEquations(doc.blocks, out);
	return out;
}

/** Recursive equation collector — descends into table cells so test cases
 *  exercising cell-resident equations (the table-cell regression test, for
 *  instance) can see what's there. Top-level `doc.blocks` doesn't include
 *  cell paragraphs. */
function collectEquations(blocks: AnyBlock[], out: EquationRun[]): void {
	for (const block of blocks) {
		for (const run of block.runs ?? []) {
			if (run.type === "equation") out.push(run as EquationRun);
		}
		for (const row of block.rows ?? []) {
			for (const cell of row.cells ?? []) {
				collectEquations(cell.blocks ?? [], out);
			}
		}
	}
}

type AnyBlock = {
	type?: string;
	runs?: Array<{ type: string }>;
	rows?: Array<{ cells?: Array<{ blocks?: AnyBlock[] }> }>;
};

describe("equation reader — AST shape", () => {
	test("ids are sequential `eqN` in document order; display flag matches m:oMathPara", async () => {
		const eqs = await equationsOf(FIXTURE);
		expect(eqs.length).toBeGreaterThan(60);
		expect(eqs[0]?.id).toBe("eq0");
		for (const [i, eq] of eqs.entries()) expect(eq.id).toBe(`eq${i}`);
		expect(eqs.filter((e) => e.display).length).toBeGreaterThan(0);
		expect(eqs.filter((e) => !e.display).length).toBeGreaterThan(0);
	});

	test("legacy `text` plaintext fallback stays populated alongside `latex`", async () => {
		const eqs = await equationsOf(FIXTURE);
		const eq = eqs.find((e) => e.latex === "x^2");
		expect(eq?.text).toBe("x2");
	});
});

describe("equation reader — atoms and decorations", () => {
	test("super/sub/mixed and Greek + accents", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("x^2");
		expect(haystack).toContain("a_i");
		expect(haystack).toContain("x_i^2");
		expect(haystack).toContain("E=mc^2");
		expect(haystack).toContain("\\alpha");
		expect(haystack).toContain("\\beta");
		expect(haystack).toContain("\\hat{n}");
		expect(haystack).toContain("\\bar{x}");
		expect(haystack).toContain("\\vec{v}");
	});

	test("wide / stretchy accents map to their `\\wide*` commands", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\widehat{x}");
		expect(haystack).toContain("\\widetilde{x}");
		expect(haystack).toContain("\\overrightarrow{AB}");
	});

	test("bar / underline / boxed decorations", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\overline{x+y}");
		expect(haystack).toContain("\\underline{x+y}");
		expect(haystack).toContain("\\boxed{x+y}");
	});

	test("labeled braces preserve the subscript/superscript label", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\underbrace{a+b+c}_{=n}");
		expect(haystack).toContain("\\overbrace{a+b+c}^{n}");
	});
});

describe("equation reader — fractions, roots, big operators", () => {
	test("fractions and \\binom (the noBar fraction wrapped in `<m:d>`)", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\frac{a}{b}");
		expect(haystack).toContain("\\frac{\\partial f}{\\partial x}");
		expect(haystack).toContain("\\frac{1}{1+\\frac{1}{x}}"); // nested
		expect(haystack).toContain("\\binom{n}{k}");
	});

	test("roots (square, nested, nth)", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\sqrt{2}");
		expect(haystack).toContain("\\sqrt{a^2+b^2}");
		expect(haystack).toContain("\\sqrt[3]{x+y}");
		expect(haystack).toContain("\\sqrt[n]{a_1a_2\\cdots a_n}");
	});

	test("n-ary operators recover sub+sup bounds + body", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\sum_{i=1}^{n}");
		expect(haystack).toContain("\\prod_{i=1}^{n}");
		expect(haystack).toContain("\\int_{0}^{\\infty}");
		expect(haystack).toContain("\\iint_{D}");
		expect(haystack).toContain("\\oint_{C}");
	});

	test("\\lim_{x→0} promotes the function name and keeps the limit subscript", async () => {
		const eqs = await equationsOf(FIXTURE);
		const lim = eqs.find((e) => e.latex.includes("\\lim_{x\\to 0}"));
		expect(lim).toBeDefined();
		expect(lim?.latex).toContain("\\frac{\\sin x}{x}");
	});
});

describe("equation reader — function application (`\\ln`, `\\sin`, `\\log_{10}`)", () => {
	test("`\\ln`, `\\sin`, `\\log` round-trip cleanly with single-space arg separator", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		// Gibbs free energy: `\Delta G = -RT \ln K`
		expect(haystack).toContain("\\ln K");
		// Nernst: `... \ln Q`
		expect(haystack).toContain("\\ln Q");
		// Half-life: `t_{1/2} = \frac{\ln 2}{k}`
		expect(haystack).toContain("\\ln 2");
		// Limit: `\lim_{x\to 0} \frac{\sin x}{x}`
		expect(haystack).toContain("\\sin x");
		// Henderson-Hasselbalch: `\log_{10} \frac{...}{...}`
		expect(haystack).toContain("\\log_{10}");
	});
});

describe("equation reader — delimiters, norms, matrices", () => {
	test("norm `\\|x\\|` survives via `<m:d>` with `‖` delimiters", async () => {
		const eqs = await equationsOf(FIXTURE);
		// User-typed `\|x\|` reads back as `\left\|x\right\|` — the canonical
		// stretchy-delimiter form for OMML's `<m:d>`. Both parse to identical
		// MathML in temml.
		const norm = eqs.find((e) => e.latex.includes("\\left\\|x\\right\\|"));
		expect(norm).toBeDefined();
	});

	test("`\\left(\\frac{a}{b}\\right)` keeps the stretchy delimiters", async () => {
		const eqs = await equationsOf(FIXTURE);
		const delim = eqs.find((e) =>
			e.latex.includes("\\left(\\frac{a}{b}\\right)"),
		);
		expect(delim).toBeDefined();
	});

	test("`\\begin{pmatrix}` shorthand (display) — `(`/`)` pair → pmatrix env", async () => {
		const eqs = await equationsOf(FIXTURE);
		const pmatrix = eqs.find((e) =>
			e.latex.includes("\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}"),
		);
		expect(pmatrix).toBeDefined();
		expect(pmatrix?.display).toBe(true);
	});

	test("aligned env (display) `<m:eqArr>` round-trips", async () => {
		const eqs = await equationsOf(FIXTURE);
		const aligned = eqs.find((e) => e.latex.startsWith("\\begin{aligned}"));
		expect(aligned).toBeDefined();
		expect(aligned?.display).toBe(true);
		expect(aligned?.latex).toContain("\\\\");
	});

	test("cases env (display) collapses to `\\left\\{...matrix...\\right.` (no native OMML cases)", async () => {
		const eqs = await equationsOf(FIXTURE);
		const cases = eqs.find((e) => e.latex.includes("\\left\\{"));
		expect(cases).toBeDefined();
		expect(cases?.latex).toContain("x\\geq 0");
	});
});

describe("equation reader — famous formulas across statistics, engineering, chemistry", () => {
	test("statistics: mean, variance, normal PDF, Bayes, binomial PMF", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		// Round-trip is structurally exact; whitespace and brace-wrapping may
		// differ from the natural-LaTeX input (both parse to identical OMML).
		expect(haystack).toContain("\\sum_{i=1}^{n}");
		expect(haystack).toContain("(x_i-\\mu)"); // variance
		expect(haystack).toContain("\\sigma \\sqrt{2\\pi}"); // normal pdf
		expect(haystack).toContain("\\frac{P(B|A)P(A)}{P(B)}"); // bayes
		expect(haystack).toContain("\\binom{n}{k}p^k"); // binomial PMF base
	});

	test("engineering: Newton gravity, Schrödinger, Euler identity, Maxwell-Ampère, Bernoulli", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("G\\frac{m_1m_2}{r^2}"); // newton-gravity
		expect(haystack).toContain("\\hbar"); // Schrödinger
		expect(haystack).toContain("\\partial \\Psi");
		expect(haystack).toContain("e^{i\\pi}+1=0"); // euler
		expect(haystack).toContain("\\nabla \\times \\vec{B}"); // maxwell-ampere
		expect(haystack).toContain("\\frac{1}{2}\\rho v^2"); // bernoulli
	});

	test("chemistry: pH, Arrhenius, ideal gas, Gibbs, Henderson-Hasselbalch", async () => {
		const eqs = await equationsOf(FIXTURE);
		const haystack = eqs.map((e) => e.latex).join("\n");
		expect(haystack).toContain("\\text{pH}");
		expect(haystack).toContain("Ae^{-E_a/RT}"); // arrhenius
		expect(haystack).toContain("PV=nRT"); // ideal-gas
		expect(haystack).toContain("\\Delta G=-RT\\ln K"); // gibbs
		expect(haystack).toContain("pK_a+\\log_{10}"); // henderson-hasselbalch
	});
});

describe("eqN locator", () => {
	test("`eqN` parses via the locator grammar (wc rejects with a clear message)", async () => {
		const result = await runCli("wc", FIXTURE, "eq0");
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("addresses a equation, not text");
	});
});

describe("markdown render", () => {
	test("inline equations emit `$…$`", async () => {
		const result = await runCli("read", FIXTURE);
		expect(result.stdout).toContain("$x^2$");
		expect(result.stdout).toContain("$\\frac{a}{b}$");
	});

	test("display equations emit `$$…$$` on their own line", async () => {
		const result = await runCli("read", FIXTURE);
		// Quadratic formula round-trips to its canonical display form.
		expect(result.stdout).toContain("$$x=\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}$$");
	});
});

// ---------------------------------------------------------------------------
// Write side: `insert --equation` and `edit --at eqN --equation`
// ---------------------------------------------------------------------------

describe("insert --equation", () => {
	test("inline LaTeX round-trips: insert x^2, read back $x^2$", async () => {
		const workspace = tempWorkspace("eq-insert-inline");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"x^2 + y^2 = r^2",
		);
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$x^2+y^2=r^2$");
	});

	test("display equations emit `$$…$$` on their own line", async () => {
		const workspace = tempWorkspace("eq-insert-display");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"\\frac{a}{b}",
			"--display",
		);
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$$\\frac{a}{b}$$");
	});

	test("Tier 1+2+3 constructs all round-trip structurally", async () => {
		const workspace = tempWorkspace("eq-insert-tiers");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const inputs = [
			"\\binom{n}{k}",
			"\\boxed{x+y}",
			"\\overline{x+y}",
			"\\widehat{X}",
			"\\overrightarrow{AB}",
			"\\underbrace{a+b}_{=n}",
			"\\|x\\|",
			"\\ln K",
			"\\log_{10} N",
			"\\sin x",
		];
		for (let i = 0; i < inputs.length; i++) {
			await runCli(
				"insert",
				docPath,
				"--after",
				`p${i}`,
				"--equation",
				inputs[i] as string,
			);
		}
		const eqs = await equationsOf(docPath);
		const haystack = eqs.map((e) => e.latex).join("\n");
		for (const probe of [
			"\\binom{n}{k}",
			"\\boxed{x+y}",
			"\\overline{x+y}",
			"\\widehat{X}",
			"\\overrightarrow{AB}",
			"\\underbrace{a+b}_{=n}",
			"\\left\\|x\\right\\|", // norm canonicalizes to the stretchy form
			"\\ln K",
			"\\log_{10} N",
			"\\sin x",
		]) {
			expect(haystack).toContain(probe);
		}
	});

	test("malformed LaTeX fails with a clear USAGE error", async () => {
		const workspace = tempWorkspace("eq-insert-bad");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"\\frac{a}{",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("Could not parse LaTeX");
	});

	test("--equation is mutex with --text/--code/--task etc.", async () => {
		const workspace = tempWorkspace("eq-insert-mutex");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"x",
			"--text",
			"y",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("only one of");
	});
});

describe("edit --at eqN --equation", () => {
	test("replaces an equation's content; locator addresses by document order", async () => {
		const workspace = tempWorkspace("eq-edit-content");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		// eq0 in the fixture is `x^2`. Replace with `x^3`.
		await runCli("edit", docPath, "--at", "eq0", "--equation", "x^3");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$x^3$");
	});

	test("--display toggles existing inline equation to display mode", async () => {
		const workspace = tempWorkspace("eq-edit-display");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		await runCli("edit", docPath, "--at", "eq0", "--display");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$$x^2$$");
	});

	test("--inline toggles existing display equation to inline mode", async () => {
		const workspace = tempWorkspace("eq-edit-inline");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header", "--force");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"x^2",
			"--display",
		);
		await runCli("edit", docPath, "--at", "eq0", "--inline");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$x^2$");
		expect(result.stdout).not.toContain("$$");
	});

	test("--display works for equations inside table cells (regression)", async () => {
		// Without the `latex` field on EquationReference, the mode-toggle
		// failed for cell-resident equations: the resolver succeeds via
		// `equationReferences.get(...)` but the lookup that walked
		// `view.body.blocks` skipped table cells, yielding a misleading
		// "LaTeX wasn't readable" error. Insert into a cell paragraph and
		// confirm --display flips inline → display via the cached `latex`.
		const workspace = tempWorkspace("eq-edit-cell-display");
		const docPath = join(workspace, "out.docx");
		const tablesFixture = join(
			import.meta.dir,
			"..",
			"fixtures",
			"tables.docx",
		);
		copyFileSync(tablesFixture, docPath);
		await runCli(
			"insert",
			docPath,
			"--after",
			"t0:r0c0:p0",
			"--equation",
			"x^2",
		);
		const insertedEqId = (await equationsOf(docPath))[0]?.id;
		expect(insertedEqId).toBe("eq0");
		const result = await runCli("edit", docPath, "--at", "eq0", "--display");
		expect(result.exitCode).toBe(0);
		const after = await equationsOf(docPath);
		expect(after[0]?.display).toBe(true);
		expect(after[0]?.latex).toBe("x^2");
	});

	test("rejects a stale eqN locator", async () => {
		const workspace = tempWorkspace("eq-edit-missing");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"eq999",
			"--equation",
			"x",
		);
		expect(result.exitCode).toBe(3); // BLOCK_NOT_FOUND
		expect(result.stdout).toContain("Equation not found");
	});

	test("--equation in a pN-pM range is rejected", async () => {
		const workspace = tempWorkspace("eq-edit-range");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1-p3",
			"--equation",
			"x",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("takes a single equation locator");
	});
});

// ---------------------------------------------------------------------------
// Tier 1 — Tracked equation insert/delete (Phase 3 partial).
// The wrappers in applyTrackedInsertion / applyTrackedDeletion must cover
// `<m:oMath>` and `<m:oMathPara>` siblings, not just `<w:r>` runs, so the
// equation itself is marked as inserted/deleted. Verified end-to-end via
// Microsoft Word for Mac's accept/reject pipeline (see commit message).
// ---------------------------------------------------------------------------

describe("tracked equation insert/delete (Tier 1)", () => {
	test("insert --equation under tracking wraps <m:oMath> in <w:ins>", async () => {
		const workspace = tempWorkspace("eq-tracked-insert-inline");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("track-changes", docPath, "on");
		await runCli("insert", docPath, "--after", "p0", "--equation", "E=mc^2");

		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:ins\b[^>]*>\s*<m:oMath\b/);

		// Accept-all in our pipeline keeps the equation (paragraph mark + content).
		const acceptResult = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--all",
		);
		expect(acceptResult.exitCode).toBe(0);
		const afterAccept = await runCli("read", docPath);
		expect(afterAccept.stdout).toContain("$E=mc^2$");
	});

	test("insert --equation --display under tracking wraps <m:oMathPara>", async () => {
		const workspace = tempWorkspace("eq-tracked-insert-display");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--equation",
			"\\frac{a}{b}",
			"--display",
		);
		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:ins\b[^>]*>\s*<m:oMathPara\b/);
	});

	test("delete an equation paragraph under tracking wraps OMML in <w:del>", async () => {
		const workspace = tempWorkspace("eq-tracked-delete");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Keep this.");
		await runCli("insert", docPath, "--after", "p0", "--equation", "E=mc^2");
		await runCli("track-changes", docPath, "on");
		await runCli("delete", docPath, "--at", "p1");

		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:del\b[^>]*>\s*<m:oMath\b/);

		// Rejecting the tracked delete restores the equation paragraph.
		const rejectResult = await runCli(
			"track-changes",
			"reject",
			docPath,
			"--all",
		);
		expect(rejectResult.exitCode).toBe(0);
		const afterReject = await runCli("read", docPath);
		expect(afterReject.stdout).toContain("$E=mc^2$");
	});

	test("track-changes list surfaces the equation insertion as one tc entry", async () => {
		const workspace = tempWorkspace("eq-tracked-list");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("track-changes", docPath, "on");
		await runCli("insert", docPath, "--after", "p0", "--equation", "x^2");
		const result = await runCli("track-changes", "list", docPath);
		const list = result.parsed as Array<{ kind: string; blockId: string }>;
		const inserts = list.filter((entry) => entry.kind === "ins");
		expect(inserts.length).toBeGreaterThan(0);
		// At least one of the entries should point at the new paragraph (p1).
		expect(inserts.some((entry) => entry.blockId === "p1")).toBe(true);
	});
});

// ---------------------------------------------------------------------------
// Tier 2 — Tracked equation content / mode edits (Phase 3 main path).
// `edit --at eqN --equation NEW` under tracking emits the OLD OMML wrapped in
// `<w:del>` next to the NEW OMML wrapped in `<w:ins>`. Our own track-changes
// accept/reject handles it cleanly. Word's accept/reject picks the right
// equation semantically; Word leaves a small structural skeleton next to
// the kept equation as a normalization quirk (cosmetic only).
// ---------------------------------------------------------------------------

describe("tracked equation edit (Tier 2)", () => {
	test("--equation NEW under tracking emits <w:del>OLD + <w:ins>NEW", async () => {
		const workspace = tempWorkspace("eq-tracked-edit-content");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("insert", docPath, "--after", "p0", "--equation", "x^2");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"eq0",
			"--equation",
			"y^3",
		);
		expect(result.exitCode).toBe(0);
		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:del\b[^>]*>\s*<m:oMath\b/);
		expect(xml).toMatch(/<w:ins\b[^>]*>\s*<m:oMath\b/);

		const list = (await runCli("track-changes", "list", docPath))
			.parsed as Array<{ kind: string }>;
		expect(list.some((entry) => entry.kind === "del")).toBe(true);
		expect(list.some((entry) => entry.kind === "ins")).toBe(true);
	});

	test("our accept-all picks the new equation", async () => {
		const workspace = tempWorkspace("eq-tracked-edit-accept");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("insert", docPath, "--after", "p0", "--equation", "x^2");
		await runCli("track-changes", docPath, "on");
		await runCli("edit", docPath, "--at", "eq0", "--equation", "y^3");
		await runCli("track-changes", "accept", docPath, "--all");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$y^3$");
		expect(result.stdout).not.toContain("$x^2$");
	});

	test("our reject-all keeps the original equation", async () => {
		const workspace = tempWorkspace("eq-tracked-edit-reject");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("insert", docPath, "--after", "p0", "--equation", "x^2");
		await runCli("track-changes", docPath, "on");
		await runCli("edit", docPath, "--at", "eq0", "--equation", "y^3");
		await runCli("track-changes", "reject", docPath, "--all");
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("$x^2$");
		expect(result.stdout).not.toContain("$y^3$");
	});

	test("--display toggle under tracking emits paired wrap (mode change)", async () => {
		const workspace = tempWorkspace("eq-tracked-mode-toggle");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Header");
		await runCli("insert", docPath, "--after", "p0", "--equation", "x^2");
		await runCli("track-changes", docPath, "on");
		await runCli("edit", docPath, "--at", "eq0", "--display");
		const xml = await readDocumentXml(docPath);
		expect(xml).toMatch(/<w:del\b[^>]*>\s*<m:oMath\b/);
		expect(xml).toMatch(/<w:ins\b[^>]*>\s*<m:oMathPara\b/);
	});

	test("stale eqN locator rejected under tracking same as untracked", async () => {
		const workspace = tempWorkspace("eq-tracked-edit-stale");
		const docPath = join(workspace, "out.docx");
		copyFileSync(FIXTURE, docPath);
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"eq999",
			"--equation",
			"x",
		);
		expect(result.exitCode).toBe(3);
		expect(result.stdout).toContain("Equation not found");
	});
});

async function readDocumentXml(docPath: string): Promise<string> {
	const JSZip = (await import("jszip")).default;
	const zip = await JSZip.loadAsync(await Bun.file(docPath).bytes());
	const file = zip.file("word/document.xml");
	if (!file) throw new Error("document.xml missing from docx");
	return await file.async("string");
}
