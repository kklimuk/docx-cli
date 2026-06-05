import { basename, extname, resolve } from "node:path";
import {
	detectEngine,
	engineByName,
	listAvailable,
	type RenderEngine,
	RenderEngineError,
	type RenderEngineName,
	renderDocxPages,
} from "@core";
import { EXIT, fail, respond, tryParseArgs, writeStdout } from "../respond";
import { parsePagesSpec } from "./parse-pages";

const HELP = `docx render — render each page of a .docx as a PNG/JPG image

Usage:
  docx render FILE [options]

Why: ground-truth visual verification. The CLI is primarily consumed by AI
agents that can read the produced images and reason about what the document
actually looks like (heading sizes, list indentation, page breaks, etc.).

Options:
  --out DIR         Directory to write images into (default: ./{basename}-pages/).
                    Created if missing; existing files in it are NOT removed
                    (so you can re-render a subset and keep prior pages).
  --engine ENGINE   word | libreoffice | auto (default: auto)
                      word         — Microsoft Word, the ground-truth renderer.
                                     macOS: drives Word for Mac via osascript,
                                     staging the docx inside Word's sandbox
                                     Container. Windows: drives Word via
                                     PowerShell COM. Linux: not supported.
                      libreoffice  — soffice --headless --convert-to pdf.
                                     Works on macOS / Linux / Windows.
                      auto         — picks the highest-fidelity engine
                                     available on this machine (Word > LO).
  --dpi N           Pixels per inch (default: 150). 72 = Word's default
                    print render; 150 = legible at 100% zoom in any
                    standard image viewer; 300 = print quality.
  --pages SPEC      Subset, e.g. "1" or "1-3" (default: all). Discontinuous
                    ranges (e.g., "1,3,5") aren't supported yet — run the
                    command multiple times for those.
  --format FMT      png (default) | jpg
  -v, --verbose     Print the full success ack JSON (default: silent on
                    success — but render always prints the page list since
                    the agent can't reconstruct them otherwise).
  -h, --help        Show this help

Examples:
  docx render report.docx
  docx render report.docx --out ./snapshots --engine libreoffice --dpi 200
  docx render report.docx --pages 1-3 --format jpg

Runtime dependencies:
  - Word engine: Microsoft Word installed locally (macOS or Windows).
    First run on macOS triggers a one-time Automation permission prompt.
  - LibreOffice engine: soffice on PATH (or installed at the default location).
  - PDF rasterization is built in via the bundled @hyzyla/pdfium WASM package
    — no extra system tools (poppler / pdftoppm / ImageMagick) required.

Run "docx render --help" or "docx --help" for the full command list.
`;

const OPTION_SPEC = {
	out: { type: "string" },
	engine: { type: "string" },
	dpi: { type: "string" },
	pages: { type: "string" },
	format: { type: "string" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);
	if (!(await Bun.file(filePath).exists())) {
		return fail("FILE_NOT_FOUND", `File not found: ${filePath}`);
	}

	const engineChoice = await resolveEngine(
		parsed.values.engine as string | undefined,
	);
	if (typeof engineChoice === "number") return engineChoice;

	const dpiRaw = parsed.values.dpi as string | undefined;
	const dpi = dpiRaw !== undefined ? Number(dpiRaw) : 150;
	if (!Number.isFinite(dpi) || dpi < 36 || dpi > 600) {
		return fail(
			"USAGE",
			`--dpi must be a number between 36 and 600, got "${dpiRaw}"`,
		);
	}

	const formatRaw = parsed.values.format as string | undefined;
	const format = formatRaw ?? "png";
	if (format !== "png" && format !== "jpg") {
		return fail("USAGE", `--format must be png or jpg, got "${formatRaw}"`);
	}

	const pagesRaw = parsed.values.pages as string | undefined;
	const range = pagesRaw !== undefined ? parsePagesSpec(pagesRaw) : undefined;
	if (typeof range === "string") return fail("USAGE", range);

	const outDir = resolve(
		(parsed.values.out as string | undefined) ??
			`./${basename(filePath, extname(filePath))}-pages`,
	);

	try {
		const result = await renderDocxPages(resolve(filePath), {
			engine: engineChoice,
			outDir,
			dpi,
			format,
			range,
		});
		// Always print the page list (not gated on --verbose): agents need
		// it to know which files to read. Same pattern as `comments add
		// --batch` printing minted ids — the output isn't reconstructable
		// downstream. The ack uses `output` (not `outDir`) because that's
		// the public field name agents are coded against.
		await respond({
			ok: true,
			operation: "render",
			path: filePath,
			engine: result.engine,
			output: result.outDir,
			pages: result.pages,
		});
		return EXIT.OK;
	} catch (error) {
		if (error instanceof RenderEngineError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}
}

async function resolveEngine(
	choice: string | undefined,
): Promise<RenderEngine | number> {
	const requested = choice ?? "auto";
	if (requested === "auto") {
		const engine = await detectEngine();
		if (engine) return engine;
		const installed = await listAvailable();
		return fail(
			"RENDER_ENGINE",
			"No render engine detected on this machine.",
			availabilityHint(installed),
		);
	}
	if (
		requested !== "word" &&
		requested !== "word-mac" &&
		requested !== "word-win" &&
		requested !== "libreoffice"
	) {
		return fail(
			"USAGE",
			`--engine must be auto, word, or libreoffice (got "${requested}")`,
		);
	}
	const engine = engineByName(requested as RenderEngineName | "word");
	if (!engine) {
		return fail(
			"RENDER_ENGINE",
			`Engine "${requested}" is not available on platform ${process.platform}`,
			requested === "word"
				? "Word engine requires macOS or Windows; use --engine libreoffice on Linux."
				: undefined,
		);
	}
	if (!(await engine.available())) {
		return fail(
			"RENDER_ENGINE",
			`Engine "${requested}" is not installed or not on PATH`,
			engine.name === "libreoffice"
				? "Install via `brew install --cask libreoffice` (macOS), `apt install libreoffice` (Linux), or libreoffice.org (Windows)."
				: engine.name === "word-mac"
					? "Microsoft Word for Mac required. Confirm the app is in /Applications and has been launched at least once."
					: "Microsoft Word for Windows required.",
		);
	}
	return engine;
}

function availabilityHint(installed: RenderEngineName[]): string {
	if (installed.length === 0) {
		const platform = process.platform;
		return platform === "darwin"
			? "macOS: install Microsoft Word, or `brew install --cask libreoffice`. PDF rasterization is bundled — no other tools needed."
			: platform === "win32"
				? "Windows: install Microsoft Word, or download LibreOffice from libreoffice.org. PDF rasterization is bundled — no other tools needed."
				: "Linux: install LibreOffice via your package manager. PDF rasterization is bundled — no other tools needed.";
	}
	return `Detected engines: ${installed.join(", ")} — but none chose auto-select; pass --engine explicitly.`;
}
