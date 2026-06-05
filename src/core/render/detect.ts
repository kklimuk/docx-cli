import { libreofficeEngine } from "./engines/libreoffice";
import type { RenderEngine, RenderEngineName } from "./engines/types";
import { wordMacEngine } from "./engines/word-mac";
import { wordWindowsEngine } from "./engines/word-windows";

/** Engines registered in priority order. Auto-selection (`--engine auto`)
 * picks the first one whose `available()` returns true on this machine —
 * Word is preferred over LibreOffice when both are present because Word's
 * render is the ground-truth most agents/users are targeting. Platform
 * gating (e.g., Word-mac only works on darwin) lives inside each engine's
 * `available()` check rather than here, so the order can stay flat. */
const ENGINES: readonly RenderEngine[] = [
	wordMacEngine,
	wordWindowsEngine,
	libreofficeEngine,
];

/** Look up an engine by name. Used by the CLI when the user passes
 * `--engine word|libreoffice` explicitly. Note: `word` resolves to whichever
 * platform-specific Word engine is appropriate (mac vs win) — there's no
 * "force Word-mac on Windows" path, that would never succeed. */
export function engineByName(
	name: RenderEngineName | "word",
): RenderEngine | undefined {
	if (name === "word") {
		return process.platform === "darwin"
			? wordMacEngine
			: process.platform === "win32"
				? wordWindowsEngine
				: undefined;
	}
	return ENGINES.find((engine) => engine.name === name);
}

/** Auto-select the highest-priority engine available on this machine. */
export async function detectEngine(): Promise<RenderEngine | undefined> {
	for (const engine of ENGINES) {
		if (await engine.available()) return engine;
	}
	return undefined;
}

/** Probe every engine and return the names of those available. Used for
 * the error message when nothing's available, to make the install hint
 * specific to the platform. */
export async function listAvailable(): Promise<RenderEngineName[]> {
	const out: RenderEngineName[] = [];
	for (const engine of ENGINES) {
		if (await engine.available()) out.push(engine.name);
	}
	return out;
}
