// Regenerates THIRD-PARTY-LICENSES: the verbatim LICENSE text of every package in
// docx-cli's production dependency closure, the attribution the MIT/BSD/ISC/Zlib
// licenses require when we redistribute them inside dist/index.js and the compiled
// binaries. Run `bun run licenses` after a dependency change; the release drift
// guard in tests/cli/installer.test.ts fails while the committed file is stale.
//
// It refuses (throws) when a package ships no license file or declares a license
// outside PERMISSIVE_LICENSES, so a copyleft dependency can't slip in unnoticed.
// Versions are deliberately left out of the output, so a version bump only shows
// up here when the license text itself changed.

import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";

/** The full THIRD-PARTY-LICENSES text for the production closure rooted at the
 *  repo's package.json `dependencies`. Deterministic: packages sort by name. */
export function buildThirdPartyLicenses(root: string): string {
	const packages = collectProductionPackages(root);
	const sections = [...packages.values()]
		.sort((left, right) => left.name.localeCompare(right.name))
		.map(formatSection);
	return `${HEADER}\n${sections.join("\n")}`;
}

/** Every package in the production closure, keyed by name. Two installed
 *  versions of one package keep the first one met; their license texts are
 *  compared so a relicense between versions can't hide behind the dedupe. */
export function collectProductionPackages(
	root: string,
): Map<string, LicensedPackage> {
	const rootManifest = readManifest(join(root, "package.json"));
	const packages = new Map<string, LicensedPackage>();
	const visited = new Set<string>();
	const pending = Object.keys(rootManifest.dependencies ?? {}).map((name) => ({
		name,
		fromDir: root,
		optional: false,
	}));
	for (let next = pending.pop(); next; next = pending.pop()) {
		const manifestPath = resolveManifest(next.name, next.fromDir);
		// A platform-specific optional dependency is legitimately absent on other
		// platforms; anything else missing means node_modules is stale.
		if (!manifestPath && next.optional) continue;
		if (!manifestPath) {
			throw new Error(
				`Cannot resolve production dependency "${next.name}" — run bun install`,
			);
		}
		if (visited.has(manifestPath)) continue;
		visited.add(manifestPath);
		const manifest = readManifest(manifestPath);
		const licensed = readLicensedPackage(manifest, dirname(manifestPath));
		const known = packages.get(licensed.name);
		if (known && known.licenseText !== licensed.licenseText) {
			throw new Error(
				`${licensed.name}: installed versions carry different license texts — attribute both before regenerating`,
			);
		}
		if (!known) packages.set(licensed.name, licensed);
		const fromDir = dirname(manifestPath);
		for (const name of Object.keys(manifest.dependencies ?? {})) {
			pending.push({ name, fromDir, optional: false });
		}
		for (const name of Object.keys(manifest.optionalDependencies ?? {})) {
			pending.push({ name, fromDir, optional: true });
		}
	}
	return packages;
}

/** Node-style lookup: `node_modules/<name>/package.json` in `fromDir` or the
 *  nearest ancestor that has it (how Bun's installed tree resolves). */
function resolveManifest(name: string, fromDir: string): string | undefined {
	for (let dir = fromDir; ; dir = dirname(dir)) {
		const candidate = join(dir, "node_modules", name, "package.json");
		if (existsSync(candidate)) return candidate;
		if (dirname(dir) === dir) return undefined;
	}
}

function readManifest(path: string): PackageManifest {
	return JSON.parse(readFileSync(path, "utf8")) as PackageManifest;
}

function readLicensedPackage(
	manifest: PackageManifest,
	packageDir: string,
): LicensedPackage {
	const license = declaredLicense(manifest);
	if (!PERMISSIVE_LICENSES.has(license)) {
		throw new Error(
			`${manifest.name}: license "${license}" is not on the permissive allowlist — review it before adding the dependency`,
		);
	}
	const licenseText = readLicenseText(manifest, license, packageDir)
		.replace(/\r\n/g, "\n")
		.trim();
	return { name: manifest.name, license, licenseText };
}

/** The package's LICENSE file; else the README's License section when it holds
 *  the full permission notice; else, for a bare MIT declaration with no text
 *  anywhere in the package, the MIT terms under the manifest's author (stated
 *  as such). Any other license with no text is refused. */
function readLicenseText(
	manifest: PackageManifest,
	license: string,
	packageDir: string,
): string {
	const entries = readdirSync(packageDir).sort();
	const licenseFile = entries.find((entry) => LICENSE_FILE_PATTERN.test(entry));
	if (licenseFile) return readFileSync(join(packageDir, licenseFile), "utf8");
	const readme = entries.find((entry) => /^readme\.md$/i.test(entry));
	const readmeSection = readme
		? readmeLicenseSection(readFileSync(join(packageDir, readme), "utf8"))
		: undefined;
	if (readmeSection?.includes("Permission is hereby granted")) {
		return `(No LICENSE file ships with this package; quoted from its README.)\n\n${readmeSection}`;
	}
	if (license === "MIT" && manifest.author) {
		return `(No LICENSE file ships with this package; package.json declares "MIT" with author "${authorName(manifest.author)}". The MIT License terms follow.)\n\nCopyright (c) ${authorName(manifest.author)}\n\n${MIT_TERMS}`;
	}
	throw new Error(`${manifest.name}: ships no license text to reproduce`);
}

function readmeLicenseSection(readme: string): string | undefined {
	const match = readme.match(
		/^#+ *licen[cs]e *\n([\s\S]*?)(?=^#+ |^\[|$(?![\s\S]))/im,
	);
	return match?.[1]?.trim();
}

function authorName(author: string | { name?: string }): string {
	return typeof author === "string" ? author : (author.name ?? "");
}

function declaredLicense(manifest: PackageManifest): string {
	if (typeof manifest.license === "string") return manifest.license;
	if (manifest.license?.type) return manifest.license.type;
	return "UNKNOWN";
}

function formatSection(licensed: LicensedPackage): string {
	return `${RULE}\n${licensed.name} — ${licensed.license}\n${RULE}\n\n${licensed.licenseText}\n`;
}

/** SPDX expressions accepted as-is. jszip's dual license is permitted because
 *  NOTICES elects its MIT option. Anything else must be reviewed by a human,
 *  added here, and attributed in NOTICES if it needs more than its LICENSE text. */
const PERMISSIVE_LICENSES = new Set([
	"MIT",
	"ISC",
	"BSD-2-Clause",
	"BSD-3-Clause",
	"Apache-2.0",
	"0BSD",
	"(MIT AND Zlib)",
	"(MIT OR GPL-3.0-or-later)",
]);

const LICENSE_FILE_PATTERN = /^(licen[cs]e|copying)(\.(md|markdown|txt))?$/i;

const MIT_TERMS = `Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.`;

export const THIRD_PARTY_LICENSES_FILE = "THIRD-PARTY-LICENSES";

const RULE = "=".repeat(80);

const HEADER = `docx-cli — third-party license texts

Generated by \`bun run licenses\` (scripts/third-party-licenses.ts) from the
installed production dependency closure. Do not edit by hand.

Every package below is redistributed inside dist/index.js and the compiled
binaries. Each section reproduces that package's own LICENSE file. NOTICES
carries the notices that are not in these files (the PDFium binary, the libxml2
C sources compiled into libxml2-wasm, and the Bun runtime) and the license
elections docx-cli makes.
`;

type PackageManifest = {
	name: string;
	version: string;
	license?: string | { type?: string };
	author?: string | { name?: string };
	dependencies?: Record<string, string>;
	optionalDependencies?: Record<string, string>;
};

type LicensedPackage = { name: string; license: string; licenseText: string };

// Last, not first: the constants above must be initialized before it runs.
if (import.meta.main) {
	const root = join(import.meta.dir, "..");
	await Bun.write(
		join(root, THIRD_PARTY_LICENSES_FILE),
		buildThirdPartyLicenses(root),
	);
}
