import type { Pkg } from "@core/ast/document/package";
import {
	type ConsistencyIssue,
	checkPackageConsistency,
	type PackagePart,
} from "@core/raw/package-consistency";
import {
	type ValidationIssue,
	validateWmlXml,
	validationXmlFor,
	WML_ROOT_TAGS,
} from "@core/raw/validate";

/** Validate a package: the ECMA-376 schema pass over every WML part plus the
 *  cross-part consistency audit, each finding filed under the part it
 *  concerns with its severity. `strict` is set (and nothing validated) for an
 *  ISO-Strict package — the transitional schemas can't judge it. */
export async function validatePackage(pkg: Pkg): Promise<Report> {
	const xmlParts = await collectXmlParts(pkg);
	const parts = xmlParts.filter((part) => WML_ROOT_TAGS.has(part.root.tag));
	if (isStrictProfile(parts)) {
		return {
			strict: true,
			checked: 0,
			parts: [],
			errors: 0,
			warnings: 0,
			fixable: 0,
		};
	}
	const results: PartReport[] = parts.map((part) => ({
		name: part.name,
		issues: validateWmlXml(validationXmlFor(part.root)).map(
			(issue): ReportedIssue => ({ ...issue, severity: "error" }),
		),
	}));
	mergeConsistencyIssues(results, checkPackageConsistency(xmlParts));
	const issues = results.flatMap((part) => part.issues);
	return {
		strict: false,
		checked: parts.length,
		parts: results,
		errors: issues.filter((issue) => issue.severity === "error").length,
		warnings: issues.filter((issue) => issue.severity === "warning").length,
		fixable: issues.filter((issue) => issue.fixable).length,
	};
}

/** Every XML package part, parsed once — the parse feeds the strict-profile
 *  sniff, schema validation (WML roots only), and the consistency audit (which
 *  also reads non-WML side-parts like commentsIds.xml). Sorted with
 *  word/document.xml first (the part a reader looks for), the rest by name. */
async function collectXmlParts(pkg: Pkg): Promise<PackagePart[]> {
	const names = pkg
		.listParts()
		.filter((name) => name.endsWith(".xml"))
		.sort((left, right) => {
			if (left === "word/document.xml") return -1;
			if (right === "word/document.xml") return 1;
			return left.localeCompare(right);
		});
	const parts: PackagePart[] = [];
	for (const name of names) {
		const tree = await pkg.readPart(name);
		const root = tree?.find((node) => !node.isText && node.tag !== "?xml");
		if (root) parts.push({ name, root });
	}
	return parts;
}

/** ISO Strict re-namespaces every part (`http://purl.oclc.org/ooxml/…`), so the
 *  transitional schemas see unknown content and every finding would be noise. */
function isStrictProfile(parts: PackagePart[]): boolean {
	const documentPart = parts.find((part) => part.name === "word/document.xml");
	return (
		documentPart?.root
			.getAttribute("xmlns:w")
			?.startsWith("http://purl.oclc.org/ooxml") ?? false
	);
}

/** File each consistency finding under the part it concerns — appended to
 *  that part's schema entry, or a new entry for a part the schema pass skips
 *  (commentsIds.xml has no wml.xsd root). */
function mergeConsistencyIssues(
	results: PartReport[],
	issues: ConsistencyIssue[],
): void {
	for (const issue of issues) {
		let entry = results.find((result) => result.name === issue.part);
		if (!entry) {
			entry = { name: issue.part, issues: [] };
			results.push(entry);
		}
		entry.issues.push({
			message: issue.message,
			kind: "consistency",
			severity: issue.severity,
			...(issue.fixable ? { fixable: true } : {}),
			...(issue.hint ? { hint: issue.hint } : {}),
		});
	}
}

/** The `--json` payload's report fields: `valid` is "no errors" (warnings
 *  don't fail validation), each issue carries `severity`, and a fixable one
 *  carries the `--fix` command as its `hint`. */
export function reportJson(
	report: Report,
	filePath: string,
): Record<string, unknown> {
	const fixCommand = fixCommandFor(filePath);
	return {
		valid: report.errors === 0,
		errors: report.errors,
		warnings: report.warnings,
		...(report.fixable > 0 ? { hint: `run ${fixCommand}` } : {}),
		parts: report.parts.map((part) => ({
			part: part.name,
			issues: part.issues.map((issue) =>
				issue.fixable ? { ...issue, hint: `run ${fixCommand}` } : issue,
			),
		})),
	};
}

/** The text report: a verdict line first (`valid — N parts checked[, W
 *  warnings]` or `invalid — …`), then each part with findings and its
 *  `severity: message` lines, then — when any finding is fixable — the one
 *  `--fix` command that cures them. */
export function reportText(report: Report, filePath: string): string {
	const checked = `${report.checked} part${plural(report.checked)} checked`;
	const warnings = `${report.warnings} warning${plural(report.warnings)}`;
	const lines: string[] = [];
	if (report.errors === 0) {
		lines.push(
			report.warnings === 0
				? `valid — ${checked}`
				: `valid — ${checked}, ${warnings}`,
		);
	} else {
		const errors = `${report.errors} error${plural(report.errors)}`;
		lines.push(
			`invalid — ${errors}${report.warnings > 0 ? `, ${warnings}` : ""} in ${checked}`,
		);
	}
	for (const part of report.parts) {
		if (part.issues.length === 0) continue;
		lines.push(`${part.name}: ${countBySeverity(part.issues)}`);
		for (const issue of part.issues) {
			const tags = [
				...(issue.kind ? [issue.kind] : []),
				...(issue.fixable ? ["fixable"] : []),
			];
			lines.push(
				`  ${issue.severity}: ${issue.message}${issue.line !== undefined ? ` (line ${issue.line})` : ""}${tags.length > 0 ? ` (${tags.join(", ")})` : ""}`,
			);
			if (issue.hint) lines.push(`    hint: ${issue.hint}`);
		}
	}
	if (report.fixable > 0) {
		lines.push(
			`${report.fixable} finding${plural(report.fixable)} marked (fixable) — run ${fixCommandFor(filePath)} to cure ${report.fixable === 1 ? "it" : "them"}`,
		);
	}
	return `${lines.join("\n")}\n`;
}

function countBySeverity(issues: ReportedIssue[]): string {
	const errors = issues.filter((issue) => issue.severity === "error").length;
	const warnings = issues.length - errors;
	return [
		...(errors > 0 ? [`${errors} error${plural(errors)}`] : []),
		...(warnings > 0 ? [`${warnings} warning${plural(warnings)}`] : []),
	].join(", ");
}

function fixCommandFor(filePath: string): string {
	const quoted = /^[\w./-]+$/.test(filePath) ? filePath : `'${filePath}'`;
	return `\`docx validate ${quoted} --fix\``;
}

const plural = (count: number) => (count === 1 ? "" : "s");

export type Report = {
	strict: boolean;
	/** WML parts the schema pass checked (the verdict line's count). */
	checked: number;
	parts: PartReport[];
	errors: number;
	warnings: number;
	fixable: number;
};

type PartReport = { name: string; issues: ReportedIssue[] };

type ReportedIssue = ValidationIssue & {
	kind?: "consistency";
	severity: "error" | "warning";
	fixable?: true;
	hint?: string;
};
