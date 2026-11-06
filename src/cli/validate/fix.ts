import { Comments, Document } from "@core";
import type { Pkg } from "@core/ast/document/package";
import { EXIT, respond, writeStdout } from "../respond";
import { type Report, reportJson, reportText, validatePackage } from "./report";

/** `docx validate FILE --fix`: cure everything the audit marked fixable —
 *  the comment-marker cures plus the sidecar reconcile (`Comments.repair`) —
 *  save through the views (a save also heals the legacy schema defects every
 *  docx-cli save heals), then re-validate the staged result and report it.
 *  `--dry-run` stages and reports without writing; `-o` writes elsewhere.
 *  Exit follows the POST-fix verdict (1 only if errors remain); a clean file
 *  is "nothing to fix", exit 0 — validate is a check, not a mutation that
 *  must change something. */
export async function runFix(
	pkg: Pkg,
	filePath: string,
	options: { output?: string; dryRun: boolean; json: boolean },
	before: Report,
): Promise<number> {
	const document = await Document.fromPackage(pkg);
	const repairs = new Comments(document).repair();
	// Stage (in memory) whenever a re-save might help — schema errors a save
	// heals included — but WRITE only when something actually changed or `-o`
	// asked for a file: an unfixable file must not be rewritten under a
	// "nothing to fix" report.
	const shouldStage =
		repairs.length > 0 || before.errors > 0 || options.output !== undefined;
	if (shouldStage) document.stage();
	const after = shouldStage ? await validatePackage(document.pkg) : before;

	const healedSchemaErrors = Math.max(
		0,
		schemaErrors(before) - schemaErrors(after),
	);
	const shouldWrite =
		repairs.length > 0 ||
		healedSchemaErrors > 0 ||
		options.output !== undefined;
	if (shouldWrite && !options.dryRun) await document.pkg.save(options.output);
	const reportPath = options.output ?? filePath;
	if (options.json) {
		await respond({
			path: filePath,
			...(options.output ? { output: options.output } : {}),
			...(options.dryRun ? { dryRun: true } : {}),
			fixed: repairs.map((repair) => ({ part: repair.part, fix: repair.fix })),
			...(healedSchemaErrors > 0 ? { healedSchemaErrors } : {}),
			...reportJson(after, reportPath),
		});
		return after.errors === 0 ? EXIT.OK : EXIT.GENERAL_ERROR;
	}

	const lines: string[] = [];
	if (repairs.length === 0 && healedSchemaErrors === 0) {
		lines.push("nothing to fix");
	} else if (repairs.length > 0) {
		const verb = options.dryRun ? "would fix" : "fixed";
		const where =
			options.output && !options.dryRun
				? ` (written to ${options.output})`
				: "";
		lines.push(
			`${verb} ${repairs.length} issue${repairs.length === 1 ? "" : "s"}${where}:`,
		);
		for (const repair of repairs) lines.push(`  ${repair.part}: ${repair.fix}`);
	}
	if (healedSchemaErrors > 0) {
		lines.push(
			`${options.dryRun ? "re-saving would also clear" : "re-saving also cleared"} ${healedSchemaErrors} schema error${healedSchemaErrors === 1 ? "" : "s"} (legacy markup every docx-cli save heals)`,
		);
	}
	if (options.dryRun) lines.push(`dry run — ${filePath} not written; result:`);
	await writeStdout(`${lines.join("\n")}\n${reportText(after, reportPath)}`);
	return after.errors === 0 ? EXIT.OK : EXIT.GENERAL_ERROR;
}

function schemaErrors(report: Report): number {
	return report.parts
		.flatMap((part) => part.issues)
		.filter((issue) => issue.kind === undefined).length;
}
