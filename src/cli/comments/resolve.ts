import { openDocView, PkgError, saveDocView } from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";
import {
	ensureCommentParaId,
	ensureCommentsExtPart,
	findCommentByNumericId,
} from "./helpers";

const HELP = `docx comments resolve — mark a comment resolved

Usage:
  docx comments resolve FILE --id cN [options]

Required:
  --id ID           Comment id (e.g., c0)

Optional:
  --unset           Mark unresolved instead of resolved
  --dry-run         Print what would change; do not write the file
  -h, --help        Show this help

Examples:
  docx comments resolve doc.docx --id c2
  docx comments resolve doc.docx --id c2 --unset
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				id: { type: "string" },
				unset: { type: "boolean" },
				"dry-run": { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const idInput = parsed.values.id as string | undefined;
	if (!idInput) return fail("USAGE", "Missing --id COMMENT_ID", HELP);
	const numericId = idInput.startsWith("c") ? idInput.slice(1) : idInput;
	const resolved = !parsed.values.unset;

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (openError) {
		if (openError instanceof PkgError) {
			if (openError.code === "FILE_NOT_FOUND") {
				return fail("FILE_NOT_FOUND", openError.message);
			}
			if (openError.code === "NOT_A_ZIP") {
				return fail("NOT_A_ZIP", openError.message);
			}
		}
		throw openError;
	}

	const commentReference = findCommentByNumericId(view, numericId);
	if (!commentReference) {
		return fail("COMMENT_NOT_FOUND", `Comment not found: c${numericId}`);
	}

	const paraId = ensureCommentParaId(view, idInput);
	if (!paraId) {
		return fail(
			"COMMENT_NOT_FOUND",
			`Comment c${numericId} could not be assigned a w14:paraId.`,
		);
	}

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.resolve",
			dryRun: true,
			path,
			commentId: `c${numericId}`,
			resolved,
		});
		return EXIT.OK;
	}

	const extRoot = ensureCommentsExtPart(view);
	let entry = extRoot.children.find(
		(child) =>
			child.tag === "w15:commentEx" &&
			child.getAttribute("w15:paraId") === paraId,
	);
	if (!entry) {
		entry = new XmlNode("w15:commentEx", { "w15:paraId": paraId });
		extRoot.children.push(entry);
	}
	if (resolved) entry.setAttribute("w15:done", "1");
	else delete entry.attributes["w15:done"];

	await saveDocView(view);

	await respond({
		ok: true,
		operation: "comments.resolve",
		path,
		commentId: `c${numericId}`,
		resolved,
	});
	return EXIT.OK;
}
