import { openDocView, PkgError } from "@core";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";

const HELP = `docx comments list — print existing comments as JSON

Usage:
  docx comments list FILE [options]

Options:
  --include-resolved   Include resolved comments (default excludes them)
  --thread cN          Print only the thread rooted at the given comment id
  -h, --help           Show this help

Examples:
  docx comments list doc.docx
  docx comments list doc.docx --include-resolved | jq '.[] | select(.author == "Jane")'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				"include-resolved": { type: "boolean" },
				thread: { type: "string" },
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

	let comments = view.doc.comments;
	if (!parsed.values["include-resolved"]) {
		comments = comments.filter((comment) => !comment.resolved);
	}
	const threadFilter = parsed.values.thread as string | undefined;
	if (threadFilter) {
		const allowed = collectThread(comments, threadFilter);
		comments = comments.filter((comment) => allowed.has(comment.id));
	}

	await respond(comments);
	return EXIT.OK;
}

function collectThread(
	comments: Awaited<ReturnType<typeof openDocView>>["doc"]["comments"],
	rootId: string,
): Set<string> {
	const allowed = new Set<string>([rootId]);
	let changed = true;
	while (changed) {
		changed = false;
		for (const comment of comments) {
			if (allowed.has(comment.id)) continue;
			if (comment.parentId && allowed.has(comment.parentId)) {
				allowed.add(comment.id);
				changed = true;
			}
		}
	}
	return allowed;
}
