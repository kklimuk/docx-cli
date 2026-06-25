import type { Document } from "@core";
import { LocatorParseError, parseLocator } from "@core";
import type { TextMatch } from "@core/find";
import type { ErrorCode } from "../respond";

/** `replace --at LOCATOR` confines a substitution to one paragraph. This is the
 *  cure for the résumé friction: a template repeats the same placeholder in
 *  every entry (`City, State`, `Position Title`, `Month Year`), so a bare
 *  `replace` (first match, document order) can't safely target THE one in the
 *  entry being filled — the agent is forced onto `find` + `edit --at pN:a-b`
 *  span surgery. Scoping a replace to the paragraph the agent already sees in
 *  the read collapses that find→offset→edit→reread loop into ONE offset-free
 *  call. Scope is a single paragraph: a body `pN` or a cell paragraph
 *  `tT:rRcC:pN` — exactly the `blockId` a match carries, so filtering is an
 *  equality test. */
export class ScopeError extends Error {
	constructor(
		public code: ErrorCode,
		message: string,
	) {
		super(message);
		this.name = "ScopeError";
	}
}

/** Validate that `at` is a single-paragraph locator SHAPE (no document needed),
 *  returning the canonical `blockId` string to filter matches by. Throws
 *  {@link ScopeError} on a non-paragraph locator. Used by the batch path, which
 *  validates entries before opening the document. */
export function validateScopeShape(at: string): string {
	let parsed: ReturnType<typeof parseLocator>;
	try {
		parsed = parseLocator(at);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			throw new ScopeError("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	if (!isParagraphLocator(parsed)) {
		throw new ScopeError(
			"INVALID_LOCATOR",
			`--at scope must be a single paragraph (pN) or cell paragraph (tT:rRcC:pN), got "${at}" — replace targets text within one paragraph`,
		);
	}
	return at;
}

/** Is `parsed` a single PARAGRAPH locator? A body `pN`, or a cell paragraph at
 *  any table-nesting depth (`tT:rRcC:pN`, `tT:rRcC:tU:rVcW:pN`, …). Rejects
 *  whole blocks (`tN`/`sN`), whole cells (no inner paragraph), and offset
 *  spans/ranges — `replace` targets text within one paragraph, and these are
 *  exactly the ids `find`/`read`/`edit` already address. */
function isParagraphLocator(parsed: ReturnType<typeof parseLocator>): boolean {
	if (parsed.kind === "block") return parsed.blockId[0] === "p";
	// Descend through nested cells to the innermost locator.
	let inner: ReturnType<typeof parseLocator> | undefined = parsed;
	while (inner?.kind === "cell") inner = inner.inner;
	return inner?.kind === "block" && inner.blockId[0] === "p";
}

/** Validate the scope shape AND that the paragraph exists in `document`,
 *  returning the `blockId` to filter matches by. The single-invocation path
 *  errors loudly on a typo'd scope rather than silently matching nothing.
 *  Throws {@link ScopeError} (mapped to a `fail` by the caller). */
export function resolveReplaceScope(document: Document, at: string): string {
	const blockId = validateScopeShape(at);
	// resolveBlock handles both body (pN) and cell (tT:rRcC:pN) forms.
	try {
		document.body.resolveBlock(at);
	} catch (error) {
		throw new ScopeError("BLOCK_NOT_FOUND", (error as Error).message);
	}
	return blockId;
}

/** Keep only the matches inside the scoped paragraph. A match's `blockId` is the
 *  paragraph's locator (`p20`, or `t1:r2c0:p0` for a cell paragraph), so the
 *  scope is an equality test. */
export function matchesInScope(
	matches: TextMatch[],
	blockId: string,
): TextMatch[] {
	return matches.filter((match) => match.blockId === blockId);
}
