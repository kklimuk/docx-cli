import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture, readDocumentXml } from "./helpers";

let workspace: string;
let docPath: string;

beforeEach(async () => {
	workspace = tempWorkspace("validate");
	docPath = join(workspace, "doc.docx");
	await runCli("create", docPath, "--text", "clean document");
});

describe("docx validate", () => {
	test("a clean document validates with exit 0", async () => {
		const result = await runCli("validate", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("valid");
	});

	test("a Word-authored document (MCE-laden) validates clean", async () => {
		const wordDoc = await freshFixture(
			"validate-word",
			"tests/fixtures/academic-paper.docx",
		);
		const result = await runCli("validate", wordDoc);
		expect(result.exitCode).toBe(0);
	});

	test("schema errors list per part and exit 1", async () => {
		await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--no-validate",
			"--xml",
			"<w:p><w:r><w:bogusChild/><w:t>x</w:t></w:r></w:p>",
		);
		const result = await runCli("validate", docPath);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain("word/document.xml: 1 error");
		expect(result.stdout).toContain("bogusChild");
	});

	test("--json reports valid/errors/parts", async () => {
		const result = await runCli("validate", docPath, "--json");
		const parsed = result.parsed as {
			valid: boolean;
			errors: number;
			parts: { part: string; issues: unknown[] }[];
		};
		expect(parsed.valid).toBe(true);
		expect(parsed.errors).toBe(0);
		expect(parsed.parts.some((part) => part.part === "word/document.xml")).toBe(
			true,
		);
	});

	test("an ISO-strict document is skipped with a note, not flooded with noise", async () => {
		const strictDoc = await freshFixture(
			"validate-strict",
			"tests/fixtures/strict-profile.docx",
		);
		const result = await runCli("validate", strictDoc);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("strict-profile");
	});
});

// The emitter-conformance pin: every fixture the repo ships must be
// schema-valid OOXML. This is what keeps the bug classes the validator audit
// found (settings trackRevisions ordering, tblGridChange attributes, comments
// mc:Ignorable/textId, nested w:rPr in math runs, stale hand-built XML) from
// regressing — Word TOLERATES most of them, so renders and round-trips can't.
// strict-profile.docx is the deliberate exception (ISO-strict namespaces; the
// skip test above covers it).
describe("every fixture validates clean", () => {
	const glob = new Bun.Glob("*.docx");
	const fixtures = [...glob.scanSync("tests/fixtures")]
		.filter((name) => name !== "strict-profile.docx")
		.sort();

	// Clean means no warnings either: the consistency audit's false-positive pin.
	test.each(fixtures)("%s", async (name) => {
		const result = await runCli("validate", `tests/fixtures/${name}`);
		expect(result.stdout.trim()).toMatch(/^valid — \d+ parts? checked$/);
		expect(result.exitCode).toBe(0);
	});
});

describe("cross-part comment consistency — severity per Word's behavior", () => {
	// Schema-valid packages Word still repairs (or refuses) on open. Graded by
	// live Word for Mac 16.114: only a commentReference naming no comment stops
	// Word opening the file (error, exit 1); everything else it repairs or drops
	// on open (warning, exit 0).
	let commented: string;

	beforeEach(async () => {
		commented = join(workspace, "commented.docx");
		await runCli("create", commented, "--text", "hello world");
		await runCli("comments", "add", commented, "--at", "p0", "--text", "n");
	});

	async function patchPart(
		name: string,
		patch: (xml: string) => string,
	): Promise<void> {
		const pkg = await Pkg.open(commented);
		pkg.writeText(name, patch(await pkg.readText(name)));
		await pkg.save();
	}

	async function addPart(name: string, xml: string): Promise<void> {
		const pkg = await Pkg.open(commented);
		pkg.writeText(name, xml);
		await pkg.save();
	}

	async function paraIdOf(commentNumber: number): Promise<string> {
		const comments = await (await Pkg.open(commented)).readText(
			"word/comments.xml",
		);
		const ids = [...comments.matchAll(/w14:paraId="([0-9A-F]{8})"/g)];
		return ids[commentNumber]?.[1] ?? "";
	}

	async function expectClean(path = commented): Promise<void> {
		const result = await runCli("validate", path);
		expect(result.stdout.trim()).toMatch(/^valid — \d+ parts? checked$/);
		expect(result.exitCode).toBe(0);
	}

	const CID =
		'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"';

	test("a consistent commented document validates clean", async () => {
		await expectClean();
	});

	test("a comment missing from a present commentsIds.xml is a fixable WARNING (exit 0)", async () => {
		await addPart("word/commentsIds.xml", `<w16cid:commentsIds ${CID}/>`);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^valid — \d+ parts checked, 1 warning\n/);
		expect(result.stdout).toContain("word/commentsIds.xml: 1 warning");
		expect(result.stdout).toMatch(
			/warning: comment c0 \(paraId [0-9A-F]{8}\) has no entry in word\/commentsIds\.xml — Word mints one on open \(consistency, fixable\)/,
		);
		expect(result.stdout).toContain(`run \`docx validate ${commented} --fix\``);

		const fixed = await runCli("validate", commented, "--fix");
		expect(fixed.exitCode).toBe(0);
		expect(fixed.stdout).toContain("fixed 1 issue:");
		const paraId = await paraIdOf(0);
		expect(fixed.stdout).toContain(
			`word/commentsIds.xml: added c0's entry (durableId ${paraId})`,
		);
		await expectClean();
	});

	test("a comment missing from commentsExtended.xml warns that a reply/resolved state is lost", async () => {
		await addPart(
			"word/commentsExtended.xml",
			'<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("loses its thread");
		expect(result.stdout).toContain("reopens");
	});

	test("a dangling commentReference is an ERROR (Word refuses to open); --fix removes it", async () => {
		await patchPart("word/comments.xml", (xml) =>
			xml.replace(/<w:comment [\s\S]*<\/w:comment>/, ""),
		);
		const result = await runCli("validate", commented, "--json");
		expect(result.exitCode).toBe(1);
		const parsed = result.parsed as {
			valid: boolean;
			errors: number;
			warnings: number;
			hint?: string;
			parts: { part: string; issues: Record<string, unknown>[] }[];
		};
		expect(parsed.valid).toBe(false);
		expect(parsed.errors).toBe(1);
		expect(parsed.hint).toContain("--fix");
		const documentPart = parsed.parts.find(
			(part) => part.part === "word/document.xml",
		);
		expect(documentPart?.issues).toContainEqual({
			message:
				"commentReference 0 points at no comment in word/comments.xml — Word refuses to open the file",
			kind: "consistency",
			severity: "error",
			fixable: true,
			hint: `run \`docx validate ${commented} --fix\``,
		});

		const fixed = await runCli("validate", commented, "--fix");
		expect(fixed.exitCode).toBe(0);
		expect(fixed.stdout).toContain(
			"word/document.xml: removed every marker of the missing comment 0",
		);
		await expectClean();
		expect(await readDocumentXml(commented)).not.toContain("commentRangeStart");
		expect(await readDocumentXml(commented)).toContain("hello world");
	});

	test("a range whose reference run was deleted is a fixable warning; --fix restores the reference", async () => {
		await patchPart("word/document.xml", (xml) =>
			xml.replace(
				/<w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="0"\/><\/w:r>/,
				"",
			),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(
			"warning: comment range 0 has no matching commentReference",
		);
		await runCli("validate", commented, "--fix");
		await expectClean();
		const comments = (await runCli("comments", "list", commented)).parsed as {
			id: string;
			anchor: { startBlockId: string };
		}[];
		expect(comments[0]?.anchor.startBlockId).toBe("p0");
	});

	test("an unpaired commentRangeStart collapses onto the reference mark, as Word does", async () => {
		await patchPart("word/document.xml", (xml) =>
			xml.replace('<w:commentRangeEnd w:id="0"/>', ""),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(
			"commentRangeStart 0 has no matching commentRangeEnd — Word collapses comment c0 to its reference mark",
		);
		await runCli("validate", commented, "--fix");
		await expectClean();
		expect(await readDocumentXml(commented)).toContain(
			'<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/>',
		);
	});

	test("an orphan comment is reported (Word deletes it) but --fix never deletes it", async () => {
		await patchPart("word/document.xml", (xml) =>
			xml
				.replace(/<w:commentRange(Start|End) w:id="0"\/>/g, "")
				.replace(
					/<w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="0"\/><\/w:r>/,
					"",
				),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(
			"comment c0 has no anchor in the document (no commentRangeStart, commentRangeEnd or commentReference) — Word silently deletes it on open (consistency)",
		);
		expect(result.stdout).toContain("hint: re-anchor its text");
		expect(result.stdout).not.toContain("fixable");
		const fixed = await runCli("validate", commented, "--fix");
		expect(fixed.stdout).toMatch(/^nothing to fix\n/);
		expect(fixed.stdout).toContain("silently deletes it");
		const comments = (await runCli("comments", "list", commented))
			.parsed as unknown[];
		expect(comments).toHaveLength(1);
	});

	test("an orphan reply (docx-cli < 0.14.1 wrote replies without markers) is re-anchored in its thread", async () => {
		await runCli("comments", "reply", commented, "--at", "c0", "--text", "r");
		await patchPart("word/document.xml", (xml) =>
			xml
				.replace(/<w:commentRange(Start|End) w:id="1"\/>/g, "")
				.replace(
					/<w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="1"\/><\/w:r>/,
					"",
				),
		);
		const result = await runCli("validate", commented);
		expect(result.stdout).toContain("(it is a reply in c0's thread)");
		expect(result.stdout).toContain("(consistency, fixable)");
		const fixed = await runCli("validate", commented, "--fix");
		expect(fixed.stdout).toContain(
			"anchored reply c1 in c0's thread (mirrored the thread's markers)",
		);
		await expectClean();
		const comments = (await runCli("comments", "list", commented)).parsed as {
			id: string;
			parentId?: string;
		}[];
		expect(comments.find((comment) => comment.id === "c1")?.parentId).toBe(
			"c0",
		);
	});

	test("a dangling paraIdParent is a report-only warning; --fix keeps the link untouched", async () => {
		await runCli("comments", "reply", commented, "--at", "c0", "--text", "r");
		await patchPart("word/commentsExtended.xml", (xml) =>
			xml.replace(
				/w15:paraIdParent="[0-9A-F]{8}"/,
				'w15:paraIdParent="0BADF00D"',
			),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain(
			"comment c1's thread parent (w15:paraIdParent 0BADF00D) names no comment — Word drops the link",
		);
		await runCli("validate", commented, "--fix");
		const extended = await (await Pkg.open(commented)).readText(
			"word/commentsExtended.xml",
		);
		expect(extended).toContain('w15:paraIdParent="0BADF00D"');
	});

	test("a stale commentsIds entry and a duplicate durableId are NOT reported, but --fix and any save clean them", async () => {
		const paraId = await paraIdOf(0);
		const ids = `<w16cid:commentsIds ${CID}><w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="0BADF00D" w16cid:durableId="11111111"/></w16cid:commentsIds>`;
		await addPart("word/commentsIds.xml", ids);
		await expectClean();
		const dryRun = await runCli("validate", commented, "--fix", "--dry-run");
		expect(dryRun.stdout).toContain("would fix 1 issue:");
		expect(dryRun.stdout).toContain(
			"word/commentsIds.xml: pruned the entry for paraId 0BADF00D (no such comment)",
		);
		expect(dryRun.stdout).toContain(`dry run — ${commented} not written`);
		expect(
			await (await Pkg.open(commented)).readText("word/commentsIds.xml"),
		).toBe(ids);

		await runCli("edit", commented, "--at", "p0", "--text", "hello there");
		const healed = await (await Pkg.open(commented)).readText(
			"word/commentsIds.xml",
		);
		expect(healed).not.toContain("0BADF00D");
		expect(healed).toContain(
			`w16cid:paraId="${paraId}" w16cid:durableId="11111111"`,
		);
	});

	test("--fix on a clean file is 'nothing to fix', exit 0, and leaves the bytes alone", async () => {
		const before = await Bun.file(commented).bytes();
		const result = await runCli("validate", commented, "--fix");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(
			/^nothing to fix\nvalid — \d+ parts checked\n$/,
		);
		expect(await Bun.file(commented).bytes()).toEqual(before);
	});

	test("--fix -o writes the cure elsewhere and leaves FILE as it was", async () => {
		await addPart("word/commentsIds.xml", `<w16cid:commentsIds ${CID}/>`);
		const output = join(workspace, "fixed.docx");
		const result = await runCli("validate", commented, "--fix", "-o", output);
		expect(result.stdout).toContain(`fixed 1 issue (written to ${output}):`);
		await expectClean(output);
		expect((await runCli("validate", commented)).stdout).toContain("1 warning");
	});

	test("--fix --json lists the fixes beside the post-fix report", async () => {
		await addPart("word/commentsIds.xml", `<w16cid:commentsIds ${CID}/>`);
		const result = await runCli("validate", commented, "--fix", "--json");
		const parsed = result.parsed as {
			fixed: { part: string; fix: string }[];
			valid: boolean;
			warnings: number;
		};
		expect(parsed.fixed).toHaveLength(1);
		expect(parsed.fixed[0]?.part).toBe("word/commentsIds.xml");
		expect(parsed.valid).toBe(true);
		expect(parsed.warnings).toBe(0);
	});

	test("-o and --dry-run without --fix are usage errors", async () => {
		const result = await runCli("validate", commented, "--dry-run");
		expect(result.exitCode).toBe(2);
	});
});

describe("reconcile on save — any mutation cures a file older docx-cli left inconsistent", () => {
	async function threadState(path: string): Promise<string[]> {
		const comments = (await runCli("comments", "list", path)).parsed as {
			id: string;
			parentId?: string;
			resolved?: boolean;
		}[];
		return comments.map(
			(comment) =>
				`${comment.id}<${comment.parentId ?? "-"}>${comment.resolved ? "resolved" : "open"}`,
		);
	}

	test("an unrelated edit adds missing sidecar entries and keeps existing thread/resolved state", async () => {
		const wordDoc = await freshFixture(
			"reconcile-paper",
			"tests/fixtures/academic-paper.docx",
		);
		await runCli("comments", "resolve", wordDoc, "--at", "c1");
		await runCli("comments", "reply", wordDoc, "--at", "c2", "--text", "r");
		const before = await threadState(wordDoc);
		expect(before).toContain("c1<->resolved");
		expect(before).toContain("c44<c2>open");
		// Simulate an older producer: no comment (the reply c44 and resolved c1 included) has an ids entry.
		const pkg = await Pkg.open(wordDoc);
		const ids = await pkg.readText("word/commentsIds.xml");
		const stripped = ids.replace(/<w16cid:commentId [^>]*\/>/g, "");
		pkg.writeText("word/commentsIds.xml", stripped);
		await pkg.save();
		expect((await runCli("validate", wordDoc)).stdout).toContain("45 warnings");

		await runCli("edit", wordDoc, "--at", "p20", "--text", "Unrelated.");
		const after = await runCli("validate", wordDoc);
		expect(after.stdout.trim()).toMatch(/^valid — \d+ parts checked$/);
		// Thread links and resolved flags survive the cure (anchors on the
		// edited paragraph legitimately move, so compare the thread state only).
		expect(await threadState(wordDoc)).toEqual(before);
	});

	test("an out-of-range comment paraId (docx-cli < 0.14.1 minted them) is re-keyed with its thread intact", async () => {
		// Word re-keys such an id on open WITHOUT carrying its sidecar entries,
		// so every reply in the file loses its thread (verified on Word for Mac).
		const doc = join(workspace, "high-para-ids.docx");
		await runCli("create", doc, "--text", "hello world");
		await runCli("comments", "add", doc, "--at", "p0", "--text", "Root");
		await runCli("comments", "reply", doc, "--at", "c0", "--text", "Reply");
		await runCli("comments", "resolve", doc, "--at", "c0");
		const before = await threadState(doc);
		expect(before).toEqual(["c0<->resolved", "c1<c0>open"]);
		const pkg = await Pkg.open(doc);
		const toHigh = (xml: string) =>
			xml.replace(
				/(paraId(?:Parent)?=")([0-9A-F])([0-9A-F]{7})"/g,
				(_, prefix, first, rest) =>
					`${prefix}${(Number.parseInt(first, 16) | 8).toString(16).toUpperCase()}${rest}"`,
			);
		for (const name of ["word/comments.xml", "word/commentsExtended.xml"]) {
			pkg.writeText(name, toHigh(await pkg.readText(name)));
		}
		await pkg.save();
		const report = await runCli("validate", doc);
		expect(report.exitCode).toBe(0);
		expect(report.stdout).toContain("2 warnings");
		expect(report.stdout).toContain("is out of range");
		expect(report.stdout).toContain("(consistency, fixable)");

		await runCli("edit", doc, "--at", "p0", "--text", "hello there");
		const after = await runCli("validate", doc);
		expect(after.stdout.trim()).toMatch(/^valid — \d+ parts checked$/);
		expect(await threadState(doc)).toEqual(before);
		const comments = await (await Pkg.open(doc)).readText("word/comments.xml");
		for (const [, paraId] of comments.matchAll(/w14:paraId="([0-9A-F]{8})"/g)) {
			expect(Number.parseInt(paraId ?? "", 16)).toBeLessThan(0x80000000);
		}
	});

	test("an unchanged sidecar keeps its original bytes on an unrelated save", async () => {
		const wordDoc = await freshFixture(
			"reconcile-bytes",
			"tests/fixtures/academic-paper.docx",
		);
		const read = async (name: string) =>
			(await Pkg.open(wordDoc)).readText(name);
		const before = await Promise.all(
			["word/commentsIds.xml", "word/commentsExtended.xml"].map(read),
		);
		await runCli("edit", wordDoc, "--at", "p20", "--text", "Unrelated.");
		const after = await Promise.all(
			["word/commentsIds.xml", "word/commentsExtended.xml"].map(read),
		);
		expect(after).toEqual(before);
	});
});

describe("legacy schema noise self-heals on any save", () => {
	test("a misnamed <w:trackChanges> and an undeclared w14 comments part are repaired by an unrelated edit", async () => {
		// Documents produced by earlier docx-cli versions (and some Word files)
		// carry two defects our views can fix: the schema-invalid legacy
		// `<w:trackChanges/>` toggle, and comment paragraphs with `w14:paraId`
		// while the root omits `mc:Ignorable`. An agent that runs `validate` after
		// its own edit read those as "I broke the file." Every save now heals both,
		// so ANY mutation leaves the document clean.
		const legacy = join(workspace, "legacy.docx");
		await runCli("create", legacy, "--text", "Clause one. Clause two.");
		await runCli("track-changes", "on", legacy);
		await runCli("comments", "add", legacy, "--at", "p0", "--text", "note");
		const pkg = await Pkg.open(legacy);
		pkg.writeText(
			"word/settings.xml",
			(await pkg.readText("word/settings.xml"))
				.replace(/<w:trackRevisions\/>/, "")
				.replace(/(<w:settings[^>]*>)/, "$1<w:trackChanges/>"),
		);
		pkg.writeText(
			"word/comments.xml",
			(await pkg.readText("word/comments.xml")).replace(
				/ mc:Ignorable="[^"]*"/,
				"",
			),
		);
		await pkg.save();
		const before = await runCli("validate", legacy);
		expect(before.exitCode).toBe(1);
		expect(before.stdout).toContain("word/settings.xml");
		expect(before.stdout).toContain("word/comments.xml");
		// Tracking still reads as ON through the legacy element.
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);

		// An unrelated mutation — no toggle, no comment — heals both parts.
		await runCli(
			"edit",
			legacy,
			"--at",
			"p0",
			"--text",
			"Clause one, amended.",
		);
		const after = await runCli("validate", legacy);
		expect(after.exitCode).toBe(0);
		expect(after.stdout).toMatch(/^valid/);
		const healed = await Pkg.open(legacy);
		const settings = await healed.readText("word/settings.xml");
		expect(settings).toContain("<w:trackRevisions/>");
		expect(settings).not.toContain("w:trackChanges");
		expect(await healed.readText("word/comments.xml")).toMatch(
			/<w:comments[^>]*mc:Ignorable="[^"]*w14/,
		);
		// The tracked edit landed as a real revision (tracking stayed on).
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);
		const revisions = (await runCli("track-changes", "list", legacy, "--json"))
			.parsed as unknown[];
		expect(revisions.length).toBeGreaterThan(0);
	});

	test("a legacy ON toggle beside a present-but-off <w:trackRevisions> stays ON after an unrelated save", async () => {
		// An earlier docx-cli's `track-changes on` added the misnamed element next
		// to a producer's `w:val="false"` real one. `read` reports ON (any ON
		// toggle wins), so the save-time migration must keep it ON — not let the
		// stale OFF element win and silently flip tracking off.
		const legacy = join(workspace, "conflict.docx");
		await runCli("create", legacy, "--text", "Clause one.");
		const pkg = await Pkg.open(legacy);
		pkg.writeText(
			"word/settings.xml",
			(await pkg.readText("word/settings.xml"))
				.replace(/(<w:settings[^>]*>)/, "$1<w:trackChanges/>")
				// The real element at its CT_Settings slot (before defaultTabStop).
				.replace(/(<w:defaultTabStop)/, '<w:trackRevisions w:val="false"/>$1'),
		);
		await pkg.save();
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);

		await runCli(
			"edit",
			legacy,
			"--at",
			"p0",
			"--text",
			"Clause one, amended.",
		);
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);
		const settings = await (await Pkg.open(legacy)).readText(
			"word/settings.xml",
		);
		expect(settings).toContain("<w:trackRevisions/>");
		expect(settings).not.toContain("w:trackChanges");
		expect((await runCli("validate", legacy)).exitCode).toBe(0);
	});
});

describe("justified alignment emits ST_Jc `both`", () => {
	test("every paragraph emitter writes a schema-valid justify that reads back", async () => {
		// ST_Jc has no `justify` — Word's spelling is `both`. The image and
		// equation paragraphs build their own <w:pPr>, so they must map it too.
		await runCli("edit", docPath, "--at", "p0", "--alignment", "justify");
		await runCli(
			"images",
			"add",
			docPath,
			"--after",
			"p0",
			"--image",
			join(import.meta.dir, "..", "fixtures", "assets", "sample.png"),
			"--alignment",
			"justify",
		);
		await runCli(
			"equations",
			"add",
			docPath,
			"--at-end",
			"--equation",
			"x^2",
			"--alignment",
			"justify",
		);
		const result = await runCli("validate", docPath);
		expect(result.stdout).toMatch(/^valid/);
		expect(result.exitCode).toBe(0);
		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown.match(/align="justify"/g)).toHaveLength(3);
	});
});
