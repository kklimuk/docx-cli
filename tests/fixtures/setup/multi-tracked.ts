import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/multi-tracked.docx — a four-paragraph doc with
 * five tracked replacements (eleven tracked-change wrappers total), one
 * per paragraph on p0..p2 plus a cross-author pair on p3. Used by B2's
 * batch accept/reject tests to verify atomic resolution of multiple `--at`
 * ids against the pre-mutation tree.
 *
 *   p0: "Alpha is first."  → "Aleph is first."   (del + ins)
 *   p1: "Beta is second."  → "Bet is second."    (del + ins)
 *   p2: "Gamma is third."  → "Gimel is third."   (del + ins)
 *   p3: "Delta is fourth." → "Delta is fourth & final." — Reviewer A
 *       replaces "fourth" with "fourth and final", then Editor replaces
 *       "and" inside A's still-pending insertion (del A + split ins A
 *       around Editor's nested del + own ins: 5 wrappers, tc6..tc10)
 *
 * After construction: 11 tracked changes; p0..p2 hold tc0..tc5 in order.
 * Each replace lives in its own paragraph so picking non-adjacent ids
 * (e.g. tc0, tc2, tc4) targets one wrapper per paragraph — exactly the
 * scenario where one-at-a-time accepting would shift later ids.
 *
 * Author/date are pinned via env vars so the fixture is reproducible.
 *
 * Built by dogfooding the CLI; doubles as a smoke test for `create`,
 * `insert`, `track-changes on`, and `replace` under tracking.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/multi-tracked.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const env = {
		DOCX_AUTHOR: "Reviewer",
		DOCX_CLI_NOW: "2026-05-07T20:00:00Z",
	};
	const result = await $.env(env)`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

mkdirSync(dirname(out), { recursive: true });

await cli(
	"create",
	out,
	"--title",
	"Multi-tracked fixture",
	"--author",
	"docx-cli",
	"--text",
	"Alpha is first.",
	"--force",
);

await cli("insert", out, "--after", "p0", "--text", "Beta is second.");
await cli("insert", out, "--after", "p1", "--text", "Gamma is third.");
await cli("insert", out, "--after", "p2", "--text", "Delta is fourth.");
await cli("track-changes", out, "on");
await cli("replace", out, "Alpha", "Aleph");
await cli("replace", out, "Beta", "Bet");
await cli("replace", out, "Gamma", "Gimel");
// Cross-author span replace (issue #13): Editor replaces text INSIDE Reviewer
// A's pending insertion, so A's <w:ins> splits around Editor's own <w:ins>
// (the trailing half with a fresh w:id) and Editor's <w:del> nests in A's.
await cli(
	"replace",
	out,
	"fourth",
	"fourth and final",
	"--author",
	"Reviewer A",
);
await cli("replace", out, "and", "&", "--author", "Editor");

const trackedJson = await cli("track-changes", "list", out, "--json");
const changes = JSON.parse(trackedJson) as Array<{
	id: string;
	kind: string;
	blockId: string;
	author: string;
}>;
console.log(`Wrote ${out}`);
console.log(`Tracked changes: ${changes.length}`);
for (const change of changes) {
	console.log(
		`  ${change.id}: ${change.kind} on ${change.blockId} (by ${change.author})`,
	);
}
