import { fail, writeStdout } from "../respond";

const HELP = `docx track-changes — manage tracked-changes

Usage:
  docx track-changes on|off FILE [options]
  docx track-changes list FILE [options]
  docx track-changes accept FILE (--at tcN | --all) [options]
  docx track-changes reject FILE (--at tcN | --all) [options]
  docx track-changes apply  FILE (--accept H ... | --reject H ...) [options]

Verbs:
  on / off  Toggle the document's track-changes mode (existing changes keep)
  list      Inventory every tracked change (insertions, deletions, moves,
            property revisions) with stable tcN / revN handles
  accept    Accept tracked changes — like Word: insertions stay, deleted
            text disappears for real
  reject    Reject tracked changes — insertions are removed, deleted text
            comes back
  apply     Finalize: accept AND reject in ONE atomic call, every handle
            resolved against the original tree — the safe way to apply a
            review, since separate accept/reject calls renumber ids between
            them

Exact-change addressing is always --at tcN (repeatable); --all targets every
change; \`apply\` takes --accept/--reject handle lists. Discover ids with
"docx track-changes list FILE"; see them in context with "docx read --current".

When tracking is on, the SUBSEQUENT insert/edit/delete/replace commands record
their edits as tracked changes (attributed via --author or $DOCX_AUTHOR on
those commands, not on the on/off toggle). Accept/reject themselves bypass
tracking — they're review decisions, not edits.

Run "docx track-changes <verb> --help" for verb-specific help.
`;

type CommandFn = (args: string[]) => Promise<number>;

export async function run(args: string[]): Promise<number> {
	const first = args[0];
	if (first === "--help" || first === "-h" || first === "help") {
		await writeStdout(HELP);
		return 0;
	}
	if (!first) {
		return fail("USAGE", "Missing arguments", HELP);
	}
	if (first === "list") {
		const module_ = (await import("./list")) as { run: CommandFn };
		return module_.run(args.slice(1));
	}
	if (first === "accept") {
		const module_ = (await import("./accept")) as { run: CommandFn };
		return module_.run(args.slice(1));
	}
	if (first === "reject") {
		const module_ = (await import("./reject")) as { run: CommandFn };
		return module_.run(args.slice(1));
	}
	if (first === "apply") {
		const module_ = (await import("./apply")) as { run: CommandFn };
		return module_.run(args.slice(1));
	}
	const module_ = (await import("./toggle")) as { run: CommandFn };
	return module_.run(args);
}
