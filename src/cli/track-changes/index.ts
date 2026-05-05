import { fail, writeStdout } from "../respond";

const HELP = `docx track-changes — manage tracked-changes

Usage:
  docx track-changes FILE on|off [options]
  docx track-changes list FILE [options]
  docx track-changes accept FILE (--at tcN | --all) [options]
  docx track-changes reject FILE (--at tcN | --all) [options]

Verbs:
  on        Set <w:trackChanges/> in word/settings.xml
  off       Remove <w:trackChanges/>
  list      List every <w:ins>/<w:del> with id, kind, author, date, location
  accept    Accept tracked changes (incorporate insertions; remove deletions)
  reject    Reject tracked changes (drop insertions; restore deletions as plain)

When tracking is on, this CLI's insert/edit/delete/replace commands
emit <w:ins>/<w:del> markers (attributed via --author or $DOCX_AUTHOR).
Accept/reject themselves bypass tracking — they're doc surgery, not edits.

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
	const module_ = (await import("./toggle")) as { run: CommandFn };
	return module_.run(args);
}
