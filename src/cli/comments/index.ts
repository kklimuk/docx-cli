import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	delete: () => import("./delete"),
	list: () => import("./list"),
	reply: () => import("./reply"),
	resolve: () => import("./resolve"),
	restore: () => import("./restore"),
};

const HELP = `docx comments — manage Word comments

Usage:
  docx comments <verb> FILE [options]

Verbs:
  add      Add a new comment anchored to a locator
  reply    Reply to an existing comment
  list     Print existing comments as JSON
  resolve  Mark a comment resolved
  delete   Remove a comment
  restore  Restore a recently deleted comment

Run "docx comments <verb> --help" for verb-specific help.
`;

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const loader = SUBCOMMANDS[verb];
	if (!loader) {
		return fail(
			"USAGE",
			`Unknown comments subcommand: ${verb}`,
			'Run "docx comments --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
