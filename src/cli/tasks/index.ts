import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	check: () => import("./check"),
	uncheck: () => import("./uncheck"),
};

const HELP = `docx tasks — author and toggle GFM task-list checkboxes

Usage:
  docx tasks <verb> FILE [options]

Verbs:
  add      Insert a task-list item (--after/--before/--at-start/--at-end; --checked)
  check    Mark an existing task's checkbox done — ☒ (--at pN)
  uncheck  Clear an existing task's checkbox — ☐ (--at pN)

Run "docx tasks <verb> --help" for verb-specific help.
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
			`Unknown tasks subcommand: ${verb}`,
			'Run "docx tasks --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
