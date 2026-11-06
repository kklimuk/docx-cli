import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	extract: () => import("./extract"),
	list: () => import("./list"),
};

const HELP = `docx charts — inspect native Word charts

Usage:
  docx charts <verb> FILE [options]

Verbs:
  list     Print the chart manifest as JSON (id, type, title, series, points)
  extract  Print one chart's data as CSV (or --json) — the cached values Word draws

Charts show in \`docx read\` as \`[chart]\` plus a <!-- docx:chart chartN … --> hint;
chartN is the --at handle.

Run "docx charts <verb> --help" for verb-specific help.
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
			`Unknown charts subcommand: ${verb}`,
			'Run "docx charts --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
