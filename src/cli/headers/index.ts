import { fail, writeStdout } from "../respond";
import { runClearMarginal } from "./clear";
import { runListMarginals } from "./list";
import { runSetMarginal } from "./set";
import { dispatcherHelp } from "./shared";

type CommandFn = (args: string[]) => Promise<number>;

// Headers and footers share the same mechanics — the only difference is the part
// name and a handful of tag/style names, all parameterized in
// `core/marginals/config.ts` as `MarginalKind`. This dispatcher binds the shared
// verb implementations to kind="header"; `cli/footers` binds them to "footer".
const SUBCOMMANDS: Record<string, CommandFn> = {
	set: (args) => runSetMarginal(args, "header"),
	list: (args) => runListMarginals(args, "header"),
	clear: (args) => runClearMarginal(args, "header"),
};

const HELP = dispatcherHelp("header");

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const handler = SUBCOMMANDS[verb];
	if (!handler) {
		return fail(
			"USAGE",
			`Unknown headers subcommand: ${verb}`,
			'Run "docx headers --help".',
		);
	}
	return handler(args.slice(1));
}
