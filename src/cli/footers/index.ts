import { runClearMarginal } from "../headers/clear";
import { runListMarginals } from "../headers/list";
import { runSetMarginal } from "../headers/set";
import { dispatcherHelp } from "../headers/shared";
import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

// Re-uses the header verb implementations with kind="footer" — see
// `cli/headers/index.ts` and `core/marginals/config.ts` (the `MarginalKind`
// parameterization), mirroring how `cli/endnotes` re-uses `cli/footnotes`.
const SUBCOMMANDS: Record<string, CommandFn> = {
	set: (args) => runSetMarginal(args, "footer"),
	list: (args) => runListMarginals(args, "footer"),
	clear: (args) => runClearMarginal(args, "footer"),
};

const HELP = dispatcherHelp("footer");

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
			`Unknown footers subcommand: ${verb}`,
			'Run "docx footers --help".',
		);
	}
	return handler(args.slice(1));
}
