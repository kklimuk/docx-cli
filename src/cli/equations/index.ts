import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	edit: () => import("./edit"),
};

const HELP = `docx equations — author and edit LaTeX equations

Usage:
  docx equations <verb> FILE [options]

Verbs:
  add   Insert an equation (--after/--before/--at-start/--at-end; --equation LATEX; --display)
  edit  Retype or re-mode an existing equation (--at eqN; --equation/--display/--inline)

In \`read\`, find/replace and edit text, an equation is its token: $latex$
inline, $$latex$$ for a display block. So \`docx replace FILE 'mc^2' 'mc^3'\`
rewrites the LaTeX of $E=mc^2$, replacing the whole token deletes it, and a
whole-paragraph \`docx edit\` keeps every token it leaves in place. Bash
expands $ inside double quotes — single-quote LaTeX, or put it in a --batch
file: {"at":"p0","markdown":"Energy is $E=mc^2$ now"}.

Run "docx equations <verb> --help" for verb-specific help.
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
			`Unknown equations subcommand: ${verb}`,
			'Run "docx equations --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
