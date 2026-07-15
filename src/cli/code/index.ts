import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	edit: () => import("./edit"),
};

const HELP = `docx code — author syntax-highlighted code blocks

Usage:
  docx code <verb> FILE [options]

Verbs:
  add   Insert a code block (--after/--before/--at-start/--at-end; --language LANG)
  edit  Replace a paragraph (or range) with a code block (--at pN | pN-pM; --language LANG)

Run "docx code <verb> --help" for verb-specific help.
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
			`Unknown code subcommand: ${verb}`,
			'Run "docx code --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
