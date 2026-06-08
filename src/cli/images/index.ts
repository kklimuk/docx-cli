import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	delete: () => import("./delete"),
	extract: () => import("./extract"),
	list: () => import("./list"),
	replace: () => import("./replace"),
};

const HELP = `docx images — manage embedded images

Usage:
  docx images <verb> FILE [options]

Verbs:
  add      Insert an image (alias for \`docx insert --image\`)
  list     Print image manifest as JSON
  extract  Dump image bytes to a directory
  replace  Swap an image's bytes
  delete   Remove an embedded image (prunes the part if unreferenced)

Run "docx images <verb> --help" for verb-specific help.
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
			`Unknown images subcommand: ${verb}`,
			'Run "docx images --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
