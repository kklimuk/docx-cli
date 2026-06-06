import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	delete: () => import("./delete"),
	list: () => import("./list"),
	replace: () => import("./replace"),
};

const HELP = `docx hyperlinks — manage hyperlinks

Usage:
  docx hyperlinks <verb> FILE [options]

Verbs:
  add      Wrap a span (pN:S-E / tN:rRcC:pK:S-E) in a hyperlink; mints a linkN
  list     Print the hyperlink manifest as JSON (each id is a linkN handle)
  replace  Change an existing hyperlink's URL (--at linkN)
  delete   Unwrap an existing hyperlink, keeping the text (--at linkN)

Hyperlinks are addressed by id (linkN); discover ids with
"docx hyperlinks list FILE". Run "docx hyperlinks <verb> --help" for
verb-specific help. See "docx info locators" for locator syntax.
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
			`Unknown hyperlinks subcommand: ${verb}`,
			'Run "docx hyperlinks --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
