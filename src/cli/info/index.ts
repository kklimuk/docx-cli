import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	schema: () => import("./schema"),
	locators: () => import("./locators"),
	skill: () => import("./skill"),
};

const HELP = `docx info — print reference material about the CLI

Usage:
  docx info <topic> [options]

Topics:
  schema    Dump AST JSON Schema (or TS source via --ts)
  locators  Dump locator grammar reference
  skill     Print the canonical Agent Skill (SKILL.md)

Run "docx info <topic> --help" for topic-specific help.
`;

export async function run(args: string[]): Promise<number> {
	const topic = args[0];
	if (!topic || topic === "--help" || topic === "-h" || topic === "help") {
		await writeStdout(HELP);
		return topic ? 0 : 2;
	}
	const loader = SUBCOMMANDS[topic];
	if (!loader) {
		return fail(
			"USAGE",
			`Unknown info topic: ${topic}`,
			'Run "docx info --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
