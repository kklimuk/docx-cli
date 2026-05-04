import { printTopHelp, VERSION } from "./help";
import { fail, writeStdout } from "./respond";

type CommandFn = (args: string[]) => Promise<number>;

const COMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	comments: () => import("./comments"),
	create: () => import("./create"),
	delete: () => import("./delete"),
	edit: () => import("./edit"),
	images: () => import("./images"),
	insert: () => import("./insert"),
	locators: () => import("./locators-cmd"),
	read: () => import("./read"),
	schema: () => import("./schema"),
	"track-changes": () => import("./track-changes"),
};

export async function main(argv: string[]): Promise<number> {
	const args = argv.slice(2);
	const cmd = args[0];

	if (!cmd) {
		process.stderr.write(
			'Usage: docx <command> [options]\nRun "docx --help" for available commands.\n',
		);
		return 2;
	}

	if (cmd === "--help" || cmd === "-h" || cmd === "help") {
		await printTopHelp();
		return 0;
	}

	if (cmd === "--version") {
		const json = args.includes("--json");
		if (json) {
			await writeStdout(`${JSON.stringify({ version: VERSION })}\n`);
		} else {
			await writeStdout(`docx ${VERSION}\n`);
		}
		return 0;
	}

	const loader = COMMANDS[cmd];
	if (!loader) {
		return fail(
			"USAGE",
			`Unknown command: ${cmd}`,
			'Run "docx --help" for the list of commands.',
		);
	}

	const mod = await loader();
	try {
		return await mod.run(args.slice(1));
	} catch (e) {
		const msg = e instanceof Error ? e.message : String(e);
		return fail("UNHANDLED", msg);
	}
}
