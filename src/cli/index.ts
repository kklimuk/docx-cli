import { printTopHelp, VERSION } from "./help";
import { fail, writeStderr, writeStdout } from "./respond";

type CommandFn = (args: string[]) => Promise<number>;

const COMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	comments: () => import("./comments"),
	create: () => import("./create"),
	delete: () => import("./delete"),
	edit: () => import("./edit"),
	find: () => import("./find"),
	hyperlinks: () => import("./hyperlinks"),
	images: () => import("./images"),
	info: () => import("./info"),
	insert: () => import("./insert"),
	outline: () => import("./outline"),
	read: () => import("./read"),
	replace: () => import("./replace"),
	tables: () => import("./tables"),
	"track-changes": () => import("./track-changes"),
	wc: () => import("./wc"),
};

export async function main(argv: string[]): Promise<number> {
	const args = argv.slice(2);
	const cmd = args[0];

	if (!cmd) {
		await writeStderr(
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
