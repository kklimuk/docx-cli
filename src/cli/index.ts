import { printTopHelp, VERSION } from "./help";
import { fail, writeStderr, writeStdout } from "./respond";

type CommandFn = (args: string[]) => Promise<number>;

const COMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	comments: () => import("./comments"),
	create: () => import("./create"),
	delete: () => import("./delete"),
	edit: () => import("./edit"),
	endnotes: () => import("./endnotes"),
	find: () => import("./find"),
	footers: () => import("./footers"),
	footnotes: () => import("./footnotes"),
	headers: () => import("./headers"),
	hyperlinks: () => import("./hyperlinks"),
	images: () => import("./images"),
	info: () => import("./info"),
	insert: () => import("./insert"),
	lists: () => import("./lists"),
	outline: () => import("./outline"),
	read: () => import("./read"),
	render: () => import("./render"),
	replace: () => import("./replace"),
	sections: () => import("./sections"),
	styles: () => import("./styles"),
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

	// `columns` was renamed to `sections` (it manages section layout, and the
	// read side already speaks sN / docx:section / cols=). Redirect rather than a
	// bare "unknown command" so an agent reaching for the old name lands right.
	if (cmd === "columns") {
		return fail(
			"USAGE",
			"`docx columns` was renamed to `docx sections`",
			"Run `docx sections --at pN-pM --columns N` (or `docx sections --help`).",
		);
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
