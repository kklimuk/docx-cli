import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	"insert-row": () => import("./insert-row"),
	"delete-row": () => import("./delete-row"),
	"insert-column": () => import("./insert-column"),
	"delete-column": () => import("./delete-column"),
	"set-widths": () => import("./set-widths"),
	merge: () => import("./merge"),
	unmerge: () => import("./unmerge"),
	borders: () => import("./borders"),
};

const HELP = `docx tables — restructure tables (rows, columns, merges, widths, borders)

Usage:
  docx tables <verb> FILE [options]

Verbs:
  insert-row      Insert a row (--at tN [--position INDEX] [--cells "a,b,c"])
  delete-row      Delete a row (--at tN:rR)
  insert-column   Insert a column (--at tN [--position INDEX] [--width TWIPS])
  delete-column   Delete a column (--at tN:cC)
  set-widths      Set column widths (--at tN --widths "20,30,50" | twips | auto)
  merge           Merge a cell region (--at tN:rR1cC1-rR2cC2)
  unmerge         Split a merge anchor (--at tN:rRcC)
  borders         Set table borders (--at tN [--style] [--size] [--color])

These verbs restructure an existing table. The rest of the table lifecycle uses
the standard verbs:
  create        docx insert FILE --after pN --table --rows N --cols M
  delete (all)  docx delete FILE --at tN
  edit a cell   docx edit FILE --at tN:rRcC:pK --text "..."
  inspect       docx read FILE --ast   (grid widths, gridSpan, vMerge per cell)

Run "docx tables <verb> --help" for verb-specific help.
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
			`Unknown tables subcommand: ${verb}`,
			'Run "docx tables --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
