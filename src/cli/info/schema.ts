// @ts-expect-error — bun bundles this as a string via the `text` import attribute
import TS_SOURCE from "@core/ast/types" with { type: "text" };
import { EXIT, respond, tryParseArgs, writeStdout } from "../respond";

const HELP = `docx info schema — print the AST type definitions

Usage:
  docx info schema [options]

Options:
  --json       Print as a JSON Schema document (this is the default — emitted
               when neither --ts nor --json is passed)
  --ts         Print TypeScript type source (from src/core/ast/types.ts, embedded at build time)
  -h, --help   Show this help

Examples:
  docx info schema | jq '.$defs.Run'
  docx info schema --ts > ast.d.ts
`;

const JSON_SCHEMA = {
	$schema: "https://json-schema.org/draft/2020-12/schema",
	$id: "https://github.com/kklimuk/docx-cli/schema",
	title: "docx-cli AST",
	type: "object",
	required: [
		"schemaVersion",
		"path",
		"properties",
		"blocks",
		"comments",
		"footnotes",
		"endnotes",
	],
	properties: {
		schemaVersion: { const: 1 },
		path: { type: "string" },
		properties: {
			type: "object",
			properties: {
				title: { type: "string" },
				author: { type: "string" },
				created: { type: "string" },
				modified: { type: "string" },
			},
		},
		blocks: { type: "array", items: { $ref: "#/$defs/Block" } },
		comments: { type: "array", items: { $ref: "#/$defs/Comment" } },
		footnotes: { type: "array", items: { $ref: "#/$defs/Footnote" } },
		endnotes: { type: "array", items: { $ref: "#/$defs/Footnote" } },
	},
	$defs: {
		Block: {
			oneOf: [
				{ $ref: "#/$defs/Paragraph" },
				{ $ref: "#/$defs/Table" },
				{ $ref: "#/$defs/SectionBreak" },
			],
		},
		Paragraph: {
			type: "object",
			required: ["id", "type", "runs"],
			properties: {
				id: { type: "string" },
				type: { const: "paragraph" },
				style: { type: "string" },
				alignment: { enum: ["left", "center", "right", "justify"] },
				list: {
					type: "object",
					required: ["level", "numId"],
					properties: {
						level: { type: "number" },
						numId: { type: "string" },
					},
				},
				taskState: { enum: ["checked", "unchecked"] },
				tabStops: {
					type: "array",
					items: {
						type: "object",
						required: ["align", "pos"],
						properties: {
							align: { type: "string" },
							pos: { type: "number" },
						},
					},
				},
				spacing: {
					type: "object",
					properties: {
						before: { type: "number" },
						after: { type: "number" },
						line: { type: "number" },
						lineRule: { enum: ["auto", "exact", "atLeast"] },
					},
				},
				indent: {
					type: "object",
					properties: {
						left: { type: "number" },
						right: { type: "number" },
						firstLine: { type: "number" },
						hanging: { type: "number" },
					},
				},
				runs: { type: "array", items: { $ref: "#/$defs/Run" } },
			},
		},
		Run: {
			oneOf: [
				{ $ref: "#/$defs/TextRun" },
				{ $ref: "#/$defs/ImageRun" },
				{ $ref: "#/$defs/BreakRun" },
				{ $ref: "#/$defs/TabRun" },
				{ $ref: "#/$defs/EquationRun" },
				{ $ref: "#/$defs/FootnoteRefRun" },
				{ $ref: "#/$defs/ChartRun" },
			],
		},
		TextRun: {
			type: "object",
			required: ["type", "text"],
			properties: {
				type: { const: "text" },
				text: { type: "string" },
				color: { type: "string" },
				colorTheme: { type: "string" },
				colorThemeTint: { type: "string" },
				colorThemeShade: { type: "string" },
				highlight: { type: "string" },
				shade: { type: "string" },
				bold: { type: "boolean" },
				italic: { type: "boolean" },
				underline: { type: "string" },
				underlineColor: { type: "string" },
				strike: { type: "boolean" },
				vertAlign: { type: "string" },
				smallCaps: { type: "boolean" },
				allCaps: { type: "boolean" },
				font: { type: "string" },
				sizeHalfPoints: { type: "number" },
				runStyle: { type: "string" },
				comments: { type: "array", items: { type: "string" } },
				trackedChange: {
					type: "object",
					required: ["id", "kind", "author", "date", "revisionId"],
					properties: {
						id: { type: "string" },
						// Only kinds that attach to a TextRun are listed here. Other
						// `TrackedChangeKind` values (sectPrChange, rowIns/rowDel,
						// cellIns/cellDel, tbl*Change, tcPrChange, checkboxToggle)
						// surface via `track-changes list` from
						// `document.trackedChangeReferences`, not via `Run.trackedChange`
						// — keep this enum tight to what readers actually see here.
						kind: {
							enum: ["ins", "del", "moveFrom", "moveTo"],
						},
						author: { type: "string" },
						date: { type: "string" },
						revisionId: { type: "string" },
					},
				},
				hyperlink: { $ref: "#/$defs/Hyperlink" },
			},
		},
		Hyperlink: {
			type: "object",
			required: ["id"],
			properties: {
				id: { type: "string" },
				url: { type: "string" },
				anchor: { type: "string" },
				tooltip: { type: "string" },
			},
		},
		ImageRun: {
			type: "object",
			required: ["type", "id", "contentType", "hash"],
			properties: {
				type: { const: "image" },
				id: { type: "string" },
				contentType: { type: "string" },
				hash: { type: "string" },
				widthEmu: { type: "number" },
				heightEmu: { type: "number" },
				alt: { type: "string" },
				floating: { type: "boolean" },
				wrap: { type: "string" },
				align: { type: "string" },
				trackedChange: {
					type: "object",
					required: ["id", "kind", "author", "date", "revisionId"],
					properties: {
						id: { type: "string" },
						kind: { enum: ["ins", "del", "moveFrom", "moveTo"] },
						author: { type: "string" },
						date: { type: "string" },
						revisionId: { type: "string" },
					},
				},
			},
		},
		BreakRun: {
			type: "object",
			required: ["type", "kind"],
			properties: {
				type: { const: "break" },
				kind: { enum: ["page", "line", "column"] },
			},
		},
		TabRun: {
			type: "object",
			required: ["type"],
			properties: { type: { const: "tab" } },
		},
		EquationRun: {
			type: "object",
			required: ["type", "id", "latex", "text", "display"],
			properties: {
				type: { const: "equation" },
				id: { type: "string" },
				latex: { type: "string" },
				text: { type: "string" },
				display: { type: "boolean" },
			},
		},
		FootnoteRefRun: {
			type: "object",
			required: ["type", "kind", "id"],
			properties: {
				type: { const: "noteRef" },
				kind: { enum: ["footnote", "endnote"] },
				id: { type: "string" },
			},
		},
		ChartRun: {
			type: "object",
			required: ["type", "kind"],
			properties: {
				type: { const: "chart" },
				kind: { enum: ["chart", "shape", "smartart", "drawing"] },
			},
		},
		Footnote: {
			type: "object",
			required: ["id", "text"],
			properties: {
				id: { type: "string" },
				text: { type: "string" },
			},
		},
		Table: {
			type: "object",
			required: ["id", "type", "grid", "rows"],
			properties: {
				id: { type: "string" },
				type: { const: "table" },
				grid: { type: "array", items: { type: "number" } },
				width: { $ref: "#/$defs/TableWidth" },
				borders: { type: "string" },
				style: { type: "string" },
				rows: {
					type: "array",
					items: {
						type: "object",
						properties: {
							cells: {
								type: "array",
								items: {
									type: "object",
									properties: {
										blocks: {
											type: "array",
											items: { $ref: "#/$defs/Block" },
										},
										gridSpan: { type: "number" },
										vMerge: { enum: ["restart", "continue"] },
										width: { $ref: "#/$defs/TableWidth" },
										shading: { type: "string" },
										trackedChange: {
											$ref: "#/$defs/TableRevision",
											description: "cellIns / cellDel (tracked column change)",
										},
									},
								},
							},
							trackedChange: {
								$ref: "#/$defs/TableRevision",
								description: "rowIns / rowDel (tracked row change)",
							},
						},
					},
				},
			},
		},
		TableRevision: {
			type: "object",
			required: ["id", "kind", "author", "date", "revisionId"],
			properties: {
				id: { type: "string" },
				kind: {
					enum: ["rowIns", "rowDel", "cellIns", "cellDel", "tcPrChange"],
				},
				author: { type: "string" },
				date: { type: "string" },
				revisionId: { type: "string" },
			},
		},
		TableWidth: {
			type: "object",
			required: ["value", "unit"],
			properties: {
				value: { type: "number" },
				unit: { enum: ["dxa", "pct", "auto", "nil"] },
			},
		},
		SectionBreak: {
			type: "object",
			required: ["id", "type"],
			properties: {
				id: { type: "string" },
				type: { const: "sectionBreak" },
				columns: { type: "number" },
				sectionType: {
					enum: ["continuous", "nextPage", "evenPage", "oddPage", "nextColumn"],
				},
				pageWidth: { type: "number" },
				pageHeight: { type: "number" },
				pageOrientation: { enum: ["portrait", "landscape"] },
				marginTop: { type: "number" },
				marginRight: { type: "number" },
				marginBottom: { type: "number" },
				marginLeft: { type: "number" },
			},
		},
		Comment: {
			type: "object",
			required: ["id", "author", "date", "text", "anchor"],
			properties: {
				id: { type: "string" },
				author: { type: "string" },
				initials: { type: "string" },
				date: { type: "string" },
				text: { type: "string" },
				parentId: { type: "string" },
				resolved: { type: "boolean" },
				anchor: {
					type: "object",
					required: ["startBlockId", "startOffset", "endBlockId", "endOffset"],
					properties: {
						startBlockId: { type: "string" },
						startOffset: { type: "number" },
						endBlockId: { type: "string" },
						endOffset: { type: "number" },
					},
				},
			},
		},
	},
};

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			json: { type: "boolean" },
			ts: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	if (parsed.values.ts) {
		await writeStdout(TS_SOURCE);
		return EXIT.OK;
	}

	await respond(JSON_SCHEMA);
	return EXIT.OK;
}
