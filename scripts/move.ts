#!/usr/bin/env bun
/**
 * Move a TypeScript file and update all import references across the project.
 * Uses the same engine as VS Code's "update imports on file move".
 *
 * Usage:
 *   bun scripts/move.ts <source> <destination>
 *
 * Examples:
 *   bun scripts/move.ts src/client/Page.tsx src/client/PageView.tsx
 *   bun scripts/move.ts src/shared/slugify.ts src/shared/utils/slugify.ts
 */

import { dirname, relative, resolve } from "path";
import * as ts from "typescript";

const [source, dest] = Bun.argv.slice(2);
if (!source || !dest) {
	console.error("Usage: bun scripts/move.ts <source> <destination>");
	throw new Error("Missing arguments");
}

const root = resolve(import.meta.dir, "..");
const oldPath = resolve(root, source);
const newPath = resolve(root, dest);

if (!(await Bun.file(oldPath).exists())) {
	throw new Error(`Source file not found: ${oldPath}`);
}

// Load tsconfig
const configPath = ts.findConfigFile(root, ts.sys.fileExists, "tsconfig.json");
if (!configPath) throw new Error("tsconfig.json not found");

const configFile = ts.readConfigFile(configPath, ts.sys.readFile);
const parsedConfig = ts.parseJsonConfigFileContent(
	configFile.config,
	ts.sys,
	root,
);

// Build a LanguageServiceHost
const files = new Map<string, { version: number; content: string }>();
for (const fileName of parsedConfig.fileNames) {
	const content = ts.sys.readFile(fileName) ?? "";
	files.set(fileName, { version: 0, content });
}

const host: ts.LanguageServiceHost = {
	getScriptFileNames: () => [...files.keys()],
	getScriptVersion: (fileName) => String(files.get(fileName)?.version ?? 0),
	getScriptSnapshot: (fileName) => {
		const entry = files.get(fileName);
		if (entry) return ts.ScriptSnapshot.fromString(entry.content);
		const content = ts.sys.readFile(fileName);
		if (content !== undefined) return ts.ScriptSnapshot.fromString(content);
		return undefined;
	},
	getCurrentDirectory: () => root,
	getCompilationSettings: () => parsedConfig.options,
	getDefaultLibFileName: ts.getDefaultLibFilePath,
	fileExists: ts.sys.fileExists,
	readFile: ts.sys.readFile,
	readDirectory: ts.sys.readDirectory,
	directoryExists: ts.sys.directoryExists,
	getDirectories: ts.sys.getDirectories,
};

const service = ts.createLanguageService(host, ts.createDocumentRegistry());
const edits = service.getEditsForFileRename(oldPath, newPath, {}, undefined);

if (edits.length === 0) {
	console.log("No import updates needed.");
} else {
	for (const fileChange of edits) {
		const filePath = fileChange.fileName;
		let content = await Bun.file(filePath).text();

		const sortedEdits = [...fileChange.textChanges].sort(
			(a, b) => b.span.start - a.span.start,
		);

		for (const change of sortedEdits) {
			content =
				content.slice(0, change.span.start) +
				change.newText +
				content.slice(change.span.start + change.span.length);
		}

		await Bun.write(filePath, content);
		console.log(`  Updated imports in ${relative(root, filePath)}`);
	}
}

// Move the file on disk
const { mkdirSync, unlinkSync } = await import("fs");
mkdirSync(dirname(newPath), { recursive: true });
await Bun.write(newPath, Bun.file(oldPath));
unlinkSync(oldPath);
console.log(`  Moved ${relative(root, oldPath)} -> ${relative(root, newPath)}`);
