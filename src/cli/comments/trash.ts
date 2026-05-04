import { dirname, join } from "node:path";

const TRASH_VERSION = 1;
const TRASH_DIR = ".docx-cli";
const TRASH_FILE = "trash.json";

export type TrashAnchor = {
	startBlockId: string;
	startOffset: number;
	endBlockId: string;
	endOffset: number;
};

export type TrashEntry = {
	file: string;
	deletedAt: string;
	commentId: string;
	anchor: TrashAnchor;
	commentXml: string;
	extXml: string | null;
};

type TrashFile = {
	version: number;
	entries: TrashEntry[];
};

export function trashPathFor(docPath: string): string {
	return join(dirname(docPath), TRASH_DIR, TRASH_FILE);
}

export async function readTrash(docPath: string): Promise<TrashFile> {
	const path = trashPathFor(docPath);
	const file = Bun.file(path);
	if (!(await file.exists())) {
		return { version: TRASH_VERSION, entries: [] };
	}
	const parsed = (await file.json()) as TrashFile;
	if (parsed.version !== TRASH_VERSION) {
		return { version: TRASH_VERSION, entries: [] };
	}
	return parsed;
}

export async function writeTrash(
	docPath: string,
	trash: TrashFile,
): Promise<void> {
	const path = trashPathFor(docPath);
	await Bun.write(path, `${JSON.stringify(trash, null, 2)}\n`);
}

export async function pushTrashEntry(
	docPath: string,
	entry: TrashEntry,
): Promise<void> {
	const trash = await readTrash(docPath);
	trash.entries.push(entry);
	await writeTrash(docPath, trash);
}

export async function popTrashEntry(
	docPath: string,
	commentId: string,
): Promise<TrashEntry | undefined> {
	const trash = await readTrash(docPath);
	const fileName = docPath.split("/").pop() ?? docPath;
	for (let index = trash.entries.length - 1; index >= 0; index--) {
		const entry = trash.entries[index];
		if (!entry) continue;
		if (entry.file !== fileName) continue;
		if (entry.commentId !== commentId) continue;
		trash.entries.splice(index, 1);
		await writeTrash(docPath, trash);
		return entry;
	}
	return undefined;
}
