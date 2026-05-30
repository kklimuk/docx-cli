import { lstat, rename, unlink } from "node:fs/promises";
import JSZip from "jszip";
import { XmlNode } from "../../parser";

export class Pkg {
	private constructor(
		private zip: JSZip,
		public readonly path: string,
	) {}

	static async open(path: string): Promise<Pkg> {
		const file = Bun.file(path);
		if (!(await file.exists())) {
			throw new PkgError("FILE_NOT_FOUND", `File not found: ${path}`);
		}
		const buf = await file.arrayBuffer();
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(buf);
		} catch (err) {
			const reason = err instanceof Error ? err.message : String(err);
			throw new PkgError("NOT_A_ZIP", `Not a valid .docx (zip): ${reason}`);
		}
		return new Pkg(zip, path);
	}

	/** Build an empty package targeting `path`. Callers populate parts via
	 * `writeText` / `writeBytes` and persist with `save()`. Used by `docx
	 * create` to assemble a fresh `.docx` from template strings without
	 * reaching for JSZip directly. */
	static empty(path: string): Pkg {
		return new Pkg(new JSZip(), path);
	}

	hasPart(name: string): boolean {
		return this.zip.file(name) !== null;
	}

	async readPart(name: string): Promise<XmlNode[] | undefined> {
		if (!this.hasPart(name)) return;
		return XmlNode.parse(await this.readText(name));
	}

	async ensurePart(name: string): Promise<XmlNode[]> {
		return (await this.readPart(name)) ?? [XmlNode.textNode("Empty")];
	}

	listParts(): string[] {
		const names: string[] = [];
		this.zip.forEach((p) => {
			names.push(p);
		});
		return names;
	}

	async readText(name: string): Promise<string> {
		const entry = this.zip.file(name);
		if (!entry) {
			throw new PkgError("PART_NOT_FOUND", `Part not found: ${name}`);
		}
		return entry.async("string");
	}

	async readBytes(name: string): Promise<Uint8Array> {
		const entry = this.zip.file(name);
		if (!entry) {
			throw new PkgError("PART_NOT_FOUND", `Part not found: ${name}`);
		}
		return entry.async("uint8array");
	}

	writeText(name: string, content: string): void {
		this.zip.file(name, content);
	}

	writeBytes(name: string, content: Uint8Array): void {
		this.zip.file(name, content);
	}

	deletePart(name: string): void {
		this.zip.remove(name);
	}

	async save(path?: string): Promise<void> {
		const target = path ?? this.path;
		const buf = await this.zip.generateAsync({
			type: "uint8array",
			compression: "DEFLATE",
			compressionOptions: { level: 6 },
		});
		await this.#writeAtomic(target, buf);
	}

	async #writeAtomic(target: string, buf: Uint8Array): Promise<void> {
		// If the target is a symlink (e.g. a .docx pointed at a cloud-sync folder),
		// rename() would replace the link with a regular file and break the link.
		// Write through the symlink instead, accepting non-atomicity for that case.
		let isSymlink = false;
		try {
			isSymlink = (await lstat(target)).isSymbolicLink();
		} catch {
			// Target doesn't exist; not a symlink.
		}
		if (isSymlink) {
			await Bun.write(target, buf);
			return;
		}

		const tmp = `${target}.docx-cli-tmp-${process.pid}-${Date.now()}`;
		try {
			await Bun.write(tmp, buf);
			await rename(tmp, target);
		} catch (err) {
			await unlink(tmp).catch(() => {});
			throw err;
		}
	}

	async toBytes(): Promise<Uint8Array> {
		return this.zip.generateAsync({
			type: "uint8array",
			compression: "DEFLATE",
			compressionOptions: { level: 6 },
		});
	}
}

export class PkgError extends Error {
	constructor(
		public code: string,
		message: string,
	) {
		super(message);
		this.name = "PkgError";
	}
}
