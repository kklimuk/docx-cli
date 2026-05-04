import JSZip from "jszip";

export class PkgError extends Error {
	constructor(
		public code: string,
		message: string,
	) {
		super(message);
		this.name = "PkgError";
	}
}

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

	hasPart(name: string): boolean {
		return this.zip.file(name) !== null;
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
		await Bun.write(target, buf);
	}

	async toBytes(): Promise<Uint8Array> {
		return this.zip.generateAsync({
			type: "uint8array",
			compression: "DEFLATE",
			compressionOptions: { level: 6 },
		});
	}
}
