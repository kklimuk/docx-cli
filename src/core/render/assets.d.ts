/** Ambient type for `.wasm` files imported with `{ type: "file" }` — Bun
 * resolves the import to a file-path string at build time (the on-disk
 * node_modules path under normal `bun` invocation, an embedded
 * `/$bunfs/root/…` path under `bun build --compile`). The actual bytes are
 * loaded at runtime via `Bun.file(path).arrayBuffer()`. */
declare module "*.wasm" {
	const path: string;
	export default path;
}
