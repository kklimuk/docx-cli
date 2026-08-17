/** Bun inlines `import … with { type: "text" }` at build time — including into a
 *  `bun build --compile` executable, which is how `docx upgrade` carries
 *  skills/docx-cli/scripts/install.sh without fetching it. tsc has no builtin
 *  knowledge of shell-script imports, so declare the shape here. */
declare module "*.sh" {
	const contents: string;
	export default contents;
}
