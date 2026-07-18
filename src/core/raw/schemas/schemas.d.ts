// Allow `import xxx from "./path.xsd" with { type: "text" }` to return a string
// at runtime under Bun, mirroring the *.xml declaration in
// src/core/create/canonical/canonical.d.ts. The bundled ECMA-376 schema files
// in this folder are the only *.xsd imports in the project.
declare module "*.xsd" {
	const content: string;
	export default content;
}
