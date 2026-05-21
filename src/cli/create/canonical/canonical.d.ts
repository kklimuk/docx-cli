// Allow `import xxx from "./path.xml" with { type: "text" }` to return a string
// at runtime under Bun. TypeScript doesn't know about Bun's text-import attribute
// for non-TS extensions, so we declare the shape here. Ambient module
// declarations are global once tsconfig picks them up — this lives next to
// the canonical theme1.xml import but applies to any *.xml import in the
// project. Today only canonical-parts.ts uses it.
declare module "*.xml" {
	const content: string;
	export default content;
}
