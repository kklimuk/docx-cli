import XMLBuilder from "fast-xml-builder";
import { XMLParser } from "fast-xml-parser";

const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>
		<w:p>
			<w:r>
				<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>
				<w:t xml:space="preserve">Hello </w:t>
			</w:r>
			<w:r>
				<w:t>world</w:t>
			</w:r>
		</w:p>
	</w:body>
</w:document>`;

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	parseAttributeValue: false,
	parseTagValue: false,
	trimValues: false,
	processEntities: true,
	ignoreDeclaration: false,
});

const tree = parser.parse(xml);
console.log("Parsed tree:");
console.log(JSON.stringify(tree, null, 2));

const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	suppressEmptyNode: false,
	format: false,
});

const out = builder.build(tree);
console.log("\nRebuilt XML:");
console.log(out);
