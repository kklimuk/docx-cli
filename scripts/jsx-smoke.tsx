import { Fragment, w } from "@core/jsx";
import { XmlNode } from "@core/parser";

void Fragment;

// Two ways to write attributes with colons:
const paragraph = (
	<w.p>
		<w.r>
			<w.rPr>
				<w.b />
				{/* (1) hyphen shortcut: w-val -> w:val */}
				<w.color w-val="800080" />
				{/* (2) JSX spread of an object with literal colon keys */}
				<w.sz {...{ "w:val": "24" }} />
			</w.rPr>
			<w.t {...{ "xml:space": "preserve" }}>Hello </w.t>
		</w.r>
		<w.r>
			<w.t>world</w.t>
		</w.r>
	</w.p>
);

if (Array.isArray(paragraph)) {
	throw new Error("Expected single node");
}

const document = (
	<>
		<w.document
			{...{
				"xmlns:w":
					"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			}}
		>
			<w.body>
				{paragraph}
				<w.sectPr />
			</w.body>
		</w.document>
	</>
);

const tree = Array.isArray(document) ? document : [document];
console.log(XmlNode.serialize(tree));
