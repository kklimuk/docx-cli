import { describe, expect, test } from "bun:test";
import { marginalText } from "@core/marginals/text";
import { XmlNode } from "@core/parser";

describe("marginalText — text boxes in a header", () => {
	test("a letterhead text box's story joins the header text, bracketed", () => {
		const tree = XmlNode.parse(
			`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml">` +
				`<w:p><w:r><w:t>Header line</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp><wps:txbx><w:txbxContent>` +
				`<w:p><w:r><w:t>Acme Corporation letterhead</w:t></w:r></w:p><w:p><w:r><w:t>Registered office London.</w:t></w:r></w:p>` +
				`</w:txbxContent></wps:txbx></wps:wsp></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect><v:textbox><w:txbxContent>` +
				`<w:p><w:r><w:t>Acme Corporation letterhead</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:hdr>`,
		);
		expect(marginalText(tree)).toBe(
			"Header line\n[Acme Corporation letterhead\nRegistered office London.]",
		);
	});
});
