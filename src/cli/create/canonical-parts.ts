import THEME1_XML from "./canonical/theme1.xml" with { type: "text" };

/** Canonical OOXML parts that Word treats as required even though the
 * spec marks them optional. Without these, Word shows an "unreadable
 * content / recover?" warning when opening the doc. LibreOffice is more
 * permissive and accepts the bare-minimum package without complaint —
 * which is why the gap went unnoticed until a user opened our output
 * in Word. */

export const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
\t<w:docDefaults>
\t\t<w:rPrDefault>
\t\t\t<w:rPr>
\t\t\t\t<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/>
\t\t\t\t<w:sz w:val="22"/>
\t\t\t\t<w:szCs w:val="22"/>
\t\t\t\t<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
\t\t\t</w:rPr>
\t\t</w:rPrDefault>
\t\t<w:pPrDefault>
\t\t\t<w:pPr>
\t\t\t\t<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
\t\t\t</w:pPr>
\t\t</w:pPrDefault>
\t</w:docDefaults>
\t<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
\t\t<w:name w:val="Normal"/>
\t\t<w:qFormat/>
\t</w:style>
</w:styles>`;

export const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
\t<w:zoom w:percent="100"/>
\t<w:defaultTabStop w:val="708"/>
\t<w:characterSpacingControl w:val="doNotCompress"/>
\t<w:compat>
\t\t<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
\t</w:compat>
</w:settings>`;

export const FONT_TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
\t<w:font w:name="Calibri">
\t\t<w:panose1 w:val="020F0502020204030204"/>
\t\t<w:charset w:val="00"/>
\t\t<w:family w:val="swiss"/>
\t\t<w:pitch w:val="variable"/>
\t</w:font>
\t<w:font w:name="Times New Roman">
\t\t<w:panose1 w:val="02020603050405020304"/>
\t\t<w:charset w:val="00"/>
\t\t<w:family w:val="roman"/>
\t\t<w:pitch w:val="variable"/>
\t</w:font>
\t<w:font w:name="Courier New">
\t\t<w:panose1 w:val="02070309020205020404"/>
\t\t<w:charset w:val="00"/>
\t\t<w:family w:val="modern"/>
\t\t<w:pitch w:val="fixed"/>
\t</w:font>
</w:fonts>`;

export const WEB_SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
\t<w:optimizeForBrowser/>
\t<w:allowPNG/>
</w:webSettings>`;

export const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
\t<Application>docx-cli</Application>
\t<DocSecurity>0</DocSecurity>
\t<ScaleCrop>false</ScaleCrop>
\t<LinksUpToDate>false</LinksUpToDate>
\t<SharedDoc>false</SharedDoc>
\t<HyperlinksChanged>false</HyperlinksChanged>
\t<AppVersion>1.0</AppVersion>
</Properties>`;

export const THEME_XML = THEME1_XML;

/** Per-part metadata: the zip path, content-type override, the relationship
 * type, and the target relative to either the package root (app.xml) or the
 * document part (everything in word/). Consumed by template.tsx and the
 * make-*-fixture scripts. */
export const CANONICAL_PARTS = {
	styles: {
		zipPath: "word/styles.xml",
		body: STYLES_XML,
		contentType:
			"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		target: "styles.xml",
		scope: "document",
	},
	settings: {
		zipPath: "word/settings.xml",
		body: SETTINGS_XML,
		contentType:
			"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
		target: "settings.xml",
		scope: "document",
	},
	webSettings: {
		zipPath: "word/webSettings.xml",
		body: WEB_SETTINGS_XML,
		contentType:
			"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
		target: "webSettings.xml",
		scope: "document",
	},
	fontTable: {
		zipPath: "word/fontTable.xml",
		body: FONT_TABLE_XML,
		contentType:
			"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
		target: "fontTable.xml",
		scope: "document",
	},
	theme: {
		zipPath: "word/theme/theme1.xml",
		body: THEME_XML,
		contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		target: "theme/theme1.xml",
		scope: "document",
	},
	app: {
		zipPath: "docProps/app.xml",
		body: APP_XML,
		contentType:
			"application/vnd.openxmlformats-officedocument.extended-properties+xml",
		relationshipType:
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
		target: "docProps/app.xml",
		scope: "package",
	},
} as const;

export type CanonicalPartKey = keyof typeof CANONICAL_PARTS;
