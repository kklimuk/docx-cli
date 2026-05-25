import { w, w14 } from "../jsx";
import type { XmlNode } from "../parser";

/** Emit the `<w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr><w:sdtContent>…☐|☒…</w:sdtContent></w:sdt>`
 *  subtree that materializes a GFM task-list checkbox — Word, Pandoc, and
 *  LibreOffice all emit this exact shape. The `w14:checked` value and the
 *  glyph in `<w:sdtContent>` co-vary (☐ = unchecked, ☒ = checked). Used by
 *  `<Paragraph taskState={...}>` in [core/blocks.tsx](../blocks.tsx) to
 *  prepend the SDT to a task-list paragraph's runs. */
export function TaskCheckbox({ checked }: { checked: boolean }): XmlNode {
	const glyph = checked ? "☒" : "☐";
	return (
		<w.sdt>
			<w.sdtPr>
				<w.id w-val={String(allocateSdtId())} />
				<w14.checkbox>
					<w14.checked w14-val={checked ? "1" : "0"} />
					<w14.checkedState w14-val="2612" />
					<w14.uncheckedState w14-val="2610" />
				</w14.checkbox>
			</w.sdtPr>
			<w.sdtContent>
				<w.r />
				<w.r>
					<w.t>{glyph}</w.t>
				</w.r>
			</w.sdtContent>
		</w.sdt>
	);
}

/** Word's SDT ids are 32-bit signed ints that mostly serve as a uniqueness
 *  key; Word doesn't error on duplicates but having them distinct keeps
 *  comparison tools cleaner. We're not coordinating with other consumers in
 *  the same package, so a process-local counter is enough. */
let nextSdtId = 1;
function allocateSdtId(): number {
	return nextSdtId++;
}
