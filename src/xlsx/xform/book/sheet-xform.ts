import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetDefModel } from "../../../types";
import { xmlDecode } from "../../../utils/utils";
import BaseXform from "../base-xform";

/**
 * A `<sheet sheetId name state r:id>` of `xl/workbook.xml` (see {@link SheetDefModel}). The name is XML-decoded when
 * parsed.
 */
class WorksheetXform extends BaseXform<SheetDefModel> {
    override render(xmlStream: XmlStream, model: SheetDefModel): void {
        xmlStream.leafNode("sheet", {
            sheetId: model.id,
            name: model.name,
            state: model.state,
            "r:id": model.rId,
        });
    }

    /** @returns `true` for `<sheet>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "sheet") {
            const { attributes } = node;
            this.model = {
                // name is required (CT_Sheet): without it, xmlDecode throws (the load fails)
                name: xmlDecode(attributes.name!),
                id: BaseXform.toIntValue(attributes.sheetId, NaN),
                state: attributes.state,
                rId: attributes["r:id"],
            };
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default WorksheetXform;
