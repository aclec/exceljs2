import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { WorksheetState } from "../../../../index";
import type { SheetDefModel } from "../../../types";
import { isOneOf, xmlDecode } from "../../../utils/utils";
import BaseXform from "../base-xform";

// ECMA-376 ST_SheetState
const SHEET_STATES: Readonly<Record<WorksheetState, true>> = { visible: true, hidden: true, veryHidden: true };

/**
 * A `<sheet sheetId name state r:id>` of `xl/workbook.xml` (see {@link SheetDefModel}). The name is XML-decoded when
 * parsed; a `state` that is not one of ST_SheetState is dropped (visible).
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

    /**
     * @returns `true` for `<sheet>`
     * @throws Error for a sheet without name (CT_Sheet requires it)
     */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "sheet") {
            const { name, sheetId, state } = node.attributes;
            if (name === undefined) {
                throw new Error("Invalid xl/workbook.xml: <sheet> without name");
            }
            this.model = {
                name: xmlDecode(name),
                id: BaseXform.toIntValue(sheetId, NaN),
                state: isOneOf(SHEET_STATES, state) ? state : undefined,
                rId: node.attributes["r:id"],
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
