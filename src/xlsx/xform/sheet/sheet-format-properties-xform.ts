import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlValue } from "../../../utils/xml-stream";
import type { WorksheetPropertiesModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * The model of `<sheetFormatPr>`: part of the worksheet properties. Parsing gives every property (0 when absent) except
 * `defaultColWidth` (only when set).
 */
export type SheetFormatPropertiesModel = Pick<
    WorksheetPropertiesModel,
    "defaultRowHeight" | "dyDescent" | "outlineLevelRow" | "outlineLevelCol" | "defaultColWidth"
>;

/**
 * `<sheetFormatPr defaultRowHeight="15" outlineLevelRow="0" outlineLevelCol="0" x14ac:dyDescent="0.25"/>`: the default
 * row height and column width, and the outline levels of a worksheet. `customHeight` is written for a default row height
 * other than 15.
 */
class SheetFormatPropertiesXform extends BaseXform<SheetFormatPropertiesModel> {
    get tag(): string {
        return "sheetFormatPr";
    }

    override render(xmlStream: XmlStream, model: SheetFormatPropertiesModel | undefined): void {
        if (model) {
            const attributes: Record<string, XmlValue | undefined> = {
                defaultRowHeight: model.defaultRowHeight,
                outlineLevelRow: model.outlineLevelRow,
                outlineLevelCol: model.outlineLevelCol,
                "x14ac:dyDescent": model.dyDescent,
            };
            if (model.defaultColWidth) {
                attributes.defaultColWidth = model.defaultColWidth;
            }

            // default value for 'defaultRowHeight' is 15, this should not be 'custom'
            if (!model.defaultRowHeight || model.defaultRowHeight !== 15) {
                attributes.customHeight = "1";
            }

            if (Object.values(attributes).some((value) => value !== undefined)) {
                xmlStream.leafNode("sheetFormatPr", attributes);
            }
        }
    }

    /** @returns `true` for `<sheetFormatPr>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "sheetFormatPr") {
            const { attributes } = node;
            const model: SheetFormatPropertiesModel = (this.model = {
                defaultRowHeight: parseFloat(attributes.defaultRowHeight || "0"),
                dyDescent: parseFloat(attributes["x14ac:dyDescent"] || "0"),
                outlineLevelRow: parseInt(attributes.outlineLevelRow || "0", 10),
                outlineLevelCol: parseInt(attributes.outlineLevelCol || "0", 10),
            });
            if (attributes.defaultColWidth) {
                model.defaultColWidth = parseFloat(attributes.defaultColWidth);
            }
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default SheetFormatPropertiesXform;
