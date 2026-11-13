import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CalculationProperties } from "../../../../index";
import BaseXform from "../base-xform";

/** `<calcPr>` of `xl/workbook.xml`: `fullCalcOnLoad` is written; nothing is parsed (the model is `{}`). */
class WorkbookCalcPropertiesXform extends BaseXform<Partial<CalculationProperties>> {
    override render(xmlStream: XmlStream, model: Partial<CalculationProperties>): void {
        xmlStream.leafNode("calcPr", {
            calcId: 171027,
            fullCalcOnLoad: model.fullCalcOnLoad ? 1 : undefined,
        });
    }

    /** @returns `true` for `<calcPr>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "calcPr") {
            this.model = {};
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default WorkbookCalcPropertiesXform;
