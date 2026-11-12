import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { PageSetup } from "../../../../index";
import BaseXform from "../base-xform";

/** The print options of a worksheet: part of the page setup (index.d.ts `PageSetup`). */
export type PrintOptionsModel = Pick<PageSetup, "showRowColHeaders" | "showGridLines" | "horizontalCentered" | "verticalCentered">;

function booleanToXml(model: boolean | undefined): "1" | undefined {
    return model ? "1" : undefined;
}

/**
 * `<printOptions headings="1" gridLines="1" horizontalCentered="1" verticalCentered="1"/>`: rendered when an option is
 * set (the model may be the whole page setup); parsing gives every option.
 */
class PrintOptionsXform extends BaseXform<PrintOptionsModel> {
    get tag(): string {
        return "printOptions";
    }

    override render(xmlStream: XmlStream, model: Partial<PrintOptionsModel> | undefined): void {
        if (model) {
            const headings = booleanToXml(model.showRowColHeaders);
            const gridLines = booleanToXml(model.showGridLines);
            const horizontalCentered = booleanToXml(model.horizontalCentered);
            const verticalCentered = booleanToXml(model.verticalCentered);
            if (headings || gridLines || horizontalCentered || verticalCentered) {
                xmlStream.leafNode(this.tag, { headings, gridLines, horizontalCentered, verticalCentered });
            }
        }
    }

    /** @returns `true` for `<printOptions>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = {
                    showRowColHeaders: node.attributes.headings === "1",
                    showGridLines: node.attributes.gridLines === "1",
                    horizontalCentered: node.attributes.horizontalCentered === "1",
                    verticalCentered: node.attributes.verticalCentered === "1",
                };
                return true;
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PrintOptionsXform;
