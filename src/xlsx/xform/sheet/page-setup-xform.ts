import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { PageSetup } from "../../../../index";
import type { PageSetupModel } from "../../../types";
import BaseXform from "../base-xform";

function booleanToXml(model: unknown): "1" | undefined {
    return model ? "1" : undefined;
}
function pageOrderToXml(model: string | undefined): "overThenDown" | undefined {
    switch (model) {
        case "overThenDown":
            return model;
        case undefined:
        default:
            return undefined;
    }
}
// ECMA-376 ST_CellComments ("none", the default, is not written). "asDisplyed" (sic) was the only spelling accepted (and
// written, an invalid value) by earlier versions: it is still accepted
function cellCommentsToXml(model: string | undefined): "atEnd" | "asDisplayed" | undefined {
    switch (model) {
        case "atEnd":
        case "asDisplayed":
            return model;
        case "asDisplyed":
            return "asDisplayed";
        case undefined:
        default:
            return undefined;
    }
}
function errorsToXml(model: string | undefined): "dash" | "blank" | "NA" | undefined {
    switch (model) {
        case "dash":
        case "blank":
        case "NA":
            return model;
        case undefined:
        default:
            return undefined;
    }
}
function pageSizeToModel(value: string | undefined): number | undefined {
    return value !== undefined ? parseInt(value, 10) : undefined;
}

/**
 * `<pageSetup>`: the print settings of a worksheet (part of the page setup model, which also holds the margins and the
 * print options). Rendered when a setting is set (values other than the defaults for the enumerations); parsing gives
 * every setting, with Excel's defaults for the absent ones. Parsed literal unions hold the attribute strings of the file,
 * unvalidated.
 */
class PageSetupXform extends BaseXform<PageSetupModel> {
    get tag(): string {
        return "pageSetup";
    }

    override render(xmlStream: XmlStream, model: PageSetupModel | undefined): void {
        if (model) {
            const attributes = {
                paperSize: model.paperSize,
                orientation: model.orientation,
                horizontalDpi: model.horizontalDpi,
                verticalDpi: model.verticalDpi,
                pageOrder: pageOrderToXml(model.pageOrder),
                blackAndWhite: booleanToXml(model.blackAndWhite),
                draft: booleanToXml(model.draft),
                cellComments: cellCommentsToXml(model.cellComments),
                errors: errorsToXml(model.errors),
                scale: model.scale,
                fitToWidth: model.fitToWidth,
                fitToHeight: model.fitToHeight,
                firstPageNumber: model.firstPageNumber,
                useFirstPageNumber: booleanToXml(model.firstPageNumber),
                usePrinterDefaults: booleanToXml(model.usePrinterDefaults),
                copies: model.copies,
            };
            if (Object.values(attributes).some((value) => value !== undefined)) {
                xmlStream.leafNode(this.tag, attributes);
            }
        }
    }

    /** @returns `true` for `<pageSetup>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag: {
                const { attributes } = node;
                // the literal unions hold the attribute strings of the file (unvalidated)
                this.model = {
                    paperSize: pageSizeToModel(attributes.paperSize),
                    orientation: (attributes.orientation || "portrait") as PageSetup["orientation"],
                    horizontalDpi: parseInt(attributes.horizontalDpi || "4294967295", 10),
                    verticalDpi: parseInt(attributes.verticalDpi || "4294967295", 10),
                    pageOrder: (attributes.pageOrder || "downThenOver") as PageSetup["pageOrder"],
                    blackAndWhite: attributes.blackAndWhite === "1",
                    draft: attributes.draft === "1",
                    cellComments: (attributes.cellComments || "None") as PageSetup["cellComments"],
                    errors: (attributes.errors || "displayed") as PageSetup["errors"],
                    scale: parseInt(attributes.scale || "100", 10),
                    fitToWidth: parseInt(attributes.fitToWidth || "1", 10),
                    fitToHeight: parseInt(attributes.fitToHeight || "1", 10),
                    firstPageNumber: parseInt(attributes.firstPageNumber || "1", 10),
                    useFirstPageNumber: attributes.useFirstPageNumber === "1",
                    usePrinterDefaults: attributes.usePrinterDefaults === "1",
                    copies: parseInt(attributes.copies || "1", 10),
                };
                return true;
            }
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PageSetupXform;
