import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { PageSetup } from "../../../../index";
import type { PageSetupModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";
import { isOneOf } from "../../../utils/utils";

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
// The enumerations of CT_PageSetup: a value that is not one of them reads as the default of the document
const ORIENTATIONS: Readonly<Record<PageSetup["orientation"], true>> = { default: true, portrait: true, landscape: true };
const PAGE_ORDERS: Readonly<Record<PageSetup["pageOrder"], true>> = { downThenOver: true, overThenDown: true };
const CELL_COMMENTS: Readonly<Record<PageSetup["cellComments"], true>> = { None: true, none: true, atEnd: true, asDisplayed: true };
const PRINT_ERRORS: Readonly<Record<PageSetup["errors"], true>> = { displayed: true, blank: true, dash: true, NA: true };
/**
 * `useFirstPageNumber`, written when the pages are numbered from a first page number: when one is set and the model
 * does not say that it is not used (a file read with `firstPageNumber` but without `useFirstPageNumber`), or when the
 * model says so (from 1, the CT_PageSetup default).
 */
function useFirstPageNumberToXml(model: PageSetupModel): "1" | undefined {
    const { firstPageNumber, useFirstPageNumber } = model;
    return useFirstPageNumber === true || (firstPageNumber !== undefined && useFirstPageNumber !== false) ? "1" : undefined;
}
/**
 * The first page number of a `<pageSetup>`: absent when the pages are numbered automatically and the file gives no
 * number; else the number (CT_PageSetup default 1) and whether it is used.
 */
function firstPageNumberToModel(attributes: SaxTag["attributes"]): Pick<PageSetupModel, "firstPageNumber" | "useFirstPageNumber"> {
    const used = parseBoolean(attributes.useFirstPageNumber);
    if (attributes.firstPageNumber === undefined && !used) {
        return {};
    }
    return { firstPageNumber: parseInt(attributes.firstPageNumber ?? "1", 10), useFirstPageNumber: used };
}

/**
 * `<pageSetup>`: the print settings of a worksheet (part of the page setup model, which also holds the margins and the
 * print options). Rendered when a setting is set (values other than the defaults for the enumerations); parsing gives
 * every setting, with Excel's defaults for the absent ones, but `firstPageNumber` and `useFirstPageNumber`: absent for
 * automatic page numbering without a first page number (see firstPageNumberToModel). An enumeration (`orientation`,
 * `pageOrder`, `cellComments`, `errors`) that is not one of its values reads as its default.
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
                useFirstPageNumber: useFirstPageNumberToXml(model),
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
                const { orientation, pageOrder, cellComments, errors } = attributes;
                this.model = {
                    paperSize: pageSizeToModel(attributes.paperSize),
                    orientation: isOneOf(ORIENTATIONS, orientation) ? orientation : "portrait",
                    horizontalDpi: parseInt(attributes.horizontalDpi || "4294967295", 10),
                    verticalDpi: parseInt(attributes.verticalDpi || "4294967295", 10),
                    pageOrder: isOneOf(PAGE_ORDERS, pageOrder) ? pageOrder : "downThenOver",
                    blackAndWhite: parseBoolean(attributes.blackAndWhite),
                    draft: parseBoolean(attributes.draft),
                    cellComments: isOneOf(CELL_COMMENTS, cellComments) ? cellComments : "None",
                    errors: isOneOf(PRINT_ERRORS, errors) ? errors : "displayed",
                    scale: parseInt(attributes.scale || "100", 10),
                    fitToWidth: parseInt(attributes.fitToWidth || "1", 10),
                    fitToHeight: parseInt(attributes.fitToHeight || "1", 10),
                    ...firstPageNumberToModel(attributes),
                    usePrinterDefaults: parseBoolean(attributes.usePrinterDefaults),
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
