import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetHyperlinkModel } from "../../../types";
import BaseXform from "../base-xform";

// The legacy unprefixed "Sheet!A1" form of an internal link (exceljs 4 accepted it): a cell or a range of a sheet. A quoted
// sheet name is anything between single quotes ('' for a quote). An unquoted one cannot hold the characters of a path or of a
// URL scheme (: / \ ?, forbidden in sheet names), so that "https://example.com/#!page1" or "../book.xlsx#Sheet1!A1" stay
// external links; without a scheme, a # is part of the name ("Q#1!B2"). It may hold a ' but not start or end with one,
// like a sheet name.
const LEGACY_LOCATION =
    /^(?:'(?:[^']|'')+'|[^'!:/\\?*[\]](?:[^!:/\\?*[\]]*[^'!:/\\?*[\]])?)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

const FILE_BEFORE_HASH = /^[^#']*\.[A-Za-z0-9]{1,5}#/;

/**
 * A `<hyperlink>` of a worksheet (an item of `<hyperlinks>`). An external target is written as a relationship (`r:id`,
 * added by worksheet-xform prepare); an internal one (a place in this workbook, see {@link HyperlinkXform.getLocation}) as a
 * `location`.
 */
class HyperlinkXform extends BaseXform<SheetHyperlinkModel> {
    get tag(): string {
        return "hyperlink";
    }

    override render(xmlStream: XmlStream, model: SheetHyperlinkModel): void {
        xmlStream.leafNode("hyperlink", {
            ref: model.address,
            "r:id": model.rId,
            tooltip: model.tooltip,
            location: HyperlinkXform.getLocation(model.target),
        });
    }

    /** @returns `true` for `<hyperlink>`: the model is `{ address, rId, tooltip }`, plus `target` for a `location` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "hyperlink") {
            const model: SheetHyperlinkModel = (this.model = {
                address: node.attributes.ref,
                rId: node.attributes["r:id"],
                tooltip: node.attributes.tooltip,
            });

            // This is an internal link
            if (node.attributes.location) {
                model.target = node.attributes.location;
            }
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }

    /**
     * An internal link (a place in this workbook) is written as a location, without a relationship.
     *
     * @param target - the target of a hyperlink: `"#Sheet2!A1"`, `"#'My Sheet'!A1:B2"`, `"#DefinedName"`, the legacy
     *   unprefixed `"Sheet2!A1"` (see LEGACY_LOCATION), or an external target (URL, file)
     * @returns the location of an internal link target (without the `#`), `undefined` for an external target or a
     *   non-string
     */
    static getLocation(this: void, target: unknown): string | undefined {
        if (typeof target !== "string") {
            return undefined;
        }
        if (target.startsWith("#")) {
            return target.slice(1);
        }
        // "Book2.xlsx#Sheet1!A1" is a link to another workbook: a file name with an extension before the #
        if (FILE_BEFORE_HASH.test(target)) {
            return undefined;
        }
        return LEGACY_LOCATION.test(target) ? target : undefined;
    }

    /** @returns whether the target is a place in this workbook (see {@link HyperlinkXform.getLocation}) */
    static isInternalLink(this: void, target: unknown): boolean {
        return HyperlinkXform.getLocation(target) !== undefined;
    }
}

export default HyperlinkXform;
