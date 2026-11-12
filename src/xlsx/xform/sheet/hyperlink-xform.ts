/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

// The legacy unprefixed "Sheet!A1" form of an internal link (exceljs 4 accepted it): a cell or a range of a sheet. A quoted
// sheet name is anything between single quotes ('' for a quote). An unquoted one cannot hold the characters of a path or of a
// URL scheme (: / \ ?, forbidden in sheet names), so that "https://example.com/#!page1" or "../book.xlsx#Sheet1!A1" stay
// external links; without a scheme, a # is part of the name ("Q#1!B2"). It may hold a ' but not start or end with one,
// like a sheet name.
const LEGACY_LOCATION =
    /^(?:'(?:[^']|'')+'|[^'!:/\\?*[\]](?:[^!:/\\?*[\]]*[^'!:/\\?*[\]])?)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

const FILE_BEFORE_HASH = /^[^#']*\.[A-Za-z0-9]{1,5}#/;

class HyperlinkXform extends BaseXform {
    get tag() {
        return "hyperlink";
    }

    render(xmlStream, model) {
        xmlStream.leafNode("hyperlink", {
            ref: model.address,
            "r:id": model.rId,
            tooltip: model.tooltip,
            location: HyperlinkXform.getLocation(model.target),
        });
    }

    parseOpen(node) {
        if (node.name === "hyperlink") {
            this.model = {
                address: node.attributes.ref,
                rId: node.attributes["r:id"],
                tooltip: node.attributes.tooltip,
            };

            // This is an internal link
            if (node.attributes.location) {
                this.model.target = node.attributes.location;
            }
            return true;
        }
        return false;
    }

    parseText() {}

    parseClose() {
        return false;
    }

    // An internal link (a place in this workbook) is written as a location, without a relationship.
    // Returns the location of an internal link target, undefined for an external target:
    // "#Sheet2!A1", "#'My Sheet'!A1:B2", "#DefinedName" and the legacy unprefixed "Sheet2!A1" (see LEGACY_LOCATION).
    static getLocation(target) {
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

    static isInternalLink(target) {
        return HyperlinkXform.getLocation(target) !== undefined;
    }
}

export default HyperlinkXform;
