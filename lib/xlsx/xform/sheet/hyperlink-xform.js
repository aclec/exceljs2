const BaseXform = require("../base-xform");

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
    // "#Sheet2!A1", "#'My Sheet'!A1:B2", "#DefinedName" and the legacy unprefixed "Sheet2!A1".
    static getLocation(target) {
        if (typeof target !== "string") {
            return undefined;
        }
        if (target.startsWith("#")) {
            return target.slice(1);
        }
        return /^[^!]+![a-zA-Z]+[\d]+$/.test(target) ? target : undefined;
    }

    static isInternalLink(target) {
        return HyperlinkXform.getLocation(target) !== undefined;
    }
}

module.exports = HyperlinkXform;
