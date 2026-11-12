const BaseXform = require("../base-xform");

class VmlTextboxXform extends BaseXform {
    get tag() {
        return "v:textbox";
    }

    conversionUnit(value, multiple, unit) {
        return `${parseFloat(value) * multiple.toFixed(2)}${unit}`;
    }

    reverseConversionUnit(inset) {
        if (!inset) {
            // no inset attribute: VML default margins (Excel usually omits it)
            return undefined;
        }
        const margins = inset.split(",").map((margin) => {
            return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, "")).toFixed(2));
        });
        return margins.every(Number.isFinite) ? margins : undefined;
    }

    render(xmlStream, model) {
        const attributes = {
            style: "mso-direction-alt:auto",
        };
        if (model && model.note) {
            const { inset } = model.note && model.note.margins;
            if (Array.isArray(inset)) {
                // a NaN/Infinity margin would make the drawing unreadable: fall back to the VML default
                if (inset.every(Number.isFinite)) {
                    attributes.inset = inset.map((margin) => this.conversionUnit(margin, 10, "mm")).join(",");
                }
            } else if (inset) {
                attributes.inset = inset;
            }
        }
        xmlStream.openNode("v:textbox", attributes);
        xmlStream.leafNode("div", { style: "text-align:left" });
        xmlStream.closeNode();
    }

    parseOpen(node) {
        switch (node.name) {
            case this.tag:
                this.model = {
                    inset: this.reverseConversionUnit(node.attributes.inset),
                };
                return true;
            default:
                return true;
        }
    }

    parseText() {}

    parseClose(name) {
        switch (name) {
            case this.tag:
                return false;
            default:
                return true;
        }
    }
}

module.exports = VmlTextboxXform;
