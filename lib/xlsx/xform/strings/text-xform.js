const BaseXform = require("../base-xform");

//   <t xml:space="preserve"> is </t>

// ECMA-376 ST_Xstring: "_xHHHH_" stands for the character U+HHHH (decoded by model). A literal one is written with its "_"
// escaped as "_x005F_", so that it reads back as written: every "_" that a reader would take as the start of an escape.
const ESCAPE_START = /_(?=x[0-9A-Fa-f]{4}_)/g;

class TextXform extends BaseXform {
    get tag() {
        return "t";
    }

    render(xmlStream, model) {
        xmlStream.openNode("t");
        if (/^\s|\n|\s$/.test(model)) {
            xmlStream.addAttribute("xml:space", "preserve");
        }
        xmlStream.writeText(typeof model === "string" ? model.replace(ESCAPE_START, "_x005F_") : model);
        xmlStream.closeNode();
    }

    get model() {
        return this._text.join("").replace(/_x([0-9A-F]{4})_/g, ($0, $1) => String.fromCharCode(parseInt($1, 16)));
    }

    parseOpen(node) {
        switch (node.name) {
            case "t":
                this._text = [];
                return true;
            default:
                return false;
        }
    }

    parseText(text) {
        this._text.push(text);
    }

    parseClose() {
        return false;
    }
}

module.exports = TextXform;
