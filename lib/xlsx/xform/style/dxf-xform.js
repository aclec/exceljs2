const BaseXform = require("../base-xform");

const AlignmentXform = require("./alignment-xform");
const BorderXform = require("./border-xform");
const FillXform = require("./fill-xform");
const FontXform = require("./font-xform");
const NumFmtXform = require("./numfmt-xform");
const ProtectionXform = require("./protection-xform");

// A differential format (conditional formatting style): the parts of a cell style it changes
class DxfXform extends BaseXform {
    constructor() {
        super();

        this.map = {
            alignment: new AlignmentXform(),
            border: new BorderXform(),
            fill: new FillXform(),
            font: new FontXform(),
            numFmt: new NumFmtXform(),
            protection: new ProtectionXform(),
        };
    }

    get tag() {
        return "dxf";
    }

    // CT_Dxf sequence: font, numFmt, fill, alignment, protection, border, extLst
    render(xmlStream, model) {
        xmlStream.openNode(this.tag);

        if (model.font) {
            this.map.font.render(xmlStream, model.font);
        }
        if (model.numFmt && model.numFmtId) {
            const numFmtModel = { id: model.numFmtId, formatCode: model.numFmt };
            this.map.numFmt.render(xmlStream, numFmtModel);
        }
        if (model.fill) {
            this.map.fill.render(xmlStream, model.fill);
        }
        if (model.alignment) {
            this.map.alignment.render(xmlStream, model.alignment);
        }
        if (model.protection) {
            this.map.protection.render(xmlStream, model.protection);
        }
        if (model.border) {
            this.map.border.render(xmlStream, model.border);
        }

        xmlStream.closeNode();
    }

    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        switch (node.name) {
            case this.tag:
                // this node is often repeated. Need to reset children
                this.reset();
                return true;
            default:
                this.parser = this.map[node.name];
                if (this.parser) {
                    this.parser.parseOpen(node);
                }
                return true;
        }
    }

    parseText(text) {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    parseClose(name) {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        if (name === this.tag) {
            // like cell styles, a dxf style holds its number format as the format code string;
            // the numFmtId is reassigned by StylesXform.addDxfStyle when written
            const numFmt = this.map.numFmt.model;
            this.model = {
                alignment: this.map.alignment.model,
                border: this.map.border.model,
                fill: this.map.fill.model,
                font: this.map.font.model,
                numFmt: numFmt && numFmt.formatCode,
                protection: this.map.protection.model,
            };
            return false;
        }

        return true;
    }
}

module.exports = DxfXform;
