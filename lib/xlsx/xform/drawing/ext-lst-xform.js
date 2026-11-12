const BaseXform = require("../base-xform");

class ExtLstXform extends BaseXform {
    get tag() {
        return "a:extLst";
    }

    // a16:creationId must be unique per picture in the workbook:
    // derive a GUID-shaped id from the drawing number and the picture index
    static creationId(drawingNumber, index) {
        const hex = (value, length) => value.toString(16).toUpperCase().padStart(length, "0");
        return `{00000000-0008-0000-${hex(drawingNumber, 4)}-${hex(index, 12)}}`;
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag);
        xmlStream.openNode("a:ext", {
            uri: "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
        });
        xmlStream.leafNode("a16:creationId", {
            "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
            id: (model && model.creationId) || ExtLstXform.creationId(0, (model && model.index) || 0),
        });
        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    parseOpen(node) {
        switch (node.name) {
            case this.tag:
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
                // unprocessed internal nodes
                return true;
        }
    }
}

module.exports = ExtLstXform;
