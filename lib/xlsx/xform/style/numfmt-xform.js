const _ = require("../../../utils/under-dash");
const defaultNumFormats = require("../../defaultnumformats");

const BaseXform = require("../base-xform");

// keyed by format code: no prototype, so that a format such as "constructor" is not found
function hashDefaultFormats() {
    const hash = Object.create(null);
    _.each(defaultNumFormats, (dnf, id) => {
        if (dnf.f) {
            hash[dnf.f] = parseInt(id, 10);
        }
        // at some point, add the other cultures here...
    });
    return hash;
}
const defaultFmtHash = hashDefaultFormats();

// NumFmt encapsulates translation between number format and xlsx
class NumFmtXform extends BaseXform {
    constructor(id, formatCode) {
        super();

        this.id = id;
        this.formatCode = formatCode;
    }

    get tag() {
        return "numFmt";
    }

    render(xmlStream, model) {
        xmlStream.leafNode("numFmt", { numFmtId: model.id, formatCode: model.formatCode });
    }

    parseOpen(node) {
        switch (node.name) {
            case "numFmt": {
                // kept as written: a backslash escapes a literal character ("\-", "\ ", "\m")
                const { formatCode } = node.attributes;
                this.model = {
                    id: parseInt(node.attributes.numFmtId, 10),
                    // Excel limits format codes to 255 characters: a longer one is not a format Excel wrote, and
                    // is ignored (its cells read as General) instead of being scanned for every cell using it
                    formatCode: formatCode.length <= NumFmtXform.MAX_FORMAT_CODE_LENGTH ? formatCode : undefined,
                };
                return true;
            }
            default:
                return false;
        }
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

NumFmtXform.MAX_FORMAT_CODE_LENGTH = 255;

NumFmtXform.getDefaultFmtId = function getDefaultFmtId(formatCode) {
    return defaultFmtHash[formatCode];
};

NumFmtXform.getDefaultFmtCode = function getDefaultFmtCode(numFmtId) {
    return defaultNumFormats[numFmtId] && defaultNumFormats[numFmtId].f;
};

module.exports = NumFmtXform;
