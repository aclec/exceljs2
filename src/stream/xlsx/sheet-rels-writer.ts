/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import utils from "../../utils/utils";
import RelType from "../../xlsx/rel-type";
import HyperlinkXform from "../../xlsx/xform/sheet/hyperlink-xform";

class HyperlinksProxy {
    constructor(sheetRelsWriter) {
        this.writer = sheetRelsWriter;
    }

    push(hyperlink) {
        this.writer.addHyperlink(hyperlink);
    }
}

class SheetRelsWriter {
    constructor(options) {
        // in a workbook, each sheet will have a number
        this.id = options.id;

        // count of all relationships
        this.count = 0;

        // keep record of all hyperlinks
        this._hyperlinks = [];

        this._workbook = options.workbook;
    }

    get stream() {
        if (!this._stream) {
            this._stream = this._workbook._openStream(`/xl/worksheets/_rels/sheet${this.id}.xml.rels`);
        }
        return this._stream;
    }

    get length() {
        return this._hyperlinks.length;
    }

    each(fn) {
        return this._hyperlinks.forEach(fn);
    }

    get hyperlinksProxy() {
        return this._hyperlinksProxy || (this._hyperlinksProxy = new HyperlinksProxy(this));
    }

    addHyperlink(hyperlink) {
        const { address, target, tooltip } = hyperlink;
        // internal links are a location in the sheet, without a relationship
        const rId = HyperlinkXform.isInternalLink(target)
            ? undefined
            : this._writeRelationship({
                  Target: target,
                  Type: RelType.Hyperlink,
                  TargetMode: "External",
              });

        // store sheet stuff for later
        this._hyperlinks.push({ rId, address, target, tooltip });
    }

    addMedia(media) {
        return this._writeRelationship(media);
    }

    addRelationship(rel) {
        return this._writeRelationship(rel);
    }

    commit() {
        if (this.count) {
            // write xml utro
            this._writeClose();
            // and close stream
            this.stream.end();
        }
    }

    // ================================================================================
    _writeOpen() {
        this.stream.write(
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
        );
    }

    _writeRelationship(relationship) {
        if (!this.count) {
            this._writeOpen();
        }

        const rId = `rId${++this.count}`;

        if (relationship.TargetMode) {
            this.stream.write(
                `<Relationship Id="${rId}"` +
                    ` Type="${relationship.Type}"` +
                    ` Target="${utils.xmlEncode(relationship.Target)}"` +
                    ` TargetMode="${relationship.TargetMode}"` +
                    "/>"
            );
        } else {
            this.stream.write(`<Relationship Id="${rId}" Type="${relationship.Type}" Target="${utils.xmlEncode(relationship.Target)}"/>`);
        }

        return rId;
    }

    _writeClose() {
        this.stream.write("</Relationships>");
    }
}

export default SheetRelsWriter;
