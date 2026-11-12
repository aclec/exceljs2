/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";
import SharedStringXform from "./shared-string-xform";

class SharedStringsXform extends BaseXform {
    constructor(model) {
        super();

        this.model = model || {
            values: [],
            count: 0,
        };
        this.hash = Object.create(null);
        this.rich = Object.create(null);
        // index of each rich text object already added: the cells of a loaded workbook that reference the same
        // shared string share its object, and serialising it again for each of them was O(cells x runs)
        this.richIndexes = new WeakMap();
    }

    get sharedStringXform() {
        return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
    }

    get values() {
        return this.model.values;
    }

    get uniqueCount() {
        return this.model.values.length;
    }

    get count() {
        return this.model.count;
    }

    getString(index) {
        return this.model.values[index];
    }

    add(value) {
        return value.richText ? this.addRichText(value) : this.addText(value);
    }

    addText(value) {
        let index = this.hash[value];
        if (index === undefined) {
            index = this.hash[value] = this.model.values.length;
            this.model.values.push(value);
        }
        this.model.count++;
        return index;
    }

    addRichText(value) {
        let index = this.richIndexes.get(value);
        if (index === undefined) {
            const xml = this.sharedStringXform.toXml(value);
            index = this.rich[xml];
            if (index === undefined) {
                index = this.rich[xml] = this.model.values.length;
                this.model.values.push(value);
            }
            this.richIndexes.set(value, index);
        }
        this.model.count++;
        return index;
    }

    // model: {values, count}, count being the number of references to the shared strings in the
    // workbook (cells), the values its unique strings
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
    //   <si><t><%=text%></t></si>
    //   <si><r><rPr></rPr><t></t></r></si>
    // </sst>

    render(xmlStream, model) {
        model = model || this._values;
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("sst", {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            count: model.count,
            uniqueCount: model.values.length,
        });

        const sx = this.sharedStringXform;
        model.values.forEach((sharedString) => {
            sx.render(xmlStream, sharedString);
        });
        xmlStream.closeNode();
    }

    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "sst":
                return true;
            case "si":
                this.parser = this.sharedStringXform;
                this.parser.parseOpen(node);
                return true;
            default:
                throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
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
                this.model.values.push(this.parser.model);
                this.model.count++;
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "sst":
                return false;
            default:
                throw new Error(`Unexpected xml node in parseClose: ${name}`);
        }
    }
}

export default SharedStringsXform;
