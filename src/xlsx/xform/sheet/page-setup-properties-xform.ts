/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class PageSetupPropertiesXform extends BaseXform {
    get tag() {
        return "pageSetUpPr";
    }

    render(xmlStream, model) {
        if (model && model.fitToPage) {
            xmlStream.leafNode(this.tag, {
                fitToPage: model.fitToPage ? "1" : undefined,
            });
            return true;
        }
        return false;
    }

    parseOpen(node) {
        if (node.name === this.tag) {
            this.model = {
                fitToPage: node.attributes.fitToPage === "1",
            };
            return true;
        }
        return false;
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

export default PageSetupPropertiesXform;
