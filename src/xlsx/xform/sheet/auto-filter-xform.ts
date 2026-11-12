/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import colCache from "../../../utils/col-cache";
import BaseXform from "../base-xform";

class AutoFilterXform extends BaseXform {
    get tag() {
        return "autoFilter";
    }

    render(xmlStream, model) {
        if (model) {
            if (typeof model === "string") {
                // assume range
                xmlStream.leafNode("autoFilter", { ref: model });
            } else {
                const getAddress = function (addr) {
                    if (typeof addr === "string") {
                        return addr;
                    }
                    return colCache.getAddress(addr.row, addr.column).address;
                };

                const firstAddress = getAddress(model.from);
                const secondAddress = getAddress(model.to);
                if (firstAddress && secondAddress) {
                    xmlStream.leafNode("autoFilter", { ref: `${firstAddress}:${secondAddress}` });
                }
            }
        }
    }

    parseOpen(node) {
        if (node.name === "autoFilter") {
            this.model = node.attributes.ref;
        }
    }
}

export default AutoFilterXform;
