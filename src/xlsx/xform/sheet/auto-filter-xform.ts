import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { AutoFilterModel } from "../../../types";
import colCache from "../../../utils/col-cache";
import BaseXform from "../base-xform";

/** A corner of an auto filter: an address (`"A1"`) or a 1-based `{ row, column }`. */
type AutoFilterCorner = Exclude<AutoFilterModel, string>["from"];

function getAddress(addr: AutoFilterCorner): string {
    if (typeof addr === "string") {
        return addr;
    }
    return colCache.getAddress(addr.row, addr.column).address;
}

/**
 * `<autoFilter ref="A1:C10"/>`: the auto filter of a worksheet. Rendered from a range string or from `{ from, to }`;
 * parsing gives the `ref` string (the filter columns inside are not read).
 */
class AutoFilterXform extends BaseXform<AutoFilterModel> {
    get tag(): string {
        return "autoFilter";
    }

    override render(xmlStream: XmlStream, model: AutoFilterModel | null | undefined): void {
        if (model) {
            if (typeof model === "string") {
                // assume range
                xmlStream.leafNode("autoFilter", { ref: model });
            } else {
                const firstAddress = getAddress(model.from);
                const secondAddress = getAddress(model.to);
                if (firstAddress && secondAddress) {
                    xmlStream.leafNode("autoFilter", { ref: `${firstAddress}:${secondAddress}` });
                }
            }
        }
    }

    /** Reads the `ref` of `<autoFilter>` (`undefined` when absent). Returns nothing: the parent decides when to delegate. */
    override parseOpen(node: SaxTag): void {
        if (node.name === "autoFilter") {
            this.model = node.attributes.ref;
        }
    }
}

export default AutoFilterXform;
