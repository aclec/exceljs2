import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ProtectionModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";

/**
 * The protection of a cell format, `<protection locked hidden>`. Only the values that differ from the defaults (locked,
 * not hidden) are written, and nothing when both are the defaults; parsing gives `null` when both are the defaults.
 */
class ProtectionXform extends BaseXform<ProtectionModel> {
    get tag(): string {
        return "protection";
    }

    override render(xmlStream: XmlStream, model: ProtectionModel): void {
        const locked = model.locked === undefined ? true : model.locked;
        const hidden = model.hidden === undefined ? false : model.hidden;
        if (!locked || hidden) {
            xmlStream.leafNode("protection", {
                locked: locked ? undefined : "0",
                hidden: hidden ? "1" : undefined,
            });
        }
    }

    /** Reads `<protection>` (returns nothing: the parent decides when to delegate to this xform). */
    override parseOpen(node: SaxTag): void {
        const model = {
            // xsd:boolean, defaults: locked, not hidden
            locked: parseBoolean(node.attributes.locked, true),
            hidden: parseBoolean(node.attributes.hidden),
        };

        // only want to record models that differ from defaults
        const isSignificant = !model.locked || model.hidden;

        this.model = isSignificant ? model : null;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default ProtectionXform;
