import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RowBreak } from "../../../../index";
import BaseXform from "../base-xform";

/**
 * A page break, `<brk id="5" max="16383" man="1"/>` (ECMA-376 CT_Break), rendered from index.d.ts `RowBreak` (its
 * properties are written as the attributes). Parsing keeps the `ref` attribute, which CT_Break does not have (so
 * `undefined`): the breaks are not read back.
 */
class PageBreaksXform extends BaseXform<Partial<RowBreak> | string> {
    get tag(): string {
        return "brk";
    }

    override render(xmlStream: XmlStream, model: Partial<RowBreak>): void {
        xmlStream.leafNode("brk", model);
    }

    /** @returns `true` for `<brk>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "brk") {
            this.model = node.attributes.ref;
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PageBreaksXform;
