import type XmlStream from "../../../utils/xml-stream";
import type { RowBreak } from "../../../../index";
import PageBreaksXform from "./page-breaks-xform";

import ListXform from "../list-xform";

/**
 * `<rowBreaks count="n" manualBreakCount="n"><brk …/>…</rowBreaks>`: the manual page breaks between rows (index.d.ts
 * `RowBreak`, see PageBreaksXform). Nothing is rendered without a break.
 */
class RowBreaksXform extends ListXform<Partial<RowBreak> | string> {
    constructor() {
        super({
            tag: "rowBreaks",
            count: true,
            childXform: new PageBreaksXform(),
        });
    }

    override render(xmlStream: XmlStream, model: readonly Partial<RowBreak>[] | undefined): void {
        if (model && model.length) {
            xmlStream.openNode(this.tag, this.$);
            if (this.count) {
                xmlStream.addAttribute(this.$count, model.length);
                xmlStream.addAttribute("manualBreakCount", model.length);
            }

            const { childXform } = this;
            model.forEach((childModel) => {
                childXform.render(xmlStream, childModel);
            });

            xmlStream.closeNode();
        } else if (this.empty) {
            xmlStream.leafNode(this.tag);
        }
    }
}

export default RowBreaksXform;
