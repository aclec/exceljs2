import type XmlStream from "../../../utils/xml-stream";
import type { RowBreakModel } from "../../../types";
import colCache from "../../../utils/col-cache";
import PageBreaksXform from "./page-breaks-xform";

import ListXform from "../list-xform";

/**
 * `<rowBreaks count="n" manualBreakCount="m"><brk …/>…</rowBreaks>`: the page breaks between rows (see
 * PageBreaksXform). Nothing is rendered without a break. Rendering fails beyond {@link RowBreaksXform.MAX_BREAKS} breaks
 * (Excel's limit); parsing reads up to one break per row (a bound on hostile input), more fail.
 */
class RowBreaksXform extends ListXform<RowBreakModel> {
    /**
     * The maximum number of breaks of a worksheet: Office requires `count` and `manualBreakCount` to be at most 1023
     * ([MS-OI29500] 2.1.599, the note on ECMA-376 Part 1 18.3.1.14 colBreaks, which refers to 18.3.1.74 rowBreaks).
     */
    static readonly MAX_BREAKS = 1023;

    constructor() {
        super({
            tag: "rowBreaks",
            count: true,
            childXform: new PageBreaksXform(),
            maxItems: colCache.MAX_ROWS,
        });
    }

    /** @throws Error for more than {@link RowBreaksXform.MAX_BREAKS} breaks, which Excel does not open */
    override render(xmlStream: XmlStream, model: readonly RowBreakModel[] | null | undefined): void {
        if (model && model.length > RowBreaksXform.MAX_BREAKS) {
            throw new Error(`Too many row breaks in a worksheet: ${model.length} (Excel allows at most ${RowBreaksXform.MAX_BREAKS})`);
        }
        if (model && model.length) {
            xmlStream.openNode(this.tag, this.$);
            if (this.count) {
                xmlStream.addAttribute(this.$count, model.length);
                xmlStream.addAttribute("manualBreakCount", model.filter((brk) => brk.man).length);
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
