import type XmlStream from "../../../utils/xml-stream";
import PageBreaksXform, { ROW_BREAK_LIMITS, type PageBreakLimits, type PageBreakModel } from "./page-breaks-xform";

import ListXform from "../list-xform";

/**
 * A list of page breaks (ECMA-376 CT_PageBreak), `<rowBreaks>` or `<colBreaks>`: `<tag count="n" manualBreakCount="m"><brk
 * …/>…</tag>` (see PageBreaksXform). Nothing is rendered without a break. Rendering fails beyond
 * {@link PageBreakListXform.MAX_BREAKS} breaks (Excel's limit); parsing reads up to one break per row or column (a bound
 * on hostile input), more fail.
 */
export class PageBreakListXform extends ListXform<PageBreakModel> {
    /**
     * The maximum number of breaks of a worksheet: Office requires `count` and `manualBreakCount` to be at most 1023
     * ([MS-OI29500] 2.1.599, the note on ECMA-376 Part 1 18.3.1.14 colBreaks, which refers to 18.3.1.74 rowBreaks).
     */
    static readonly MAX_BREAKS = 1023;

    /** What the breaks are between, in the error messages: `"row"` or `"column"`. */
    readonly kind: string;

    /**
     * @param tag - `rowBreaks` or `colBreaks`
     * @param kind - `"row"` or `"column"` (the error messages)
     * @param limits - the limits of the attributes of the breaks; the maximum number of breaks parsed is `limits.lastId`
     */
    constructor(tag: string, kind: string, limits: PageBreakLimits) {
        super({
            tag,
            count: true,
            childXform: new PageBreaksXform(limits),
            maxItems: limits.lastId,
        });
        this.kind = kind;
    }

    /** @throws Error for more than {@link PageBreakListXform.MAX_BREAKS} breaks, which Excel does not open */
    override render(xmlStream: XmlStream, model: readonly PageBreakModel[] | null | undefined): void {
        if (model && model.length > PageBreakListXform.MAX_BREAKS) {
            throw new Error(
                `Too many ${this.kind} breaks in a worksheet: ${model.length} (Excel allows at most ${PageBreakListXform.MAX_BREAKS})`
            );
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

/** `<rowBreaks>`: the page breaks between rows (see {@link PageBreakListXform}). */
class RowBreaksXform extends PageBreakListXform {
    constructor() {
        super("rowBreaks", "row", ROW_BREAK_LIMITS);
    }
}

export default RowBreaksXform;
