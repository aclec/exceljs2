import { COL_BREAK_LIMITS } from "./page-breaks-xform";
import { PageBreakListXform } from "./row-breaks-xform";

/**
 * `<colBreaks>`: the page breaks between columns, `<brk id="3" max="1048575" man="1"/>` (`id` the column number, `min` /
 * `max` 0-based rows). See {@link PageBreakListXform}: at most 1023 breaks are rendered, and one per column (16,384)
 * parsed.
 */
class ColBreaksXform extends PageBreakListXform {
    constructor() {
        super("colBreaks", "column", COL_BREAK_LIMITS);
    }
}

export default ColBreaksXform;
