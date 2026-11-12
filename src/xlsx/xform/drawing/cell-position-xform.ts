import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { AnchorModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import IntegerXform from "../simple/integer-xform";
import colCache from "../../../utils/col-cache";

// a value of a corner: 0 when its element is absent or not a number (CT_Marker requires the four)
function coordinate(value: number | null | undefined): number {
    return typeof value === "number" && !Number.isNaN(value) ? value : 0;
}

// the cell of an anchor is in the worksheet grid (0-based): a huge index from the file would make readers of the anchor
// (Anchor col/row) work up to it
function cell(value: number | null | undefined, max: number): number {
    return Math.min(Math.max(coordinate(value), 0), max);
}

/** The options of a {@link CellPositionXform}. */
export interface CellPositionXformOptions {
    /** The element: `xdr:from` or `xdr:to`. */
    readonly tag: string;
}

/**
 * A corner of an anchor, `<xdr:from>` / `<xdr:to>`: its 0-based column and row, and its offsets in EMU, in this order
 * (CT_Marker). Parsing clamps the column and row to the grid; a value whose element is absent (or not a number) is 0.
 */
class CellPositionXform extends BaseXform<AnchorModel> {
    /** The element. */
    readonly tag: string;

    declare map: {
        readonly "xdr:col": IntegerXform;
        readonly "xdr:colOff": IntegerXform;
        readonly "xdr:row": IntegerXform;
        readonly "xdr:rowOff": IntegerXform;
    };

    constructor(options: CellPositionXformOptions) {
        super();

        this.tag = options.tag;
        this.map = {
            "xdr:col": new IntegerXform({ tag: "xdr:col", zero: true }),
            "xdr:colOff": new IntegerXform({ tag: "xdr:colOff", zero: true }),
            "xdr:row": new IntegerXform({ tag: "xdr:row", zero: true }),
            "xdr:rowOff": new IntegerXform({ tag: "xdr:rowOff", zero: true }),
        };
    }

    override render(xmlStream: XmlStream, model: AnchorModel): void {
        xmlStream.openNode(this.tag);

        this.map["xdr:col"].render(xmlStream, model.nativeCol);
        this.map["xdr:colOff"].render(xmlStream, model.nativeColOff);

        this.map["xdr:row"].render(xmlStream, model.nativeRow);
        this.map["xdr:rowOff"].render(xmlStream, model.nativeRowOff);

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside the corner is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.reset();
                break;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                break;
            }
        }
        return true;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag: {
                const { map } = this;
                // the children were reset when the element opened: a value whose element is absent is null
                this.model = {
                    nativeCol: cell(map["xdr:col"].model, colCache.MAX_COLS - 1),
                    nativeColOff: coordinate(map["xdr:colOff"].model),
                    nativeRow: cell(map["xdr:row"].model, colCache.MAX_ROWS - 1),
                    nativeRowOff: coordinate(map["xdr:rowOff"].model),
                };
                return false;
            }
            default:
                // not quite sure how we get here!
                return true;
        }
    }
}

export default CellPositionXform;
