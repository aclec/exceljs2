import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** The box of a note in cells (fractional, 0-based): left / right columns, top / bottom rows. */
export interface VmlAnchorRect {
    left: number;
    top: number;
    right: number;
    bottom: number;
}

/** What {@link VmlAnchorXform} renders: the box of the note, or else the cell it is attached to. */
export interface VmlAnchorRenderModel {
    anchor?: VmlAnchorRect;
    /** The cell of the note (1-based row and column). */
    refAddress: { readonly row: number; readonly col: number };
}

/** The 8 numbers of `<x:Anchor>`: column, x offset, row, y offset of the top-left then of the bottom-right corner. */
type AnchorNumbers = [number, number, number, number, number, number, number, number];

/**
 * The position of the box of a note, `<x:Anchor>` (this is what draws the triangle in the commented cell). Parsing keeps
 * the text in {@link text}, not in the model.
 */
class VmlAnchorXform extends BaseXform<VmlAnchorRenderModel> {
    /** While parsing: the text of the element (`""` until a text is read). */
    declare text?: string;

    get tag(): string {
        return "x:Anchor";
    }

    /** @returns the anchor numbers of a box */
    getAnchorRect(anchor: VmlAnchorRect): AnchorNumbers {
        const l = Math.floor(anchor.left);
        const lf = Math.floor((anchor.left - l) * 68);
        const t = Math.floor(anchor.top);
        const tf = Math.floor((anchor.top - t) * 18);
        const r = Math.floor(anchor.right);
        const rf = Math.floor((anchor.right - r) * 68);
        const b = Math.floor(anchor.bottom);
        const bf = Math.floor((anchor.bottom - b) * 18);
        return [l, lf, t, tf, r, rf, b, bf];
    }

    /** @returns the anchor numbers of the default box of a note: right of its cell, over the rows around it */
    getDefaultRect(ref: VmlAnchorRenderModel["refAddress"]): AnchorNumbers {
        const l = ref.col;
        const lf = 6;
        const t = Math.max(ref.row - 2, 0);
        const tf = 14;
        const r = l + 2;
        const rf = 2;
        const b = t + 4;
        const bf = 16;
        return [l, lf, t, tf, r, rf, b, bf];
    }

    override render(xmlStream: XmlStream, model: VmlAnchorRenderModel): void {
        const rect = model.anchor ? this.getAnchorRect(model.anchor) : this.getDefaultRect(model.refAddress);

        xmlStream.leafNode("x:Anchor", undefined, rect.join(", "));
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.text = "";
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        this.text = text;
    }

    override parseClose(): boolean {
        return false;
    }
}

export default VmlAnchorXform;
