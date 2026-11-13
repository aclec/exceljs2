import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { NoteModel } from "../../../types";
import BaseXform from "../base-xform";

/** The parsed model of a {@link VmlTextboxXform}: the margins of the text, in cm. */
export interface VmlTextboxModel {
    inset: number[] | undefined;
}

/** What {@link VmlTextboxXform} renders: a note, whose `margins.inset` are written. */
export type VmlTextboxRenderModel = Pick<NoteModel, "note">;

/**
 * The text box of a VML note, `<v:textbox style inset><div/></v:textbox>`: its margins (`inset`), in mm in the file and
 * in cm in the model (`note.margins.inset`). Invalid margins are not written (the VML default applies) and parse as
 * `undefined`.
 */
class VmlTextboxXform extends BaseXform<VmlTextboxModel | VmlTextboxRenderModel> {
    /** The parsed margins (render takes a note: see the type parameter). */
    declare model: VmlTextboxModel | null | undefined;

    get tag(): string {
        return "v:textbox";
    }

    /** @returns `value * multiple` (rounded as `multiple.toFixed(2)`), followed by `unit` */
    conversionUnit(value: unknown, multiple: number, unit: string): string {
        // parseFloat converts its argument to a string, as String() does
        return `${parseFloat(String(value)) * Number(multiple.toFixed(2))}${unit}`;
    }

    /** @returns the margins (cm) of an `inset` attribute ("1mm,1mm,2mm,2mm"), `undefined` when absent or invalid */
    reverseConversionUnit(inset: string | undefined): number[] | undefined {
        if (!inset) {
            // no inset attribute: VML default margins (Excel usually omits it)
            return undefined;
        }
        const margins = inset.split(",").map((margin) => {
            return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, "")).toFixed(2));
        });
        return margins.every(Number.isFinite) ? margins : undefined;
    }

    override render(xmlStream: XmlStream, model: VmlTextboxRenderModel): void {
        const attributes: { style: string; inset?: string } = {
            style: "mso-direction-alt:auto",
        };
        if (model && model.note) {
            const { margins } = model.note;
            const inset: unknown = margins && margins.inset;
            if (Array.isArray(inset)) {
                // a NaN/Infinity margin would make the drawing unreadable: fall back to the VML default
                if (inset.every(Number.isFinite)) {
                    attributes.inset = inset.map((margin) => this.conversionUnit(margin, 10, "mm")).join(",");
                }
            } else if (inset && (typeof inset === "string" || typeof inset === "number")) {
                // not the documented array: a VML inset, written as it is
                attributes.inset = String(inset);
            }
        }
        xmlStream.openNode("v:textbox", attributes);
        xmlStream.leafNode("div", { style: "text-align:left" });
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<v:textbox>` is accepted (and skipped) */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = {
                    inset: this.reverseConversionUnit(node.attributes.inset),
                };
                return true;
            default:
                return true;
        }
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        switch (name) {
            case this.tag:
                return false;
            default:
                return true;
        }
    }
}

export default VmlTextboxXform;
