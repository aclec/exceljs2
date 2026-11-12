import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes } from "../../../utils/xml-stream";
import type { VmlNoteShapeModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import VmlTextboxXform from "./vml-textbox-xform";
import VmlClientDataXform, { type VmlClientDataRenderModel } from "./vml-client-data-xform";

/**
 * The VML shape of a note, `<v:shape>`: a hidden box with its text box and its Excel data (`<x:ClientData>`). Rendered
 * for a note of the comments part, parsed as a {@link VmlNoteShapeModel}.
 */
class VmlShapeXform extends BaseXform<VmlNoteShapeModel | VmlClientDataRenderModel> {
    /** The parsed shape (render takes a note: see the type parameter). */
    declare model: VmlNoteShapeModel | null | undefined;

    declare map: { readonly "v:textbox": VmlTextboxXform; readonly "x:ClientData": VmlClientDataXform };

    constructor() {
        super();
        this.map = {
            "v:textbox": new VmlTextboxXform(),
            "x:ClientData": new VmlClientDataXform(),
        };
    }

    get tag(): string {
        return "v:shape";
    }

    /** @param index - the index of the shape: its id is `_x0000_s{1025 + index}` (see VmlNotesXform.shapeIndex) */
    override render(xmlStream: XmlStream, model: VmlClientDataRenderModel, index: number): void {
        xmlStream.openNode("v:shape", VmlShapeXform.V_SHAPE_ATTRIBUTES(model, index));

        xmlStream.leafNode("v:fill", { color2: "infoBackground [80]" });
        xmlStream.leafNode("v:shadow", { color: "none [81]", obscured: "t" });
        xmlStream.leafNode("v:path", { "o:connecttype": "none" });
        this.map["v:textbox"].render(xmlStream, model);
        this.map["x:ClientData"].render(xmlStream, model);

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<v:shape>` is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        switch (node.name) {
            case this.tag:
                this.reset();
                this.model = {
                    margins: {
                        insetmode: node.attributes["o:insetmode"],
                    },
                    anchor: "",
                    editAs: "",
                    protection: {},
                };
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
                // the model was created when the element opened; the children were reset then (null when absent)
                const model = this.model!;
                const textbox = this.map["v:textbox"].model;
                const clientData = this.map["x:ClientData"].model;
                model.margins.inset = textbox && textbox.inset;
                model.protection = clientData && clientData.protection;
                model.anchor = clientData && clientData.anchor;
                model.editAs = clientData && clientData.editAs;
                return false;
            }
            default:
                return true;
        }
    }

    /** @returns the attributes of `<v:shape>` for a note and its index */
    static V_SHAPE_ATTRIBUTES(this: void, model: VmlClientDataRenderModel, index: number): XmlAttributes {
        return {
            id: `_x0000_s${1025 + index}`,
            type: "#_x0000_t202",
            style: "position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden",
            fillcolor: "infoBackground [80]",
            strokecolor: "none [81]",
            "o:insetmode": model.note.margins && model.note.margins.insetmode,
        };
    }
}

export default VmlShapeXform;
