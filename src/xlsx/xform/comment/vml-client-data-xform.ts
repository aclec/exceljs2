import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CommentRefModel, VmlNoteShapeModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";

import VmlAnchorXform, { type VmlAnchorRect } from "./vml-anchor-xform";
import VmlProtectionXform from "./style/vml-protection-xform";
import VmlPositionXform from "./style/vml-position-xform";

// the editAs of a note, by the number of <x:MoveWithCells/> and <x:SizeWithCells/> elements present
const POSITION_TYPE = ["twoCells", "oneCells", "absolute"] as const;

/** What {@link VmlClientDataXform} renders: a note of the comments part (see `CommentRefModel`). */
export type VmlClientDataRenderModel = Pick<CommentRefModel, "note" | "refAddress"> & { anchor?: VmlAnchorRect };

/** The parsed model of a {@link VmlClientDataXform}. */
export type VmlClientDataModel = Pick<VmlNoteShapeModel, "anchor" | "editAs"> & {
    protection: { locked?: string; lockText?: string };
};

/**
 * The Excel data of a VML note, `<x:ClientData ObjectType="Note">`: how the box moves with the cells, its anchor, its
 * protection and its cell (0-based `<x:Row>` / `<x:Column>`).
 */
class VmlClientDataXform extends BaseXform<VmlClientDataModel | VmlClientDataRenderModel> {
    /** The parsed data (render takes a note: see the type parameter). */
    declare model: VmlClientDataModel | null | undefined;

    declare map: {
        readonly "x:Anchor": VmlAnchorXform;
        readonly "x:Locked": VmlProtectionXform;
        readonly "x:LockText": VmlProtectionXform;
        readonly "x:SizeWithCells": VmlPositionXform;
        readonly "x:MoveWithCells": VmlPositionXform;
    };

    constructor() {
        super();
        this.map = {
            "x:Anchor": new VmlAnchorXform(),
            "x:Locked": new VmlProtectionXform({ tag: "x:Locked" }),
            "x:LockText": new VmlProtectionXform({ tag: "x:LockText" }),
            "x:SizeWithCells": new VmlPositionXform({ tag: "x:SizeWithCells" }),
            "x:MoveWithCells": new VmlPositionXform({ tag: "x:MoveWithCells" }),
        };
    }

    get tag(): string {
        return "x:ClientData";
    }

    override render(xmlStream: XmlStream, model: VmlClientDataRenderModel): void {
        const { protection, editAs } = model.note;
        xmlStream.openNode(this.tag, { ObjectType: "Note" });
        this.map["x:MoveWithCells"].render(xmlStream, editAs, POSITION_TYPE);
        this.map["x:SizeWithCells"].render(xmlStream, editAs, POSITION_TYPE);
        this.map["x:Anchor"].render(xmlStream, model);
        // a note model always has its protection (the defaults of Note)
        this.map["x:Locked"].render(xmlStream, protection?.locked);
        xmlStream.leafNode("x:AutoFill", undefined, "False");
        this.map["x:LockText"].render(xmlStream, protection?.lockText);
        xmlStream.leafNode("x:Row", undefined, model.refAddress.row - 1);
        xmlStream.leafNode("x:Column", undefined, model.refAddress.col - 1);
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<x:ClientData>` is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                // the model is built when the element closes
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
            case this.tag:
                this.normalizeModel();
                return false;
            default:
                return true;
        }
    }

    /** Builds the model from the children, at the end of the element. */
    normalizeModel(): void {
        const position = Object.assign({}, this.map["x:MoveWithCells"].model, this.map["x:SizeWithCells"].model);
        const len = Object.keys(position).length;
        this.model = {
            anchor: this.map["x:Anchor"].text,
            protection: {
                locked: this.map["x:Locked"].text,
                lockText: this.map["x:LockText"].text,
            },
            editAs: POSITION_TYPE[len],
        };
    }
}

export default VmlClientDataXform;
