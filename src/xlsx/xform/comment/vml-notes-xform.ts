import type { SaxInput, SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { VmlNotesModel } from "../../../types";

import BaseXform, { ownEntry } from "../base-xform";
import VmlShapeXform from "./vml-shape-xform";
import type { VmlClientDataRenderModel } from "./vml-client-data-xform";

/** What {@link VmlNotesXform} renders: the notes of a worksheet, and the first block of VML shape ids it claims. */
export interface VmlNotesRenderModel {
    readonly comments: readonly VmlClientDataRenderModel[];
    /** The first block of shape ids of the worksheet (default 1). */
    readonly shapeBlock?: number;
}

/**
 * The VML drawing of the notes of a worksheet, `xl/drawings/vmlDrawingN.vml`: one hidden shape per note (it is what
 * draws the triangle of a commented cell). Parsed as the list of the shapes ({@link VmlNotesModel}).
 */
class VmlNotesXform extends BaseXform<VmlNotesModel | VmlNotesRenderModel> {
    /** The parsed shapes (render takes the notes of a worksheet: see the type parameter). */
    declare model: VmlNotesModel | null | undefined;

    /** The namespaces of the root element. */
    static readonly DRAWING_ATTRIBUTES: XmlAttributes = {
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:x": "urn:schemas-microsoft-com:office:excel",
    };

    /**
     * VML shape ids must be unique across the workbook. They come in blocks of 1024 (block n holds ids n * 1024 to
     * n * 1024 + 1023) that the drawing claims in `<o:idmap data="n,m,...">`; as Excel does, the shapes of block n are
     * numbered from n * 1024 + 1, so a block holds 1023 notes. Each worksheet claims contiguous blocks, as many as its
     * notes need.
     */
    static readonly SHAPES_PER_BLOCK = 1023;

    /** @returns the number of blocks of a worksheet with `count` notes (at least 1) */
    static blockCount(this: void, count: number): number {
        return Math.max(1, Math.ceil(count / VmlNotesXform.SHAPES_PER_BLOCK));
    }

    /** @returns the idmap data of a worksheet whose `count` notes use the blocks from `firstBlock` ("1,2,3") */
    static idmap(this: void, firstBlock: number, count: number): string {
        return Array.from({ length: VmlNotesXform.blockCount(count) }, (_, i) => firstBlock + i).join(",");
    }

    /** @returns the index for VmlShapeXform (which numbers shapes from 1025) of the note at `index` */
    static shapeIndex(this: void, firstBlock: number, index: number): number {
        const block = firstBlock + Math.floor(index / VmlNotesXform.SHAPES_PER_BLOCK);
        return (block - 1) * 1024 + (index % VmlNotesXform.SHAPES_PER_BLOCK);
    }

    declare map: { readonly "v:shape": VmlShapeXform };

    constructor() {
        super();
        this.map = {
            "v:shape": new VmlShapeXform(),
        };
    }

    get tag(): string {
        return "xml";
    }

    /** @returns the parsed shapes */
    override async parseStream(stream: SaxInput): Promise<VmlNotesModel | null | undefined> {
        await super.parseStream(stream);
        return this.model;
    }

    override render(xmlStream: XmlStream, model: VmlNotesRenderModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, VmlNotesXform.DRAWING_ATTRIBUTES);

        // shape ids are unique across the workbook: model.shapeBlock is the first of the blocks of this worksheet
        const firstBlock = model.shapeBlock || 1;
        xmlStream.openNode("o:shapelayout", { "v:ext": "edit" });
        xmlStream.leafNode("o:idmap", { "v:ext": "edit", data: VmlNotesXform.idmap(firstBlock, model.comments.length) });
        xmlStream.closeNode();

        xmlStream.openNode("v:shapetype", {
            id: "_x0000_t202",
            coordsize: "21600,21600",
            "o:spt": 202,
            path: "m,l,21600r21600,l21600,xe",
        });
        xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
        xmlStream.leafNode("v:path", { gradientshapeok: "t", "o:connecttype": "rect" });
        xmlStream.closeNode();

        const shapeXform = this.map["v:shape"];
        model.comments.forEach((comment, index) => {
            shapeXform.render(xmlStream, comment, VmlNotesXform.shapeIndex(firstBlock, index));
        });

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the drawing is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.reset();
                this.model = {
                    comments: [],
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
                // a shape outside the root element (no model) is ignored; the shape has its model once its element closed
                const shape = this.map["v:shape"].model;
                if (this.model && shape) {
                    this.model.comments.push(shape);
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                return false;
            default:
                // could be some unrecognised tags
                return true;
        }
    }
}

export default VmlNotesXform;
