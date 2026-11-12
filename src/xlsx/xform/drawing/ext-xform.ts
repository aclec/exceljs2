import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ImageRangeModel } from "../../../types";
import BaseXform from "../base-xform";

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
const EMU_PER_PIXEL_AT_96_DPI = 9525;

/** The size of an image, in pixels (at 96 dpi). */
export type ExtModel = NonNullable<ImageRangeModel["ext"]>;

/** The options of an {@link ExtXform}. */
export interface ExtXformOptions {
    /** The element, e.g. `xdr:ext`. */
    readonly tag: string;
}

/** The size of a one-cell anchor, `<xdr:ext cx cy>`: in EMU in the file, in pixels (96 dpi) in the model. */
class ExtXform extends BaseXform<ExtModel> {
    /** The element. */
    readonly tag: string;

    declare map: Record<string, never>;

    constructor(options: ExtXformOptions) {
        super();

        this.tag = options.tag;
        this.map = {};
    }

    override render(xmlStream: XmlStream, model: ExtModel): void {
        xmlStream.openNode(this.tag);

        const width = Math.floor(model.width * EMU_PER_PIXEL_AT_96_DPI);
        const height = Math.floor(model.height * EMU_PER_PIXEL_AT_96_DPI);

        xmlStream.addAttribute("cx", width);
        xmlStream.addAttribute("cy", height);

        xmlStream.closeNode();
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            this.model = {
                width: parseInt(node.attributes.cx || "0", 10) / EMU_PER_PIXEL_AT_96_DPI,
                height: parseInt(node.attributes.cy || "0", 10) / EMU_PER_PIXEL_AT_96_DPI,
            };
            return true;
        }
        return false;
    }

    override parseText(/* text */): void {}

    override parseClose(/* name */): boolean {
        return false;
    }
}

export default ExtXform;
