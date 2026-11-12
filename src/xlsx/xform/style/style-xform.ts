import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XfModel } from "../../../types";
import BaseXform from "../base-xform";

import AlignmentXform from "./alignment-xform";
import ProtectionXform from "./protection-xform";

/** The options of a {@link StyleXform}. */
export interface StyleXformOptions {
    /** Write and read the `xfId` attribute (the `<xf>` of `cellXfs`, which refer to a cell style). */
    readonly xfId?: boolean;
}

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
//   Optional <protection>
// </xf>

/**
 * A cell format, `<xf>` of `cellXfs` / `cellStyleXfs` (see {@link XfModel}): the ids of its number format, font, fill
 * and border, with an `apply...` flag for each part that is set, and its alignment and protection.
 */
class StyleXform extends BaseXform<XfModel> {
    /** Whether `xfId` is written and read. */
    readonly xfId: boolean;

    declare map: { readonly alignment: AlignmentXform; readonly protection: ProtectionXform };

    constructor(options?: StyleXformOptions) {
        super();

        this.xfId = !!(options && options.xfId);
        this.map = {
            alignment: new AlignmentXform(),
            protection: new ProtectionXform(),
        };
    }

    get tag(): string {
        return "xf";
    }

    override render(xmlStream: XmlStream, model: XfModel): void {
        xmlStream.openNode("xf", {
            numFmtId: model.numFmtId || 0,
            fontId: model.fontId || 0,
            fillId: model.fillId || 0,
            borderId: model.borderId || 0,
        });
        if (this.xfId) {
            xmlStream.addAttribute("xfId", model.xfId || 0);
        }

        if (model.numFmtId) {
            xmlStream.addAttribute("applyNumberFormat", "1");
        }
        if (model.fontId) {
            xmlStream.addAttribute("applyFont", "1");
        }
        if (model.fillId) {
            xmlStream.addAttribute("applyFill", "1");
        }
        if (model.borderId) {
            xmlStream.addAttribute("applyBorder", "1");
        }
        if (model.alignment) {
            xmlStream.addAttribute("applyAlignment", "1");
        }
        if (model.protection) {
            xmlStream.addAttribute("applyProtection", "1");
        }

        // Rendering tags causes close of XML stream: the attributes must be added before rendering the children
        if (model.alignment) {
            this.map.alignment.render(xmlStream, model.alignment);
        }
        if (model.protection) {
            this.map.protection.render(xmlStream, model.protection);
        }

        xmlStream.closeNode();
    }

    /** @returns `true` for `<xf>` and its `<alignment>` / `<protection>` */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        // used during sax parsing of xml to build font object
        switch (node.name) {
            case "xf": {
                const { attributes } = node;
                const model: XfModel = {
                    numFmtId: BaseXform.toIntValue(attributes.numFmtId, NaN),
                    fontId: BaseXform.toIntValue(attributes.fontId, NaN),
                    fillId: BaseXform.toIntValue(attributes.fillId, NaN),
                    borderId: BaseXform.toIntValue(attributes.borderId, NaN),
                };
                if (this.xfId) {
                    model.xfId = BaseXform.toIntValue(attributes.xfId, NaN);
                }
                this.model = model;
                return true;
            }
            case "alignment":
                this.parser = this.map.alignment;
                this.parser.parseOpen(node);
                return true;
            case "protection":
                this.parser = this.map.protection;
                this.parser.parseOpen(node);
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the children are inside <xf>, which created the model
                if (this.map.protection === this.parser) {
                    this.model!.protection = this.map.protection.model;
                } else {
                    this.model!.alignment = this.map.alignment.model;
                }
                this.parser = undefined;
            }
            return true;
        }
        return name !== "xf";
    }
}

export default StyleXform;
