import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DxfModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";

import AlignmentXform from "./alignment-xform";
import BorderXform from "./border-xform";
import FillXform from "./fill-xform";
import FontXform from "./font-xform";
import NumFmtXform from "./numfmt-xform";
import ProtectionXform from "./protection-xform";

/**
 * A differential format, `<dxf>` of the styles part (the style of a conditional formatting rule or a table column), see
 * {@link DxfModel}. Its number format is rendered only when the model has both `numFmt` and `numFmtId`. Unknown child
 * elements are skipped.
 */
class DxfXform extends BaseXform<DxfModel> {
    declare map: {
        readonly alignment: AlignmentXform;
        readonly border: BorderXform;
        readonly fill: FillXform;
        readonly font: FontXform;
        readonly numFmt: NumFmtXform;
        readonly protection: ProtectionXform;
    };

    constructor() {
        super();

        this.map = {
            alignment: new AlignmentXform(),
            border: new BorderXform(),
            fill: new FillXform(),
            font: new FontXform(),
            numFmt: new NumFmtXform(),
            protection: new ProtectionXform(),
        };
    }

    get tag(): string {
        return "dxf";
    }

    // CT_Dxf sequence: font, numFmt, fill, alignment, border, protection, extLst
    override render(xmlStream: XmlStream, model: DxfModel): void {
        xmlStream.openNode(this.tag);

        if (model.font) {
            this.map.font.render(xmlStream, model.font);
        }
        if (model.numFmt && model.numFmtId) {
            const numFmtModel = { id: model.numFmtId, formatCode: model.numFmt };
            this.map.numFmt.render(xmlStream, numFmtModel);
        }
        if (model.fill) {
            this.map.fill.render(xmlStream, model.fill);
        }
        if (model.alignment) {
            this.map.alignment.render(xmlStream, model.alignment);
        }
        if (model.border) {
            this.map.border.render(xmlStream, model.border);
        }
        if (model.protection) {
            this.map.protection.render(xmlStream, model.protection);
        }

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<dxf>` is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        switch (node.name) {
            case this.tag:
                // this node is often repeated. Need to reset children
                this.reset();
                return true;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                return true;
            }
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
                this.parser = undefined;
            }
            return true;
        }
        if (name === this.tag) {
            // like cell styles, a dxf style holds its number format as the format code string;
            // the numFmtId is reassigned by StylesXform.addDxfStyle when written
            const numFmt = this.map.numFmt.model;
            // (the children were reset when <dxf> opened: the parts that are absent are null)
            this.model = {
                alignment: this.map.alignment.model,
                border: this.map.border.model,
                fill: this.map.fill.model,
                font: this.map.font.model,
                numFmt: numFmt && numFmt.formatCode,
                protection: this.map.protection.model,
            };
            return false;
        }

        return true;
    }
}

export default DxfXform;
