import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ColorModel, OutlinePropertiesModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import ColorXform from "../style/color-xform";
import PageSetupPropertiesXform, { type PageSetupPropertiesModel } from "./page-setup-properties-xform";
import OutlinePropertiesXform from "./outline-properties-xform";

/** The model of `<sheetPr>`: parts of the worksheet properties and page setup. */
export interface SheetPropertiesModel {
    tabColor?: ColorModel;
    pageSetup?: PageSetupPropertiesModel;
    outlineProperties?: OutlinePropertiesModel;
}

/**
 * `<sheetPr>`: the tab color, outline properties and `fitToPage` of a worksheet. Rendered only when one of its children
 * is (the element is rolled back otherwise). Parsing gives `null` when none of the children has a model.
 */
class SheetPropertiesXform extends BaseXform<SheetPropertiesModel> {
    declare map: {
        readonly tabColor: ColorXform;
        readonly pageSetUpPr: PageSetupPropertiesXform;
        readonly outlinePr: OutlinePropertiesXform;
    };

    constructor() {
        super();

        this.map = {
            tabColor: new ColorXform("tabColor"),
            pageSetUpPr: new PageSetupPropertiesXform(),
            outlinePr: new OutlinePropertiesXform(),
        };
    }

    get tag(): string {
        return "sheetPr";
    }

    override render(xmlStream: XmlStream, model: SheetPropertiesModel | undefined): void {
        if (model) {
            xmlStream.addRollback();
            xmlStream.openNode("sheetPr");

            let inner = false;
            inner = this.map.tabColor.render(xmlStream, model.tabColor) || inner;
            // CT_SheetPr order: tabColor, outlinePr, pageSetUpPr
            inner = this.map.outlinePr.render(xmlStream, model.outlineProperties) || inner;
            inner = this.map.pageSetUpPr.render(xmlStream, model.pageSetup) || inner;

            if (inner) {
                xmlStream.closeNode();
                xmlStream.commit();
            } else {
                xmlStream.rollback();
            }
        }
    }

    /** @returns `true` for `<sheetPr>` and its known children */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (node.name === this.tag) {
            this.reset();
            return true;
        }
        const parser = ownEntry(this.map, node.name);
        if (parser) {
            this.parser = parser;
            parser.parseOpen(node);
            return true;
        }
        return false;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @returns `false` at the end of `<sheetPr>` (the model is then built), else `true` */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        // color-xform is not typed yet: its model is a ColorModel
        const tabColor = this.map.tabColor.model as ColorModel | null | undefined;
        const pageSetup = this.map.pageSetUpPr.model;
        const outlineProperties = this.map.outlinePr.model;
        if (tabColor || pageSetup || outlineProperties) {
            const model: SheetPropertiesModel = (this.model = {});
            if (tabColor) {
                model.tabColor = tabColor;
            }
            if (pageSetup) {
                model.pageSetup = pageSetup;
            }
            if (outlineProperties) {
                model.outlineProperties = outlineProperties;
            }
        } else {
            this.model = null;
        }
        return false;
    }
}

export default SheetPropertiesXform;
