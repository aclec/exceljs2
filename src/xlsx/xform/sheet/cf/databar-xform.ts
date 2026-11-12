import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, ColorModel, DataBarRuleModel } from "../../../../types";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

import ColorXform from "../../style/color-xform";
import CfvoXform from "./cfvo-xform";

/** The model of `<dataBar>` (merged into the rule). */
export interface DatabarModel {
    cfvo: CfvoModel[];
    showValue: boolean | undefined;
    /** The color of the bars (`undefined` for an `auto` color). */
    color?: ColorModel | undefined;
}

/** The thresholds of a data bar as Excel 2007 reads them: the x14 types autoMin / autoMax are min / max. */
function mainCfvoType(type: string): CfvoModel["type"] | undefined {
    switch (type) {
        case "autoMin":
            return "min";
        case "autoMax":
            return "max";
        default:
            return undefined;
    }
}

/** `<dataBar>` of a `dataBar` rule (Excel 2007): its two thresholds and the color of the bars. */
class DatabarXform extends CompositeXform<DatabarModel> {
    /** The thresholds written when the rule has none (CT_DataBar requires two). */
    static readonly DEFAULT_CFVO: readonly CfvoModel[] = [{ type: "min" }, { type: "max" }];
    /** The color written when the rule has none (Excel's default). */
    static readonly DEFAULT_COLOR: ColorModel = { argb: "FF638EC6" };
    /** The main `<cfvo>` type of the x14 threshold types (written in the extension). */
    static readonly MAIN_CFVO_TYPE: Readonly<Partial<Record<CfvoModel["type"], CfvoModel["type"]>>> = { autoMin: "min", autoMax: "max" };

    /** The xform of the thresholds. */
    readonly cfvoXform = new CfvoXform();
    /** The xform of the color of the bars. */
    readonly colorXform = new ColorXform("color");

    declare map: { readonly cfvo: CfvoXform; readonly color: ColorXform };

    constructor() {
        super();

        this.map = {
            cfvo: this.cfvoXform,
            color: this.colorXform,
        };
    }

    get tag(): string {
        return "dataBar";
    }

    /** @param model - the rule */
    override render(xmlStream: XmlStream, model: Pick<DataBarRuleModel, "showValue" | "cfvo" | "color">): void {
        xmlStream.openNode(this.tag, { showValue: BaseXform.toBoolAttribute(model.showValue, true) });

        // CT_DataBar requires two cfvo and a color: fall back to Excel's defaults
        const cfvoList = model.cfvo && model.cfvo.length ? model.cfvo : DatabarXform.DEFAULT_CFVO;
        for (const cfvo of cfvoList) {
            // the x14 types autoMin / autoMax (written in the extension) are min / max for Excel 2007
            const type = mainCfvoType(cfvo.type);
            this.cfvoXform.render(xmlStream, type ? { ...cfvo, type } : cfvo);
        }
        this.colorXform.render(xmlStream, model.color || DatabarXform.DEFAULT_COLOR);

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): DatabarModel {
        return {
            cfvo: [],
            showValue: BaseXform.toBoolValue(attributes.showValue),
        };
    }

    override onParserClose(name: string): void {
        // the element was opened before its children (createNewModel)
        const model = this.model!;
        switch (name) {
            case "cfvo":
                // set by the start tag of <cfvo>
                model.cfvo.push(this.cfvoXform.model!);
                break;
            case "color":
                // color-xform is not typed yet: its model is a ColorModel (undefined for an auto color)
                model.color = this.colorXform.model as ColorModel | undefined;
                break;
            default:
                break;
        }
    }
}

export default DatabarXform;
