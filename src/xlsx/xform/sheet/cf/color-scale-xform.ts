import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, ColorModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import ColorXform from "../../style/color-xform";
import CfvoXform from "./cfvo-xform";

/** The model of `<colorScale>`: the thresholds and their colors (`undefined` for an `auto` color). */
export interface ColorScaleModel {
    cfvo: CfvoModel[];
    color: (ColorModel | undefined)[];
}

/** `<colorScale>` of a `colorScale` rule: its `<cfvo>` thresholds, then their `<color>`s (merged into the rule). */
class ColorScaleXform extends CompositeXform<ColorScaleModel> {
    /** The xform of the thresholds. */
    readonly cfvoXform = new CfvoXform();
    /** The xform of the colors. */
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
        return "colorScale";
    }

    /**
     * @param model - the rule
     * @throws Error when the rule has no `cfvo` or no `color` (CT_ColorScale requires them)
     */
    override render(
        xmlStream: XmlStream,
        model: { readonly cfvo?: readonly CfvoModel[]; readonly color?: readonly (ColorModel | undefined)[] }
    ): void {
        const { cfvo, color } = model;
        if (!cfvo || !color) {
            throw new Error("A colorScale conditional formatting rule needs cfvo and color");
        }
        xmlStream.openNode(this.tag);

        for (const item of cfvo) {
            this.cfvoXform.render(xmlStream, item);
        }
        for (const item of color) {
            this.colorXform.render(xmlStream, item);
        }

        xmlStream.closeNode();
    }

    override createNewModel(): ColorScaleModel {
        return {
            cfvo: [],
            color: [],
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
                model.color.push(this.colorXform.model as ColorModel | undefined);
                break;
            default:
                break;
        }
    }
}

export default ColorScaleXform;
