import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, ColorModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import ColorXform from "../../style/color-xform";
import CfvoXform from "./cfvo-xform";

/** The model of `<colorScale>`: the thresholds and their colors (`undefined` for a `<color>` without color attribute). */
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
     * Writes the thresholds, then one `<color>` per color (`<color/>` for an `undefined` one: the colors match the
     * thresholds by position).
     *
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
            if (!this.colorXform.render(xmlStream, item)) {
                xmlStream.leafNode(this.colorXform.tag);
            }
        }

        xmlStream.closeNode();
    }

    override createNewModel(): ColorScaleModel {
        return {
            cfvo: [],
            color: [],
        };
    }

    override onParserClose(name: string, model: ColorScaleModel): void {
        // the element was opened before its children (createNewModel)
        switch (name) {
            case "cfvo": {
                // set by the start tag of <cfvo>
                const cfvo = this.cfvoXform.model;
                if (cfvo) {
                    model.cfvo.push(cfvo);
                }
                break;
            }
            case "color":
                // undefined for a <color> without color attribute
                model.color.push(this.colorXform.model ?? undefined);
                break;
            default:
                break;
        }
    }
}

export default ColorScaleXform;
