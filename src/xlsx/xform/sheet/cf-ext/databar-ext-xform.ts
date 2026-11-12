import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, ColorModel, DataBarRuleModel } from "../../../../types";
import { parseBoolean } from "../../../../utils/utils";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";
import { isLiteralOf } from "../attribute-values";

import ColorXform from "../../style/color-xform";
import CfvoExtXform from "./cfvo-ext-xform";

/** The settings of an x14 data bar that are colors (the child elements other than the thresholds). */
type DatabarExtColor = "borderColor" | "negativeFillColor" | "negativeBorderColor" | "axisColor";

/** [MS-XLSX] ST_DataBarAxisPosition. */
const isAxisPosition = isLiteralOf<NonNullable<DataBarRuleModel["axisPosition"]>>({ auto: true, middle: true, none: true });
/** [MS-XLSX] ST_DataBarDirection. */
const isDirection = isLiteralOf<NonNullable<DataBarRuleModel["direction"]>>({ context: true, leftToRight: true, rightToLeft: true });

/**
 * The model of `<x14:dataBar>` (merged into the rule): the parsed attributes, with the defaults of [MS-XLSX] CT_DataBar
 * (except `minLength` / `maxLength`: 10 / 90; also for an unknown `axisPosition` or `direction`), the thresholds and the
 * colors (`undefined` for an `auto` color).
 */
export type DatabarExtModel = Required<
    Pick<
        DataBarRuleModel,
        | "minLength"
        | "maxLength"
        | "border"
        | "gradient"
        | "negativeBarColorSameAsPositive"
        | "negativeBarBorderColorSameAsPositive"
        | "axisPosition"
        | "direction"
    >
> & { cfvo: CfvoModel[] } & Partial<Record<DatabarExtColor, ColorModel | undefined>>;

/**
 * `<x14:dataBar>` of an x14 `dataBar` rule: the settings that Excel 2007 ignores, the thresholds (which may be autoMin /
 * autoMax) and the colors other than the color of the bars.
 */
class DatabarExtXform extends CompositeXform<DatabarExtModel> {
    /** The settings of `<x14:dataBar>` that Excel 2007 ignores (its `<dataBar>` lengths are not read). */
    static readonly X14_SETTINGS: readonly (keyof DataBarRuleModel)[] = [
        "minLength",
        "maxLength",
        "border",
        "negativeBarColorSameAsPositive",
        "negativeBarBorderColorSameAsPositive",
        "axisPosition",
        "direction",
        "borderColor",
        "negativeFillColor",
        "negativeBorderColor",
        "axisColor",
    ];

    /** Excel's thresholds for a data bar created without explicit ones. */
    static readonly DEFAULT_CFVO: readonly CfvoModel[] = [{ type: "autoMin" }, { type: "autoMax" }];

    /** The xform of the thresholds. */
    readonly cfvoXform = new CfvoExtXform();
    /** The xforms of the colors. */
    readonly borderColorXform = new ColorXform("x14:borderColor");
    /** The xform of the border colour of the negative bars. */
    readonly negativeBorderColorXform = new ColorXform("x14:negativeBorderColor");
    /** The xform of the fill colour of the negative bars. */
    readonly negativeFillColorXform = new ColorXform("x14:negativeFillColor");
    /** The xform of the axis colour. */
    readonly axisColorXform = new ColorXform("x14:axisColor");

    declare map: {
        readonly "x14:cfvo": CfvoExtXform;
        readonly "x14:borderColor": ColorXform;
        readonly "x14:negativeBorderColor": ColorXform;
        readonly "x14:negativeFillColor": ColorXform;
        readonly "x14:axisColor": ColorXform;
    };

    constructor() {
        super();

        this.map = {
            "x14:cfvo": this.cfvoXform,
            "x14:borderColor": this.borderColorXform,
            "x14:negativeBorderColor": this.negativeBorderColorXform,
            "x14:negativeFillColor": this.negativeFillColorXform,
            "x14:axisColor": this.axisColorXform,
        };
    }

    /**
     * Excel 2007 data bars are gradient bars with none of the x14 settings: any other data bar, e.g. one read from a file
     * with its extension, needs one.
     *
     * @returns whether the data bar is written in the extension too
     */
    static isExt(this: void, rule: DataBarRuleModel): boolean {
        return (
            !rule.gradient ||
            DatabarExtXform.X14_SETTINGS.some((key) => rule[key] !== undefined) ||
            (rule.cfvo || []).some((cfvo) => cfvo.type === "autoMin" || cfvo.type === "autoMax")
        );
    }

    get tag(): string {
        return "x14:dataBar";
    }

    /**
     * Writes the x14 settings: the defaults of [MS-XLSX] CT_DataBar, except for the lengths of a new data bar (Excel's 0
     * and 100), and the default thresholds autoMin / autoMax.
     *
     * @param model - the rule
     */
    override render(xmlStream: XmlStream, model: Omit<DataBarRuleModel, "type">): void {
        xmlStream.openNode(this.tag, {
            minLength: BaseXform.toIntAttribute(model.minLength, 0, true),
            maxLength: BaseXform.toIntAttribute(model.maxLength, 100, true),
            border: BaseXform.toBoolAttribute(model.border, false),
            gradient: BaseXform.toBoolAttribute(model.gradient, true),
            negativeBarColorSameAsPositive: BaseXform.toBoolAttribute(model.negativeBarColorSameAsPositive, false),
            negativeBarBorderColorSameAsPositive: BaseXform.toBoolAttribute(model.negativeBarBorderColorSameAsPositive, true),
            axisPosition: BaseXform.toAttribute(model.axisPosition, "auto"),
            direction: BaseXform.toAttribute(model.direction, "context"),
        });

        const cfvoList = model.cfvo && model.cfvo.length ? model.cfvo : DatabarExtXform.DEFAULT_CFVO;
        for (const cfvo of cfvoList) {
            this.cfvoXform.render(xmlStream, cfvo);
        }

        // CT_DataBar order: cfvo, fillColor, borderColor, negativeFillColor, negativeBorderColor, axisColor
        this.borderColorXform.render(xmlStream, model.borderColor);
        this.negativeFillColorXform.render(xmlStream, model.negativeFillColor);
        this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor);
        this.axisColorXform.render(xmlStream, model.axisColor);

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): DatabarExtModel {
        return {
            cfvo: [],
            minLength: BaseXform.toIntValue(attributes.minLength, 10),
            maxLength: BaseXform.toIntValue(attributes.maxLength, 90),
            border: parseBoolean(attributes.border, false),
            gradient: parseBoolean(attributes.gradient, true),
            negativeBarColorSameAsPositive: parseBoolean(attributes.negativeBarColorSameAsPositive, false),
            negativeBarBorderColorSameAsPositive: parseBoolean(attributes.negativeBarBorderColorSameAsPositive, true),
            axisPosition: isAxisPosition(attributes.axisPosition) ? attributes.axisPosition : "auto",
            direction: isDirection(attributes.direction) ? attributes.direction : "context",
        };
    }

    override onParserClose(name: string): void {
        // the element was opened before its children (createNewModel)
        const { model } = this;
        if (!model) {
            return;
        }
        switch (name) {
            case "x14:cfvo": {
                // set by the start tag of <x14:cfvo>
                const cfvo = this.cfvoXform.model;
                if (cfvo) {
                    model.cfvo.push(cfvo);
                }
                break;
            }

            // undefined for a color element without color attribute
            case "x14:borderColor":
                model.borderColor = this.borderColorXform.model ?? undefined;
                break;
            case "x14:negativeBorderColor":
                model.negativeBorderColor = this.negativeBorderColorXform.model ?? undefined;
                break;
            case "x14:negativeFillColor":
                model.negativeFillColor = this.negativeFillColorXform.model ?? undefined;
                break;
            case "x14:axisColor":
                model.axisColor = this.axisColorXform.model ?? undefined;
                break;
        }
    }
}

export default DatabarExtXform;
