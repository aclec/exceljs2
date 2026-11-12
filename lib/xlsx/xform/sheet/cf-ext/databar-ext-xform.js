const BaseXform = require("../../base-xform");
const CompositeXform = require("../../composite-xform");

const ColorXform = require("../../style/color-xform");
const CfvoExtXform = require("./cfvo-ext-xform");

class DatabarExtXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            "x14:cfvo": (this.cfvoXform = new CfvoExtXform()),
            "x14:borderColor": (this.borderColorXform = new ColorXform("x14:borderColor")),
            "x14:negativeBorderColor": (this.negativeBorderColorXform = new ColorXform("x14:negativeBorderColor")),
            "x14:negativeFillColor": (this.negativeFillColorXform = new ColorXform("x14:negativeFillColor")),
            "x14:axisColor": (this.axisColorXform = new ColorXform("x14:axisColor")),
        };
    }

    // Excel 2007 data bars are gradient bars with none of the x14 settings: any other data bar,
    // e.g. one read from a file with its extension, needs one
    static isExt(rule) {
        return (
            !rule.gradient ||
            DatabarExtXform.X14_SETTINGS.some((key) => rule[key] !== undefined) ||
            (rule.cfvo || []).some((cfvo) => cfvo.type === "autoMin" || cfvo.type === "autoMax")
        );
    }

    get tag() {
        return "x14:dataBar";
    }

    // the defaults of [MS-XLSX] CT_DataBar, except for the lengths of a new data bar: Excel's 0 and 100
    render(xmlStream, model) {
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
        cfvoList.forEach((cfvo) => {
            this.cfvoXform.render(xmlStream, cfvo);
        });

        // CT_DataBar order: cfvo, fillColor, borderColor, negativeFillColor, negativeBorderColor, axisColor
        this.borderColorXform.render(xmlStream, model.borderColor);
        this.negativeFillColorXform.render(xmlStream, model.negativeFillColor);
        this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor);
        this.axisColorXform.render(xmlStream, model.axisColor);

        xmlStream.closeNode();
    }

    createNewModel({ attributes }) {
        return {
            cfvo: [],
            minLength: BaseXform.toIntValue(attributes.minLength, 10),
            maxLength: BaseXform.toIntValue(attributes.maxLength, 90),
            border: BaseXform.toBoolValue(attributes.border, false),
            gradient: BaseXform.toBoolValue(attributes.gradient, true),
            negativeBarColorSameAsPositive: BaseXform.toBoolValue(attributes.negativeBarColorSameAsPositive, false),
            negativeBarBorderColorSameAsPositive: BaseXform.toBoolValue(attributes.negativeBarBorderColorSameAsPositive, true),
            axisPosition: BaseXform.toStringValue(attributes.axisPosition, "auto"),
            direction: BaseXform.toStringValue(attributes.direction, "context"),
        };
    }

    onParserClose(name, parser) {
        const [, prop] = name.split(":");
        switch (prop) {
            case "cfvo":
                this.model.cfvo.push(parser.model);
                break;

            default:
                this.model[prop] = parser.model;
                break;
        }
    }
}

// the settings of <x14:dataBar> that Excel 2007 ignores (its <dataBar> lengths are not read)
DatabarExtXform.X14_SETTINGS = [
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

// Excel's defaults for a data bar created without explicit thresholds
DatabarExtXform.DEFAULT_CFVO = [{ type: "autoMin" }, { type: "autoMax" }];

module.exports = DatabarExtXform;
