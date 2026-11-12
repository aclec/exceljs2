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

    static isExt(rule) {
        // not all databars need ext
        // TODO: refine this
        return !rule.gradient;
    }

    get tag() {
        return "x14:dataBar";
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            minLength: BaseXform.toIntAttribute(model.minLength, 0, true),
            maxLength: BaseXform.toIntAttribute(model.maxLength, 100, true),
            border: BaseXform.toBoolAttribute(model.border, false),
            gradient: BaseXform.toBoolAttribute(model.gradient, true),
            negativeBarColorSameAsPositive: BaseXform.toBoolAttribute(model.negativeBarColorSameAsPositive, true),
            negativeBarBorderColorSameAsPositive: BaseXform.toBoolAttribute(model.negativeBarBorderColorSameAsPositive, true),
            axisPosition: BaseXform.toAttribute(model.axisPosition, "auto"),
            direction: BaseXform.toAttribute(model.direction, "leftToRight"),
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
            minLength: BaseXform.toIntValue(attributes.minLength, 0),
            maxLength: BaseXform.toIntValue(attributes.maxLength, 100),
            border: BaseXform.toBoolValue(attributes.border, false),
            gradient: BaseXform.toBoolValue(attributes.gradient, true),
            negativeBarColorSameAsPositive: BaseXform.toBoolValue(attributes.negativeBarColorSameAsPositive, true),
            negativeBarBorderColorSameAsPositive: BaseXform.toBoolValue(attributes.negativeBarBorderColorSameAsPositive, true),
            axisPosition: BaseXform.toStringValue(attributes.axisPosition, "auto"),
            direction: BaseXform.toStringValue(attributes.direction, "leftToRight"),
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

// Excel's defaults for a data bar created without explicit thresholds
DatabarExtXform.DEFAULT_CFVO = [{ type: "autoMin" }, { type: "autoMax" }];

module.exports = DatabarExtXform;
