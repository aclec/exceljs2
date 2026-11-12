const uuidv4 = require("../../../../utils/uuid");
const BaseXform = require("../../base-xform");
const CompositeXform = require("../../composite-xform");

const DxfXform = require("../../style/dxf-xform");
const CfRuleXform = require("../cf/cf-rule-xform");
const DatabarExtXform = require("./databar-ext-xform");
const IconSetExtXform = require("./icon-set-ext-xform");
const FExtXform = require("./f-ext-xform");

// the differential format of an x14 rule: same content as <dxf>
class DxfExtXform extends DxfXform {
    get tag() {
        return "x14:dxf";
    }
}

const extIcons = {
    "3Triangles": true,
    "3Stars": true,
    "5Boxes": true,
};

class CfRuleExtXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            "x14:dataBar": (this.databarXform = new DatabarExtXform()),
            "x14:iconSet": (this.iconSetXform = new IconSetExtXform()),
            // formula rules (expression, cellIs) carry their formulae and style inline
            "xm:f": (this.fExtXform = new FExtXform()),
            "x14:dxf": (this.dxfXform = new DxfExtXform()),
        };
    }

    get tag() {
        return "x14:cfRule";
    }

    static isExt(rule) {
        // is this rule primitive?
        if (CfRuleXform.isX14(rule)) {
            return true;
        }
        if (rule.type === "dataBar") {
            return DatabarExtXform.isExt(rule);
        }
        if (rule.type === "iconSet") {
            if (rule.custom || extIcons[rule.iconSet]) {
                return true;
            }
        }
        return false;
    }

    prepare(model) {
        if (CfRuleExtXform.isExt(model)) {
            model.x14Id = `{${uuidv4()}}`.toUpperCase();
        }
    }

    render(xmlStream, model) {
        if (!CfRuleExtXform.isExt(model)) {
            return;
        }

        switch (model.type) {
            case "dataBar":
                this.renderDataBar(xmlStream, model);
                break;
            case "iconSet":
                this.renderIconSet(xmlStream, model);
                break;
            case "expression":
            case "cellIs":
                this.renderFormula(xmlStream, model);
                break;
        }
    }

    // expression and cellIs rules (see CfRuleXform.isX14): CT_CfRule of [MS-XLSX] 2.6.2
    renderFormula(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: model.type,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            operator: model.type === "cellIs" ? model.operator : undefined,
            id: model.x14Id,
        });

        // like the main <cfRule>, an expression has a single formula
        const formulae = model.formulae || [];
        (model.type === "expression" ? formulae.slice(0, 1) : formulae).forEach((formula) => {
            this.fExtXform.render(xmlStream, formula);
        });
        if (model.style) {
            this.dxfXform.render(xmlStream, model.style);
        }

        xmlStream.closeNode();
    }

    renderDataBar(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "dataBar",
            id: model.x14Id,
        });

        this.databarXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    renderIconSet(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "iconSet",
            priority: model.priority,
            id: model.x14Id || `{${uuidv4()}}`,
        });

        this.iconSetXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    createNewModel({ attributes }) {
        return {
            type: attributes.type,
            x14Id: attributes.id,
            priority: BaseXform.toIntValue(attributes.priority),
            stopIfTrue: BaseXform.toBoolValue(attributes.stopIfTrue),
            operator: attributes.operator,
        };
    }

    onParserClose(name, parser) {
        switch (name) {
            case "xm:f":
                this.model.formulae = this.model.formulae || [];
                this.model.formulae.push(parser.model);
                break;
            case "x14:dxf":
                this.model.style = parser.model;
                break;
            default:
                Object.assign(this.model, parser.model);
                break;
        }
    }
}

module.exports = CfRuleExtXform;
