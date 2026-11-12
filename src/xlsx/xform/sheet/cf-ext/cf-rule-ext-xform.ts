import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type {
    CfFormula,
    ConditionalFormattingRuleModel,
    DataBarRuleModel,
    DxfModel,
    FormulaRuleModel,
    IconSetRuleModel,
} from "../../../../types";
import uuidv4 from "../../../../utils/uuid";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

import DxfXform from "../../style/dxf-xform";
import CfRuleXform, { getTextFormula, getTimePeriodFormula, opType } from "../cf/cf-rule-xform";
import DatabarExtXform from "./databar-ext-xform";
import IconSetExtXform from "./icon-set-ext-xform";
import FExtXform from "./f-ext-xform";

// the differential format of an x14 rule: same content as <dxf>
class DxfExtXform extends DxfXform {
    override get tag(): string {
        return "x14:dxf";
    }
}

// the icon sets of Excel 2010+, written in the x14 extension only
const EXT_ICONS: ReadonlySet<string | undefined> = new Set(["3Triangles", "3Stars", "5Boxes"]);

/** The attributes of the `<x14:cfRule>` of a formula rule that depend on its type. */
interface FormulaRuleAttributes {
    type: string;
    operator?: string;
    text?: string;
    timePeriod?: string;
}

const isFormula = (formula: CfFormula | undefined): formula is CfFormula => Boolean(formula);

/**
 * An `<x14:cfRule>` of the x14 extension of a worksheet: the x14 version of a data bar (see DatabarExtXform.isExt), the
 * custom and x14 icon sets, and the formula rules that Excel 2007 cannot read (see CfRuleXform.isX14), with their
 * formulae and format inline. Linked to a main rule by its `id` (`x14Id`). Parsing gives a rule with the parsed
 * attributes (text rules as `containsText`, like the main `<cfRule>`), merged with the model of its children.
 */
class CfRuleExtXform extends CompositeXform<ConditionalFormattingRuleModel> {
    /** The xform of a data bar. */
    readonly databarXform = new DatabarExtXform();
    /** The xform of an icon set. */
    readonly iconSetXform = new IconSetExtXform();
    /** The xform of a formula. */
    readonly fExtXform = new FExtXform();
    /** The xform of the format. */
    readonly dxfXform = new DxfExtXform();

    declare map: {
        readonly "x14:dataBar": DatabarExtXform;
        readonly "x14:iconSet": IconSetExtXform;
        readonly "xm:f": FExtXform;
        readonly "x14:dxf": DxfExtXform;
    };

    constructor() {
        super();

        this.map = {
            "x14:dataBar": this.databarXform,
            "x14:iconSet": this.iconSetXform,
            // formula rules (expression, cellIs) carry their formulae and style inline
            "xm:f": this.fExtXform,
            "x14:dxf": this.dxfXform,
        };
    }

    get tag(): string {
        return "x14:cfRule";
    }

    /** @returns whether the rule is written in the x14 extension */
    static isExt(this: void, rule: ConditionalFormattingRuleModel): boolean {
        if (CfRuleXform.isX14(rule)) {
            return true;
        }
        if (rule.type === "dataBar") {
            return DatabarExtXform.isExt(rule);
        }
        if (rule.type === "iconSet") {
            if (rule.custom || EXT_ICONS.has(rule.iconSet)) {
                return true;
            }
        }
        return false;
    }

    /** Gives an x14 rule a new `x14Id` (mutates the rule). */
    override prepare(model: ConditionalFormattingRuleModel): void {
        if (CfRuleExtXform.isExt(model)) {
            model.x14Id = `{${uuidv4()}}`.toUpperCase();
        }
    }

    /** Writes an x14 rule (nothing for the other rules). */
    override render(xmlStream: XmlStream, model: ConditionalFormattingRuleModel): void {
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
            case "containsText":
            case "timePeriod":
                this.renderFormula(xmlStream, model);
                break;
            case "top10":
            case "aboveAverage":
            case "colorScale":
            default:
                break;
        }
    }

    /** @returns the attributes of a formula rule that depend on its type */
    static formulaAttributes(this: void, model: FormulaRuleModel): FormulaRuleAttributes {
        switch (model.type) {
            case "cellIs":
                return { type: "cellIs", operator: model.operator };
            case "containsText": {
                const operator = CfRuleXform.textOperator(model);
                return { type: CfRuleXform.textType(model), operator, text: operator ? model.text : undefined };
            }
            case "timePeriod":
                return { type: "timePeriod", timePeriod: model.timePeriod };
            case "expression":
            default:
                return { type: model.type };
        }
    }

    /** @returns the formulae written for a formula rule (derived ones for a text or time period rule without formulae) */
    static formulae(this: void, model: FormulaRuleModel): CfFormula[] {
        const formulae = model.formulae || [];
        switch (model.type) {
            case "expression":
                // like the main <cfRule>, an expression has a single formula
                return formulae.slice(0, 1);
            case "containsText":
                // Excel writes the text compared with as a second formula when it is a reference
                return formulae.length ? formulae : [getTextFormula(model)].filter(isFormula);
            case "timePeriod":
                return formulae.length ? formulae : [getTimePeriodFormula(model)].filter(isFormula);
            case "cellIs":
            default:
                return formulae;
        }
    }

    /** Writes a formula rule (see CfRuleXform.isX14): CT_CfRule of [MS-XLSX] 2.6.2. */
    renderFormula(xmlStream: XmlStream, model: FormulaRuleModel): void {
        const { type, operator, text, timePeriod } = CfRuleExtXform.formulaAttributes(model);
        xmlStream.openNode(this.tag, {
            type,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            operator,
            text,
            timePeriod,
            id: model.x14Id,
        });

        for (const formula of CfRuleExtXform.formulae(model)) {
            this.fExtXform.render(xmlStream, formula);
        }
        if (model.style) {
            this.dxfXform.render(xmlStream, model.style);
        }

        xmlStream.closeNode();
    }

    /** Writes the x14 version of a data bar. */
    renderDataBar(xmlStream: XmlStream, model: DataBarRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "dataBar",
            id: model.x14Id,
        });

        this.databarXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    /** Writes an x14 icon set (a rule without `x14Id` gets a new id, not kept). */
    renderIconSet(xmlStream: XmlStream, model: IconSetRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "iconSet",
            priority: model.priority,
            id: model.x14Id || `{${uuidv4()}}`,
        });

        this.iconSetXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): ConditionalFormattingRuleModel {
        // the attributes of the file, unvalidated
        return {
            // text rules: containsText with the type as operator, like the main <cfRule>
            ...opType(attributes),
            x14Id: attributes.id,
            priority: BaseXform.toIntValue(attributes.priority),
            stopIfTrue: BaseXform.toBoolValue(attributes.stopIfTrue),
            text: attributes.text,
            timePeriod: attributes.timePeriod,
        } as ConditionalFormattingRuleModel;
    }

    override onParserClose(name: string): void {
        // the element was opened before its children (createNewModel)
        const model = this.model!;
        switch (name) {
            case "xm:f":
                // set by the start tag of <xm:f>
                model.formulae = model.formulae || [];
                model.formulae.push(this.fExtXform.model!);
                break;
            case "x14:dxf":
                // dxf-xform is not typed yet: its model is a DxfModel
                model.style = this.dxfXform.model as DxfModel;
                break;
            case "x14:dataBar":
                Object.assign(model, this.databarXform.model);
                break;
            case "x14:iconSet":
                Object.assign(model, this.iconSetXform.model);
                break;
        }
    }
}

export default CfRuleExtXform;
