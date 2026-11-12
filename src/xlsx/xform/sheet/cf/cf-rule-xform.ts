/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

import Range from "../../../../doc/range";

import DatabarXform from "./databar-xform";
import DatabarExtXform from "../cf-ext/databar-ext-xform";
import ExtLstRefXform from "./ext-lst-ref-xform";
import FormulaXform from "./formula-xform";
import ColorScaleXform from "./color-scale-xform";
import IconSetXform from "./icon-set-xform";

const extIcons = {
    "3Triangles": true,
    "3Stars": true,
    "5Boxes": true,
};

// the operator of a text rule is the type of its <cfRule>; the four comparing a text also have an
// operator attribute (ST_ConditionalFormattingOperator)
const TEXT_OPERATORS = {
    containsText: "containsText",
    notContainsText: "notContains",
    beginsWith: "beginsWith",
    endsWith: "endsWith",
};

// a text as a string literal of a formula
const textLiteral = (text) => `"${String(text === undefined ? "" : text).replace(/"/g, '""')}"`;

const getTextFormula = (model) => {
    if (model.formulae && model.formulae[0]) {
        return model.formulae[0];
    }

    const range = new Range(model.ref);
    const { tl } = range;
    const text = textLiteral(model.text);
    switch (model.operator || "containsText") {
        case "containsText":
            return `NOT(ISERROR(SEARCH(${text},${tl})))`;
        case "notContainsText":
            return `ISERROR(SEARCH(${text},${tl}))`;
        case "beginsWith":
            return `LEFT(${tl},LEN(${text}))=${text}`;
        case "endsWith":
            return `RIGHT(${tl},LEN(${text}))=${text}`;
        case "containsBlanks":
            return `LEN(TRIM(${tl}))=0`;
        case "notContainsBlanks":
            return `LEN(TRIM(${tl}))>0`;
        case "containsErrors":
            return `ISERROR(${tl})`;
        case "notContainsErrors":
            return `NOT(ISERROR(${tl}))`;
        default:
            return undefined;
    }
};

const getTimePeriodFormula = (model) => {
    if (model.formulae && model.formulae[0]) {
        return model.formulae[0];
    }

    const range = new Range(model.ref);
    const { tl } = range;
    switch (model.timePeriod) {
        case "thisWeek":
            return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
        case "lastWeek":
            return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
        case "nextWeek":
            return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
        case "yesterday":
            return `FLOOR(${tl},1)=TODAY()-1`;
        case "today":
            return `FLOOR(${tl},1)=TODAY()`;
        case "tomorrow":
            return `FLOOR(${tl},1)=TODAY()+1`;
        case "last7Days":
            return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
        case "lastMonth":
            return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
        case "thisMonth":
            return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
        case "nextMonth":
            return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
        default:
            return undefined;
    }
};

// Does a formula reference a cell of another sheet ("Other!A1", "'My sheet'!A1", "Sheet1:Sheet3!A1",
// "[1]Data!A1")? In a formula, "!" only appears in such references, in string literals and at the
// end of error literals (#REF!, #DIV/0!...).
const referencesOtherSheet = (formula) => {
    if (typeof formula !== "string") {
        return false;
    }
    let inString = false;
    let inSheetName = false;
    let inError = false;
    for (const char of formula) {
        if (inError && !/[A-Za-z0-9/_]/.test(char)) {
            inError = false;
            if (char === "!") {
                continue;
            }
        }
        if (char === '"' && !inSheetName) {
            inString = !inString;
        } else if (char === "'" && !inString) {
            inSheetName = !inSheetName;
        } else if (char === "#" && !inString && !inSheetName) {
            inError = true;
        } else if (char === "!" && !inString && !inSheetName) {
            return true;
        }
    }
    return false;
};

const opType = (attributes) => {
    const { type, operator } = attributes;
    switch (type) {
        case "containsText":
        case "notContainsText":
        case "beginsWith":
        case "endsWith":
        case "containsBlanks":
        case "notContainsBlanks":
        case "containsErrors":
        case "notContainsErrors":
            return {
                type: "containsText",
                operator: type,
            };

        default:
            return { type, operator };
    }
};

const FORMULA_TYPES = ["expression", "cellIs", "containsText", "timePeriod"];

class CfRuleXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            dataBar: (this.databarXform = new DatabarXform()),
            extLst: (this.extLstRefXform = new ExtLstRefXform()),
            formula: (this.formulaXform = new FormulaXform()),
            colorScale: (this.colorScaleXform = new ColorScaleXform()),
            iconSet: (this.iconSetXform = new IconSetXform()),
        };
    }

    get tag() {
        return "cfRule";
    }

    // the rules whose formulae and format can be written inline in the x14 extension
    static isFormulaRule(rule) {
        return FORMULA_TYPES.includes(rule.type);
    }

    // Excel 2007 conditional formats cannot reference other sheets: Excel 2010+ writes formula rules
    // that do only in the x14 extension of the worksheet, with their formulae and format inline.
    // Formula rules read from that extension alone are flagged x14 and written back there too.
    static isX14(rule) {
        if (!CfRuleXform.isFormulaRule(rule)) {
            return false;
        }
        return rule.x14 === true || (rule.formulae || []).some(referencesOtherSheet);
    }

    // the type of the <cfRule> of a text rule is its operator
    static textType(model) {
        return model.operator || "containsText";
    }

    // the operator attribute of a text rule, only for those comparing a text
    static textOperator(model) {
        return TEXT_OPERATORS[CfRuleXform.textType(model)];
    }

    static isPrimitive(rule) {
        // is this rule primitive?
        if (CfRuleXform.isX14(rule)) {
            return false;
        }
        if (rule.type === "iconSet") {
            if (rule.custom || extIcons[rule.iconSet]) {
                return false;
            }
        }
        return true;
    }

    render(xmlStream, model) {
        switch (model.type) {
            case "expression":
                this.renderExpression(xmlStream, model);
                break;
            case "cellIs":
                this.renderCellIs(xmlStream, model);
                break;
            case "top10":
                this.renderTop10(xmlStream, model);
                break;
            case "aboveAverage":
                this.renderAboveAverage(xmlStream, model);
                break;
            case "dataBar":
                this.renderDataBar(xmlStream, model);
                break;
            case "colorScale":
                this.renderColorScale(xmlStream, model);
                break;
            case "iconSet":
                this.renderIconSet(xmlStream, model);
                break;
            case "containsText":
                this.renderText(xmlStream, model);
                break;
            case "timePeriod":
                this.renderTimePeriod(xmlStream, model);
                break;
        }
    }

    renderExpression(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "expression",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
        });

        // formula is optional in CT_CfRule
        if (model.formulae && model.formulae[0]) {
            this.formulaXform.render(xmlStream, model.formulae[0]);
        }

        xmlStream.closeNode();
    }

    renderCellIs(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "cellIs",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            operator: model.operator,
        });

        (model.formulae || []).forEach((formula) => {
            this.formulaXform.render(xmlStream, formula);
        });

        xmlStream.closeNode();
    }

    renderTop10(xmlStream, model) {
        xmlStream.leafNode(this.tag, {
            type: "top10",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            percent: BaseXform.toBoolAttribute(model.percent, false),
            bottom: BaseXform.toBoolAttribute(model.bottom, false),
            rank: BaseXform.toIntValue(model.rank, 10, true),
        });
    }

    renderAboveAverage(xmlStream, model) {
        xmlStream.leafNode(this.tag, {
            type: "aboveAverage",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            aboveAverage: BaseXform.toBoolAttribute(model.aboveAverage, true),
        });
    }

    renderDataBar(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "dataBar",
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
        });

        this.databarXform.render(xmlStream, model);
        // reference to the x14 counterpart of this rule, when it has one
        if (DatabarExtXform.isExt(model)) {
            this.extLstRefXform.render(xmlStream, model);
        }

        xmlStream.closeNode();
    }

    renderColorScale(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "colorScale",
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
        });

        this.colorScaleXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    renderIconSet(xmlStream, model) {
        // iconset is all primitive or all extLst
        if (!CfRuleXform.isPrimitive(model)) {
            return;
        }

        xmlStream.openNode(this.tag, {
            type: "iconSet",
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
        });

        this.iconSetXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    renderText(xmlStream, model) {
        const operator = CfRuleXform.textOperator(model);
        xmlStream.openNode(this.tag, {
            type: CfRuleXform.textType(model),
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            operator,
            text: operator ? model.text : undefined,
        });

        const formula = getTextFormula(model);
        if (formula) {
            this.formulaXform.render(xmlStream, formula);
        }

        xmlStream.closeNode();
    }

    renderTimePeriod(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: "timePeriod",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: BaseXform.toBoolAttribute(model.stopIfTrue, false),
            timePeriod: model.timePeriod,
        });

        const formula = getTimePeriodFormula(model);
        if (formula) {
            this.formulaXform.render(xmlStream, formula);
        }

        xmlStream.closeNode();
    }

    createNewModel({ attributes }) {
        return {
            ...opType(attributes),
            dxfId: BaseXform.toIntValue(attributes.dxfId),
            priority: BaseXform.toIntValue(attributes.priority),
            stopIfTrue: BaseXform.toBoolValue(attributes.stopIfTrue),
            text: attributes.text,
            timePeriod: attributes.timePeriod,
            percent: BaseXform.toBoolValue(attributes.percent),
            bottom: BaseXform.toBoolValue(attributes.bottom),
            rank: BaseXform.toIntValue(attributes.rank),
            aboveAverage: BaseXform.toBoolValue(attributes.aboveAverage),
        };
    }

    onParserClose(name, parser) {
        switch (name) {
            case "dataBar":
            case "extLst":
            case "colorScale":
            case "iconSet":
                // merge parser model with ours
                Object.assign(this.model, parser.model);
                break;

            case "formula":
                // except - formula is a string and appends to formulae
                this.model.formulae = this.model.formulae || [];
                this.model.formulae.push(parser.model);
                break;
        }
    }
}

// shared with the x14 rules
CfRuleXform.opType = opType;
CfRuleXform.getTextFormula = getTextFormula;
CfRuleXform.getTimePeriodFormula = getTimePeriodFormula;

export default CfRuleXform;
