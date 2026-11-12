import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type {
    AboveAverageRuleModel,
    CellIsRuleModel,
    CfFormula,
    ColorScaleRuleModel,
    ConditionalFormattingRuleModel,
    ContainsTextRuleModel,
    DataBarRuleModel,
    ExpressionRuleModel,
    FormulaRuleModel,
    IconSetRuleModel,
    TimePeriodRuleModel,
    Top10RuleModel,
} from "../../../../types";
import type { CellIsOperators, TimePeriodTypes } from "../../../../../index";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";
import { isLiteralOf, parseOptionalBoolean } from "../attribute-values";

import Range from "../../../../doc/range";

import DatabarXform from "./databar-xform";
import DatabarExtXform from "../cf-ext/databar-ext-xform";
import ExtLstRefXform from "./ext-lst-ref-xform";
import FormulaXform from "./formula-xform";
import ColorScaleXform from "./color-scale-xform";
import IconSetXform from "./icon-set-xform";
import { isX14OnlyIconSet } from "./icon-sets";

// the operator of a text rule is the type of its <cfRule>; the four comparing a text also have an
// operator attribute (ST_ConditionalFormattingOperator)
function textOperatorAttribute(type: string): string | undefined {
    switch (type) {
        case "containsText":
            return "containsText";
        case "notContainsText":
            return "notContains";
        case "beginsWith":
            return "beginsWith";
        case "endsWith":
            return "endsWith";
        default:
            return undefined;
    }
}

// a text as a string literal of a formula
const textLiteral = (text: string | undefined): string => `"${String(text === undefined ? "" : text).replace(/"/g, '""')}"`;

/**
 * The range of a conditional formatting rule (its `sqref`).
 * @throws Error for a rule without range (a malformed file); it used to be a TypeError
 */
function rangeOfRule(ref: string | undefined): Range {
    if (ref === undefined) {
        throw new Error("Conditional formatting rule without range (sqref)");
    }
    return new Range(ref);
}

/**
 * The formula of a text rule: its first formula, else the one derived from its operator and text for the top left cell
 * of its range (`ref`, set by render).
 *
 * @returns `undefined` for an unknown operator
 */
function getTextFormula(this: void, model: ContainsTextRuleModel): CfFormula | undefined {
    if (model.formulae && model.formulae[0]) {
        return model.formulae[0];
    }

    const range = rangeOfRule(model.ref);
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
}

/**
 * The formula of a time period rule: its first formula, else the one derived from its time period for the top left cell
 * of its range (`ref`, set by render).
 *
 * @returns `undefined` for an unknown time period
 */
function getTimePeriodFormula(this: void, model: TimePeriodRuleModel): CfFormula | undefined {
    if (model.formulae && model.formulae[0]) {
        return model.formulae[0];
    }

    const range = rangeOfRule(model.ref);
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
        case undefined:
        default:
            return undefined;
    }
}

// Does a formula reference a cell of another sheet ("Other!A1", "'My sheet'!A1", "Sheet1:Sheet3!A1",
// "[1]Data!A1")? In a formula, "!" only appears in such references, in string literals and at the
// end of error literals (#REF!, #DIV/0!...).
const referencesOtherSheet = (formula: unknown): boolean => {
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

const isCellIsOperator = isLiteralOf<CellIsOperators>({
    equal: true,
    notEqual: true,
    greaterThan: true,
    greaterThanOrEqual: true,
    lessThan: true,
    lessThanOrEqual: true,
    between: true,
    notBetween: true,
});

const isTimePeriod = isLiteralOf<TimePeriodTypes>({
    lastWeek: true,
    thisWeek: true,
    nextWeek: true,
    yesterday: true,
    today: true,
    tomorrow: true,
    last7Days: true,
    lastMonth: true,
    thisMonth: true,
    nextMonth: true,
});

/**
 * The rule of a `<cfRule>` / `<x14:cfRule>` from the attributes common to both (`priority`, `stopIfTrue` and those of its
 * type), without its children: a rule of the variant of its `type`, a text rule being `containsText` with its type as
 * operator. An unknown `operator` (cellIs) or `timePeriod` is absent.
 *
 * @returns `null` for a type that the model does not have (`uniqueValues`, `duplicateValues`, an unknown one): such a
 *   rule is not read (it could not be written back)
 */
function ruleOfAttributes(this: void, attributes: Readonly<Record<string, string | undefined>>): ConditionalFormattingRuleModel | null {
    const { type } = attributes;
    const priority = BaseXform.toIntValue(attributes.priority);
    const stopIfTrue = parseOptionalBoolean(attributes.stopIfTrue);
    switch (type) {
        case "expression":
        case "colorScale":
        case "dataBar":
        case "iconSet":
            return { type, priority, stopIfTrue };
        case "cellIs": {
            const { operator } = attributes;
            return { type, priority, stopIfTrue, operator: isCellIsOperator(operator) ? operator : undefined };
        }
        case "top10":
            return {
                type,
                priority,
                stopIfTrue,
                percent: parseOptionalBoolean(attributes.percent),
                bottom: parseOptionalBoolean(attributes.bottom),
                rank: BaseXform.toIntValue(attributes.rank),
            };
        case "aboveAverage":
            return { type, priority, stopIfTrue, aboveAverage: parseOptionalBoolean(attributes.aboveAverage) };
        case "containsText":
        case "notContainsText":
        case "beginsWith":
        case "endsWith":
        case "containsBlanks":
        case "notContainsBlanks":
        case "containsErrors":
        case "notContainsErrors":
            return { type: "containsText", priority, stopIfTrue, operator: type, text: attributes.text };
        case "timePeriod": {
            const { timePeriod } = attributes;
            return { type, priority, stopIfTrue, timePeriod: isTimePeriod(timePeriod) ? timePeriod : undefined };
        }
        case undefined:
        default:
            return null;
    }
}

const FORMULA_TYPES: readonly string[] = ["expression", "cellIs", "containsText", "timePeriod"];

/** The `stopIfTrue` attribute (not written when false). */
function stopIfTrueAttribute(model: ConditionalFormattingRuleModel): "1" | "0" | boolean | undefined {
    return BaseXform.toBoolAttribute(model.stopIfTrue, false);
}

/**
 * A `<cfRule>` of a `<conditionalFormatting>` (Excel 2007 rules). The rules that Excel 2007 cannot read (see
 * {@link CfRuleXform.isPrimitive}) are written in the x14 extension instead, by cf-rule-ext-xform; a data bar is written
 * in both. Parsing gives a rule of the variant of its type (see {@link ruleOfAttributes}), merged with the model of its
 * children; the model is `null` for a rule of a type that the model does not have.
 */
class CfRuleXform extends CompositeXform<ConditionalFormattingRuleModel | null> {
    /** See {@link ruleOfAttributes} (shared with the x14 rules). */
    static readonly ruleOfAttributes = ruleOfAttributes;
    /** See {@link getTextFormula} (shared with the x14 rules). */
    static readonly getTextFormula = getTextFormula;
    /** See {@link getTimePeriodFormula} (shared with the x14 rules). */
    static readonly getTimePeriodFormula = getTimePeriodFormula;

    /** The xform of a data bar. */
    readonly databarXform = new DatabarXform();
    /** The xform of the reference to the x14 version of the rule. */
    readonly extLstRefXform = new ExtLstRefXform();
    /** The xform of a formula. */
    readonly formulaXform = new FormulaXform();
    /** The xform of a color scale. */
    readonly colorScaleXform = new ColorScaleXform();
    /** The xform of an icon set. */
    readonly iconSetXform = new IconSetXform();

    declare map: {
        readonly dataBar: DatabarXform;
        readonly extLst: ExtLstRefXform;
        readonly formula: FormulaXform;
        readonly colorScale: ColorScaleXform;
        readonly iconSet: IconSetXform;
    };

    constructor() {
        super();

        this.map = {
            dataBar: this.databarXform,
            extLst: this.extLstRefXform,
            formula: this.formulaXform,
            colorScale: this.colorScaleXform,
            iconSet: this.iconSetXform,
        };
    }

    get tag(): string {
        return "cfRule";
    }

    /** @returns whether the formulae and format of the rule can be written inline in the x14 extension */
    static isFormulaRule(this: void, rule: ConditionalFormattingRuleModel): rule is FormulaRuleModel {
        return FORMULA_TYPES.includes(rule.type);
    }

    /**
     * Excel 2007 conditional formats cannot reference other sheets: Excel 2010+ writes formula rules that do only in the
     * x14 extension of the worksheet, with their formulae and format inline. Formula rules read from that extension alone
     * are flagged `x14` and written back there too.
     *
     * @returns whether the rule is written in the x14 extension only
     */
    static isX14(this: void, rule: ConditionalFormattingRuleModel): boolean {
        if (!CfRuleXform.isFormulaRule(rule)) {
            return false;
        }
        return rule.x14 === true || (rule.formulae || []).some(referencesOtherSheet);
    }

    /** @returns the type of the `<cfRule>` of a text rule: its operator */
    static textType(this: void, model: ContainsTextRuleModel): string {
        return model.operator || "containsText";
    }

    /** @returns the operator attribute of a text rule, only for those comparing a text */
    static textOperator(this: void, model: ContainsTextRuleModel): string | undefined {
        return textOperatorAttribute(CfRuleXform.textType(model));
    }

    /**
     * An icon set is written in the x14 extension only when it is custom (`custom`, or `icons`: a rule read from a file
     * has both) or uses an icon set that only the extension has (3Stars, 3Triangles, 5Boxes, NoIcons).
     *
     * @returns whether the icon set rule is written in the x14 extension only
     */
    static isX14IconSet(this: void, rule: IconSetRuleModel): boolean {
        return rule.custom === true || rule.icons !== undefined || isX14OnlyIconSet(rule.iconSet);
    }

    /** @returns whether the rule is written in `<conditionalFormatting>` (a data bar is also written in the extension) */
    static isPrimitive(this: void, rule: ConditionalFormattingRuleModel): boolean {
        if (CfRuleXform.isX14(rule)) {
            return false;
        }
        return rule.type !== "iconSet" || !CfRuleXform.isX14IconSet(rule);
    }

    /** Writes the rule (nothing for an unknown type, or an icon set written in the extension). */
    override render(xmlStream: XmlStream, model: ConditionalFormattingRuleModel): void {
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

    /** Writes an `expression` rule (its first formula, if any: the formula is optional in CT_CfRule). */
    renderExpression(xmlStream: XmlStream, model: ExpressionRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "expression",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
        });

        // formula is optional in CT_CfRule
        const formula = model.formulae && model.formulae[0];
        if (formula) {
            this.formulaXform.render(xmlStream, formula);
        }

        xmlStream.closeNode();
    }

    /** Writes a `cellIs` rule. */
    renderCellIs(xmlStream: XmlStream, model: CellIsRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "cellIs",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
            operator: model.operator,
        });

        (model.formulae || []).forEach((formula) => {
            this.formulaXform.render(xmlStream, formula);
        });

        xmlStream.closeNode();
    }

    /** Writes a `top10` rule (`rank` defaults to 10). */
    renderTop10(xmlStream: XmlStream, model: Top10RuleModel): void {
        xmlStream.leafNode(this.tag, {
            type: "top10",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
            percent: BaseXform.toBoolAttribute(model.percent, false),
            bottom: BaseXform.toBoolAttribute(model.bottom, false),
            rank: model.rank === undefined ? 10 : parseInt(String(model.rank), 10),
        });
    }

    /** Writes an `aboveAverage` rule. */
    renderAboveAverage(xmlStream: XmlStream, model: AboveAverageRuleModel): void {
        xmlStream.leafNode(this.tag, {
            type: "aboveAverage",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
            aboveAverage: BaseXform.toBoolAttribute(model.aboveAverage, true),
        });
    }

    /** Writes a `dataBar` rule, with the reference to its x14 version when it has one. */
    renderDataBar(xmlStream: XmlStream, model: DataBarRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "dataBar",
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
        });

        this.databarXform.render(xmlStream, model);
        // reference to the x14 counterpart of this rule, when it has one
        if (DatabarExtXform.isExt(model)) {
            this.extLstRefXform.render(xmlStream, model);
        }

        xmlStream.closeNode();
    }

    /** Writes a `colorScale` rule. */
    renderColorScale(xmlStream: XmlStream, model: ColorScaleRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "colorScale",
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
        });

        this.colorScaleXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    /** Writes an `iconSet` rule, unless it is written in the extension. */
    renderIconSet(xmlStream: XmlStream, model: IconSetRuleModel): void {
        // iconset is all primitive or all extLst
        if (!CfRuleXform.isPrimitive(model)) {
            return;
        }

        xmlStream.openNode(this.tag, {
            type: "iconSet",
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
        });

        this.iconSetXform.render(xmlStream, model);

        xmlStream.closeNode();
    }

    /** Writes a text rule (type: its operator), with its formula (see {@link getTextFormula}). */
    renderText(xmlStream: XmlStream, model: ContainsTextRuleModel): void {
        const operator = CfRuleXform.textOperator(model);
        xmlStream.openNode(this.tag, {
            type: CfRuleXform.textType(model),
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
            operator,
            text: operator ? model.text : undefined,
        });

        const formula = getTextFormula(model);
        if (formula) {
            this.formulaXform.render(xmlStream, formula);
        }

        xmlStream.closeNode();
    }

    /** Writes a `timePeriod` rule, with its formula (see {@link getTimePeriodFormula}). */
    renderTimePeriod(xmlStream: XmlStream, model: TimePeriodRuleModel): void {
        xmlStream.openNode(this.tag, {
            type: "timePeriod",
            dxfId: model.dxfId,
            priority: model.priority,
            stopIfTrue: stopIfTrueAttribute(model),
            timePeriod: model.timePeriod,
        });

        const formula = getTimePeriodFormula(model);
        if (formula) {
            this.formulaXform.render(xmlStream, formula);
        }

        xmlStream.closeNode();
    }

    /** The rule of the attributes (see {@link ruleOfAttributes}: `null` for a type the model does not have), with its `dxfId`. */
    override createNewModel({ attributes }: SaxTag): ConditionalFormattingRuleModel | null {
        const rule = ruleOfAttributes(attributes);
        if (rule) {
            rule.dxfId = BaseXform.toIntValue(attributes.dxfId);
        }
        return rule;
    }

    /** Merges the model of a child into the rule: the settings of a data bar, color scale or icon set into a rule of that type. */
    override onParserClose(name: string): void {
        // the element was opened before its children (createNewModel); null for a rule that is not read
        const { model } = this;
        if (!model) {
            return;
        }
        switch (name) {
            case "dataBar":
                if (model.type === "dataBar") {
                    mergeChild(model, this.databarXform.model);
                }
                break;
            case "extLst": {
                const ref = this.extLstRefXform.model;
                if (ref && ref.x14Id !== undefined) {
                    model.x14Id = ref.x14Id;
                }
                break;
            }
            case "colorScale":
                if (model.type === "colorScale") {
                    mergeChild(model, this.colorScaleXform.model);
                }
                break;
            case "iconSet":
                if (model.type === "iconSet") {
                    mergeChild(model, this.iconSetXform.model);
                }
                break;

            case "formula": {
                // appended to the formulae; set by the start tag of <formula>
                const formula = this.formulaXform.model;
                if (formula !== null && formula !== undefined) {
                    model.formulae = model.formulae || [];
                    model.formulae.push(formula);
                }
                break;
            }
        }
    }
}

/**
 * Merges the model of a child element into its rule (the keys the child has, all settings of the rule: the type of
 * `child` is checked against the rule).
 */
function mergeChild<T extends object>(this: void, rule: T, child: Partial<T> | null | undefined): void {
    if (child) {
        Object.assign(rule, child);
    }
}

export { ruleOfAttributes, mergeChild, getTextFormula, getTimePeriodFormula };
export default CfRuleXform;
