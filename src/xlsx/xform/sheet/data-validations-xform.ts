import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DataValidationFormula, DataValidationModel, DataValidationsModel, WorksheetLoadBudget } from "../../../types";
import _ from "../../../utils/under-dash";
import utils from "../../../utils/utils";
import colCache from "../../../utils/col-cache";
import BaseXform from "../base-xform";
import Range from "../../../doc/range";
import { wholeRangeAddress } from "../../../doc/data-validations";

const addressRegex = /^[A-Z]{1,3}\d+$/;
const rangeRegex = /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/;

// A number literal of a formula: an optional sign, digits with an optional fraction (or a fraction alone), an optional
// exponent. Anchored, and no two quantifiers can match the same characters: linear on any input.
const numberLiteralRegex = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * A formula of a `<dataValidation>` read from a file: for `whole`, `decimal` and `textLength` a number literal is a
 * number, for `date` it is a serial date, read as a Date; any other formula (`$B$1`, `TODAY()`, `A1+1`, a literal out of
 * range) is its text.
 */
function parseFormula(type: DataValidationModel["type"], text: string): DataValidationFormula {
    switch (type) {
        case "whole":
        case "decimal":
        case "textLength":
        case "date": {
            if (!numberLiteralRegex.test(text)) {
                return text;
            }
            const value = Number(text);
            if (type !== "date") {
                return Number.isFinite(value) ? value : text;
            }
            const date = utils.excelToDate(value);
            return Number.isNaN(date.getTime()) ? text : date;
        }
        case "any":
        case "list":
        case "custom":
        default:
            return text;
    }
}

/**
 * The text of a formula of a `<dataValidation>`: a string as is (a formula, or a literal), a number as a number (for a
 * `date` rule, as milliseconds since 1970 like `new Date(number)`: its serial date), a Date as its serial date.
 *
 * @throws Error for an invalid Date (or a number that is not one, for a `date` rule), or a number that is not finite
 */
function formulaText(type: DataValidationModel["type"], formula: DataValidationFormula): string {
    if (typeof formula === "string") {
        return formula;
    }
    if (formula instanceof Date || type === "date") {
        const date = formula instanceof Date ? formula : new Date(formula);
        if (Number.isNaN(date.getTime())) {
            throw new Error(`Invalid date in the formulae of a data validation: ${String(formula)}`);
        }
        return String(utils.dateToExcel(date));
    }
    if (!Number.isFinite(formula)) {
        throw new Error(`A formula of a data validation is not a finite number: ${String(formula)}`);
    }
    return String(formula);
}

// Ranges read from a file are expanded to one model entry per cell only while the
// workbook stays under this budget; larger ranges are kept as 'A1:B2' keys.
const MAX_EXPANDED_CELLS = 100000;

/** The attributes of a `<dataValidation>` copied as strings. */
type StringAttribute = "operator" | "promptTitle" | "prompt" | "errorStyle" | "errorTitle" | "error";
/** The attributes of a `<dataValidation>` parsed as booleans. */
type BooleanAttribute = "allowBlank" | "showInputMessage" | "showErrorMessage";

function assign(
    dataValidation: DataValidationModel,
    attributes: Readonly<Record<string, string | undefined>>,
    name: StringAttribute,
    defaultValue?: string
): void {
    const value = attributes[name];
    // the attribute strings of the file, unvalidated (operator, errorStyle)
    const target = dataValidation as Record<StringAttribute, string | undefined>;
    if (value !== undefined) {
        target[name] = value;
    } else if (defaultValue !== undefined) {
        target[name] = defaultValue;
    }
}

function assignBool(
    dataValidation: DataValidationModel,
    attributes: Readonly<Record<string, string | undefined>>,
    name: BooleanAttribute
): void {
    const value = attributes[name];
    if (value !== undefined) {
        dataValidation[name] = utils.parseBoolean(value);
    }
}

/** A data validation of the model, with the rectangle (`sqref`) it is written for. */
type SqrefDataValidation = DataValidationModel & { sqref: string };

/** An entry of the model while it is squeezed. */
interface DvEntry {
    address: string;
    dataValidation: DataValidationModel;
    marked: boolean;
}

/**
 * Squeezes alike data validations together into rectangular ranges, to reduce file size and speed up Excel load time.
 *
 * @throws Error for a key of the model that is not a cell address or a range in its normal form (`A1`, `A1:B2`)
 */
function optimiseDataValidations(model: DataValidationsModel | undefined): SqrefDataValidation[] {
    const keys = model ? Object.keys(model) : [];
    const dvList: DvEntry[] = [];
    for (const address of keys) {
        // removed validations are kept as undefined entries
        const dataValidation = model![address];
        if (dataValidation) {
            dvList.push({ address, dataValidation, marked: false });
        }
    }
    dvList.sort((a, b) => _.strcmp(a.address, b.address));
    // the model is not empty (dvList has entries) from here on
    const dvs = model!;
    const dvMap = new Map(dvList.map((dv) => [dv.address, dv]));
    const matchCol = (addr: { row: number; address: string }, height: number, col: number): boolean => {
        for (let i = 0; i < height; i++) {
            const otherAddress = colCache.encodeAddress(addr.row + i, col);
            if (!dvs[otherAddress] || !_.isEqual(dvs[addr.address], dvs[otherAddress])) {
                return false;
            }
        }
        return true;
    };
    const result: SqrefDataValidation[] = [];
    for (const dv of dvList) {
        if (dv.marked) {
            continue;
        }
        const addr = colCache.decodeEx(dv.address);
        if ("dimensions" in addr) {
            if (!dvMap.has(addr.dimensions)) {
                // not a range in its normal form (e.g. "$A$1:B2", "B2:A1")
                throw new Error(`Invalid data validation address: ${dv.address}`);
            }
            result.push({
                ...dv.dataValidation,
                sqref: dv.address,
            });
            continue;
        }
        if ("error" in addr) {
            throw new Error(`Invalid data validation address: ${dv.address}`);
        }

        // iterate downwards - finding matching cells
        let height = 1;
        let otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        while (dvs[otherAddress] && _.isEqual(dv.dataValidation, dvs[otherAddress])) {
            height++;
            otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        }

        // iterate rightwards...

        let width = 1;
        while (matchCol(addr, height, addr.col + width)) {
            width++;
        }

        // mark all included addresses (entries of the model, found by matchCol; the first one is the key itself, unless it
        // is not a cell address in its normal form)
        for (let i = 0; i < height; i++) {
            for (let j = 0; j < width; j++) {
                otherAddress = colCache.encodeAddress(addr.row + i, addr.col + j);
                const entry = dvMap.get(otherAddress);
                if (!entry) {
                    throw new Error(`Invalid data validation address: ${dv.address}`);
                }
                entry.marked = true;
            }
        }

        if (height > 1 || width > 1) {
            const bottom = addr.row + (height - 1);
            const right = addr.col + (width - 1);
            result.push({
                ...dv.dataValidation,
                sqref: `${dv.address}:${colCache.encodeAddress(bottom, right)}`,
            });
        } else {
            result.push({
                ...dv.dataValidation,
                sqref: dv.address,
            });
        }
    }
    return result;
}

/** The options of a {@link DataValidationsXform}. */
export interface DataValidationsXformOptions {
    /**
     * The budget of expanded cells, shared by all the worksheets of a workbook being loaded (a budget per
     * `<dataValidations>` element would be multiplied by the number of worksheets). Default: a budget of its own for
     * each `<dataValidations>`.
     */
    loadBudget?: Pick<WorksheetLoadBudget, "dataValidationCells">;
}

/**
 * `<dataValidations>` of a worksheet. Rendering squeezes alike validations of adjacent cells into rectangles. Parsing
 * gives the model keyed by cell address, sharing one validation object between the cells of a `<dataValidation>`; ranges
 * are expanded per cell within a budget of cells ({@link DataValidationsXform.MAX_EXPANDED_CELLS} per workbook), larger
 * ones are kept as range keys (never expanded: a hostile `A1:XFD1048576` stays one key). Whole columns / rows are read as
 * the rectangle they cover.
 */
class DataValidationsXform extends BaseXform<DataValidationsModel> {
    /** The budget of cells that the ranges of a workbook being loaded are expanded to. */
    static readonly MAX_EXPANDED_CELLS = MAX_EXPANDED_CELLS;

    /** The budget shared by the worksheets of a workbook being loaded, if any. */
    readonly loadBudget: Pick<WorksheetLoadBudget, "dataValidationCells"> | undefined;

    /** While parsing: the budget of the current `<dataValidations>`. */
    declare _budget: Pick<WorksheetLoadBudget, "dataValidationCells">;
    /** While parsing: the `sqref` of the current `<dataValidation>`. */
    declare _address: string | undefined;
    /** While parsing: the current validation. */
    declare _dataValidation: DataValidationModel;
    /** While parsing: the chunks of text of the current `<formula1>` / `<formula2>`, if any. */
    declare _formula: string[] | undefined;

    constructor(options?: DataValidationsXformOptions) {
        super();
        this.loadBudget = options && options.loadBudget;
    }

    get tag(): string {
        return "dataValidations";
    }

    /**
     * Writes the validations (nothing when there is none); a formula that is a Date (or a number, for a `date` rule) is
     * written as a serial date.
     *
     * @throws Error for an invalid key (see {@link optimiseDataValidations}) or formula (see {@link formulaText})
     */
    override render(xmlStream: XmlStream, model: DataValidationsModel | undefined): void {
        const optimizedModel = optimiseDataValidations(model);
        if (optimizedModel.length) {
            xmlStream.openNode("dataValidations", { count: optimizedModel.length });

            for (const value of optimizedModel) {
                xmlStream.openNode("dataValidation");

                if (value.type !== "any") {
                    xmlStream.addAttribute("type", value.type);

                    if (value.operator && value.type !== "list" && value.operator !== "between") {
                        xmlStream.addAttribute("operator", value.operator);
                    }
                    if (value.allowBlank) {
                        xmlStream.addAttribute("allowBlank", "1");
                    }
                }
                if (value.showInputMessage) {
                    xmlStream.addAttribute("showInputMessage", "1");
                }
                if (value.promptTitle) {
                    xmlStream.addAttribute("promptTitle", value.promptTitle);
                }
                if (value.prompt) {
                    xmlStream.addAttribute("prompt", value.prompt);
                }
                if (value.showErrorMessage) {
                    xmlStream.addAttribute("showErrorMessage", "1");
                }
                if (value.errorStyle) {
                    xmlStream.addAttribute("errorStyle", value.errorStyle);
                }
                if (value.errorTitle) {
                    xmlStream.addAttribute("errorTitle", value.errorTitle);
                }
                if (value.error) {
                    xmlStream.addAttribute("error", value.error);
                }
                xmlStream.addAttribute("sqref", value.sqref);
                (value.formulae || []).forEach((formula: DataValidationFormula, index) => {
                    xmlStream.openNode(`formula${index + 1}`);
                    xmlStream.writeText(formulaText(value.type, formula));
                    xmlStream.closeNode();
                });
                xmlStream.closeNode();
            }
            xmlStream.closeNode();
        }
    }

    /** @returns `true` for `<dataValidations>`, `<dataValidation>`, `<formula1>` and `<formula2>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "dataValidations":
                // a plain object (it becomes the public worksheet.dataValidations.model): its keys are only cell and range
                // references matched by addressRegex / rangeRegex, never raw strings of the file
                this.model = {};
                this._budget = this.loadBudget || { dataValidationCells: MAX_EXPANDED_CELLS };
                return true;

            case "dataValidation": {
                const { attributes } = node;
                this._address = attributes.sqref;
                // the type attribute of the file, unvalidated
                const dataValidation: DataValidationModel = {
                    type: (attributes.type || "any") as DataValidationModel["type"],
                    formulae: [],
                };

                if (attributes.type) {
                    assignBool(dataValidation, attributes, "allowBlank");
                }
                assignBool(dataValidation, attributes, "showInputMessage");
                assignBool(dataValidation, attributes, "showErrorMessage");

                switch (dataValidation.type) {
                    case "any":
                    case "list":
                    case "custom":
                        break;
                    case "whole":
                    case "decimal":
                    case "date":
                    case "textLength":
                    default:
                        assign(dataValidation, attributes, "operator", "between");
                        break;
                }
                assign(dataValidation, attributes, "promptTitle");
                assign(dataValidation, attributes, "prompt");
                assign(dataValidation, attributes, "errorStyle");
                assign(dataValidation, attributes, "errorTitle");
                assign(dataValidation, attributes, "error");

                this._dataValidation = dataValidation;
                return true;
            }

            case "formula1":
            case "formula2":
                this._formula = [];
                return true;

            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this._formula) {
            this._formula.push(text);
        }
    }

    /** @returns `false` at the end of `<dataValidations>`, else `true` */
    override parseClose(name: string): boolean {
        switch (name) {
            case "dataValidations":
                return false;
            case "dataValidation": {
                const dataValidation = this._dataValidation;
                // created by <dataValidations>, which holds this element
                const model = this.model!;
                const budget = this._budget;
                if (!dataValidation.formulae || !dataValidation.formulae.length) {
                    delete dataValidation.formulae;
                    delete dataValidation.operator;
                }
                // The four known cases: 1. E4:L9 N4:U9  2.E4 L9  3. N4:U9  4. E4
                const list = (this._address || "").split(/\s+/g);
                for (const token of list) {
                    // whole columns / rows ('D:E', '4:4') are read as the rectangle they cover; outside the grid
                    // (null), they are ignored like the other tokens that are not cell references
                    const addr = wholeRangeAddress(token.replace(/\$/g, "")) || "";
                    if (rangeRegex.test(addr)) {
                        const range = new Range(addr);
                        if (range.count <= budget.dataValidationCells) {
                            budget.dataValidationCells -= range.count;
                            range.forEachAddress((address: string) => {
                                model[address] = dataValidation;
                            });
                        } else {
                            // Too large to expand per cell (e.g. a whole column, or a hostile
                            // A1:XFD1048576): keep the range as a single key. DataValidations.find
                            // and the writer both understand range keys.
                            model[range.range] = dataValidation;
                        }
                    } else if (addressRegex.test(addr)) {
                        model[addr] = dataValidation;
                    }
                }
                return true;
            }
            case "formula1":
            case "formula2": {
                const formula = parseFormula(this._dataValidation.type, (this._formula || []).join(""));
                // a formula is read inside <dataValidation>, which created the formulae
                this._dataValidation.formulae!.push(formula);
                this._formula = undefined;
                return true;
            }
            default:
                return true;
        }
    }
}

export default DataValidationsXform;
