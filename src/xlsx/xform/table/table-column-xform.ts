import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { TableColumnProperties } from "../../../../index";
import type { ListItemOptions, TableColumnModel } from "../../../types";
import { MAX_FORMULA_LENGTH } from "../../../utils/shared-formula";
import { isOneOf, parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";
import InvalidTablePartError from "./invalid-table-part-error";

// the totals row functions of ECMA-376 ST_TotalsRowFunction
const TOTALS_ROW_FUNCTIONS: Readonly<Record<NonNullable<TableColumnProperties["totalsRowFunction"]>, true>> = {
    none: true,
    sum: true,
    min: true,
    max: true,
    average: true,
    count: true,
    countNums: true,
    stdDev: true,
    var: true,
    custom: true,
};

// the formula children of CT_TableColumn (CT_TableFormula), in schema order, with the model flag of their `array` attribute
const TABLE_FORMULAS = {
    calculatedColumnFormula: "calculatedColumnFormulaArray",
    totalsRowFormula: "totalsRowFormulaArray",
} as const;
type TableFormulaName = keyof typeof TABLE_FORMULAS;
const TABLE_FORMULA_NAMES: Readonly<Record<TableFormulaName, true>> = { calculatedColumnFormula: true, totalsRowFormula: true };

/** A formula child being parsed. */
interface TableFormulaBeingParsed {
    name: TableFormulaName;
    text: string;
    array: boolean;
}

/**
 * A `<tableColumn id name totalsRowLabel totalsRowFunction dataDxfId>` of a table (see {@link TableColumnModel}), with its
 * `<calculatedColumnFormula>` and `<totalsRowFormula>` children (CT_TableColumn sequence; the totals row formula is written
 * for the `custom` function only: the other functions are SUBTOTALs of the cells). The style of the data cells is
 * `dataDxfId`; the `dxfId` attribute that exceljs 4.x / 5.0 wrote (not in CT_TableColumn) is still read. The other children
 * (xmlColumnPr, extLst) are skipped. A `totalsRowFunction` that is not one of ST_TotalsRowFunction parses as absent (the
 * default: `none`).
 */
class TableColumnXform extends BaseXform<TableColumnModel> {
    /** While parsing: the number of open elements of the column (its own included), 0 outside a column. */
    depth = 0;

    /** While parsing: the formula child being read (a direct child of the column). */
    formula: TableFormulaBeingParsed | undefined = undefined;

    get tag(): string {
        return "tableColumn";
    }

    /**
     * Sets `id`: the 1-based index of the column (`options.index` from ListXform).
     *
     * @throws Error without `options.index`
     */
    override prepare(model: TableColumnModel, options: ListItemOptions): void {
        if (options.index === undefined) {
            throw new Error("TableColumnXform.prepare needs the index of the column (options.index)");
        }
        model.id = options.index + 1;
    }

    /**
     * @param model - prepared (it has its `id`)
     * @throws Error for a column without `id` (not prepared)
     */
    override render(xmlStream: XmlStream, model: TableColumnModel): void {
        if (model.id === undefined) {
            throw new Error(`Table column ${JSON.stringify(model.name)} has no id: prepare the table before writing it`);
        }
        const attributes = {
            id: model.id.toString(),
            name: model.name,
            totalsRowLabel: model.totalsRowLabel,
            totalsRowFunction: model.totalsRowFunction,
            dataDxfId: model.dxfId,
        };
        const formulas: Record<TableFormulaName, string | undefined> = {
            calculatedColumnFormula: model.calculatedColumnFormula,
            totalsRowFormula: model.totalsRowFunction === "custom" ? model.totalsRowFormula : undefined,
        };
        if (!formulas.calculatedColumnFormula && !formulas.totalsRowFormula) {
            xmlStream.leafNode(this.tag, attributes);
            return;
        }
        xmlStream.openNode(this.tag, attributes);
        for (const name of ["calculatedColumnFormula", "totalsRowFormula"] as const) {
            const formula = formulas[name];
            if (formula) {
                xmlStream.leafNode(name, { array: model[TABLE_FORMULAS[name]] ? "1" : undefined }, formula);
            }
        }
        xmlStream.closeNode();
    }

    /**
     * @returns `true` for `<tableColumn>` and the elements inside it
     * @throws InvalidTablePartError for a column without name (CT_TableColumn requires it)
     */
    override parseOpen(node: SaxTag): boolean {
        if (this.depth > 0) {
            // a child of the column: only its direct formula children are read
            this.depth++;
            if (this.depth === 2 && isOneOf(TABLE_FORMULA_NAMES, node.name)) {
                this.formula = { name: node.name, text: "", array: parseBoolean(node.attributes.array) };
            }
            return true;
        }
        if (node.name === this.tag) {
            // dataDxfId (CT_TableColumn), else the dxfId that exceljs 4.x / 5.0 wrote
            const { name, totalsRowLabel, totalsRowFunction, dataDxfId, dxfId } = node.attributes;
            if (name === undefined) {
                throw new InvalidTablePartError("Invalid table: <tableColumn> without name");
            }
            this.model = {
                name,
                totalsRowLabel,
                totalsRowFunction: isOneOf(TOTALS_ROW_FUNCTIONS, totalsRowFunction) ? totalsRowFunction : undefined,
                dxfId: dataDxfId ?? dxfId,
            };
            this.depth = 1;
            this.formula = undefined;
            return true;
        }
        return false;
    }

    /** @throws InvalidTablePartError for a formula longer than the Excel limit ({@link MAX_FORMULA_LENGTH}) */
    override parseText(text: string): void {
        const { formula, model } = this;
        if (formula && this.depth === 2) {
            formula.text += text;
            if (formula.text.length > MAX_FORMULA_LENGTH) {
                throw new InvalidTablePartError(
                    `Invalid table: the ${formula.name} of column ${JSON.stringify(model?.name)} is longer than ${MAX_FORMULA_LENGTH} characters`
                );
            }
        }
    }

    override parseClose(): boolean {
        if (this.depth > 1) {
            const { formula, model } = this;
            if (this.depth === 2 && formula) {
                // an empty formula is absent; array: xsd:boolean, default false
                if (model && formula.text) {
                    model[formula.name] = formula.text;
                    if (formula.array) {
                        model[TABLE_FORMULAS[formula.name]] = true;
                    }
                }
                this.formula = undefined;
            }
            this.depth--;
            return true;
        }
        // the column ends with its own closing tag (the elements are well-formed)
        this.depth = 0;
        return false;
    }
}

export default TableColumnXform;
