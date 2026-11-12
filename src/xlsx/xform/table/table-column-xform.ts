import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { TableColumnProperties } from "../../../../index";
import type { ListItemOptions, TableColumnModel } from "../../../types";
import { isOneOf } from "../../../utils/utils";
import BaseXform from "../base-xform";

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

/**
 * A `<tableColumn id name totalsRowLabel totalsRowFunction dxfId>` of a table (see {@link TableColumnModel}). Its
 * children (calculatedColumnFormula, totalsRowFormula...) are skipped when parsing. A `totalsRowFunction` that is not one
 * of ST_TotalsRowFunction parses as absent (the default: `none`).
 */
class TableColumnXform extends BaseXform<TableColumnModel> {
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
        xmlStream.leafNode(this.tag, {
            id: model.id.toString(),
            name: model.name,
            totalsRowLabel: model.totalsRowLabel,
            totalsRowFunction: model.totalsRowFunction,
            dxfId: model.dxfId,
        });
    }

    /**
     * @returns `true` for `<tableColumn>`
     * @throws Error for a column without name (CT_TableColumn requires it)
     */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            const { name, totalsRowLabel, totalsRowFunction, dxfId } = node.attributes;
            if (name === undefined) {
                throw new Error("Invalid table: <tableColumn> without name");
            }
            this.model = {
                name,
                totalsRowLabel,
                totalsRowFunction: isOneOf(TOTALS_ROW_FUNCTIONS, totalsRowFunction) ? totalsRowFunction : undefined,
                dxfId,
            };
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        // a tableColumn may have children (calculatedColumnFormula, totalsRowFormula, ...):
        // the column only ends with its own closing tag
        return name !== this.tag;
    }
}

export default TableColumnXform;
