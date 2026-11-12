import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { ListItemOptions, StyleLookup, TableColumnModel, TableModel } from "../../../types";

import BaseXform, { ownEntry } from "../base-xform";
import InvalidTablePartError from "./invalid-table-part-error";
import ListXform from "../list-xform";

import AutoFilterXform from "./auto-filter-xform";
import TableColumnXform from "./table-column-xform";
import TableStyleInfoXform from "./table-style-info-xform";

/** The reconcile options of a {@link TableXform}: the parsed styles (for the dxfs of the columns). */
export interface TableReconcileOptions {
    readonly styles: StyleLookup;
}

/**
 * A table part, `xl/tables/tableN.xml` (CT_Table order: autoFilter, tableColumns, tableStyleInfo), see
 * {@link TableModel}. The auto filter is written only with a header row; parsing sets the `filterButton` of the columns
 * from it. `totalsRowCount` and `headerRowCount` are counts (a table has a totals row when its count is not 0).
 */
class TableXform extends BaseXform<TableModel> {
    /** The namespaces of `<table>`. */
    static readonly TABLE_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "xr xr3",
        "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "xmlns:xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
        // 'xr:uid': '{00000000-000C-0000-FFFF-FFFF00000000}',
    };

    declare map: {
        readonly autoFilter: AutoFilterXform;
        readonly tableColumns: ListXform<TableColumnModel>;
        readonly tableStyleInfo: TableStyleInfoXform;
    };

    constructor() {
        super();

        this.map = {
            autoFilter: new AutoFilterXform(),
            tableColumns: new ListXform({
                tag: "tableColumns",
                count: true,
                empty: true,
                childXform: new TableColumnXform(),
            }),
            tableStyleInfo: new TableStyleInfoXform(),
        };
    }

    /** Sets the `colId` (filter) and `id` of the columns. */
    override prepare(model: TableModel, options: ListItemOptions): void {
        this.map.autoFilter.prepare(model);
        this.map.tableColumns.prepare(model.columns, options);
    }

    get tag(): string {
        return "table";
    }

    override render(xmlStream: XmlStream, model: TableModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, {
            ...TableXform.TABLE_ATTRIBUTES,
            id: model.id,
            name: model.name,
            displayName: model.displayName || model.name,
            ref: model.tableRef,
            totalsRowCount: model.totalsRow ? "1" : undefined,
            totalsRowShown: model.totalsRow ? undefined : "1",
            headerRowCount: model.headerRow ? "1" : "0",
        });

        // the autoFilter buttons live in the header row: a table without one has no autoFilter
        if (model.headerRow) {
            this.map.autoFilter.render(xmlStream, model);
        }
        this.map.tableColumns.render(xmlStream, model.columns);
        // tableStyleInfo is optional: a table read from a file without it has no table style (`style: null`); the
        // document table defaults the style of the tables it creates (see Table.validate)
        if (model.style) {
            this.map.tableStyleInfo.render(xmlStream, model.style);
        }

        xmlStream.closeNode();
    }

    /**
     * @returns `true`: every element of the part is accepted (unknown ones are skipped)
     * @throws Error for a table without name (CT_Table requires it)
     */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        const { name, attributes } = node;
        switch (name) {
            case this.tag: {
                this.reset();
                const { name: tableName } = attributes;
                if (tableName === undefined) {
                    throw new InvalidTablePartError("Invalid table part: <table> without name");
                }
                this.model = {
                    name: tableName,
                    displayName: attributes.displayName || tableName,
                    tableRef: attributes.ref,
                    // xsd:unsignedInt, default 0
                    totalsRow: BaseXform.toIntValue(attributes.totalsRowCount, 0) > 0,
                    // xsd:unsignedInt, default 1 (Excel omits it): a count that is not a number is the default
                    headerRow: BaseXform.toIntValue(attributes.headerRowCount, 1) !== 0,
                    // set when the element closes
                    columns: [],
                };
                break;
            }
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                break;
            }
        }
        return true;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @throws InvalidTablePartError for a table without `<tableColumns>` (CT_Table requires it) */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag: {
                // the model was created when the element opened; the children were reset then (null when absent)
                const { model } = this;
                const columns = this.map.tableColumns.model;
                if (!model) {
                    // a </table> without its <table> (not well-formed XML: the parser rejects it)
                    return false;
                }
                if (!columns) {
                    throw new InvalidTablePartError(`Invalid table ${JSON.stringify(model.name)}: no <tableColumns>`);
                }
                model.columns = columns;
                const autoFilter = this.map.autoFilter.model;
                if (autoFilter) {
                    model.autoFilterRef = autoFilter.autoFilterRef;
                    // with an autoFilter every column shows its button unless a filterColumn hides it
                    for (const column of columns) {
                        column.filterButton = true;
                    }
                    // filterColumns are sparse: match them to table columns by colId
                    for (const filterColumn of autoFilter.columns) {
                        const column = columns[BaseXform.toIntValue(filterColumn.colId, NaN)];
                        if (column) {
                            column.filterButton = filterColumn.filterButton;
                        }
                    }
                }
                model.style = this.map.tableStyleInfo.model;
                return false;
            }
            default:
                // could be some unrecognised tags
                return true;
        }
    }

    /** Resolves the `dxfId` of each column (an index into the dxfs of the styles part) to its `style`. */
    override reconcile(model: TableModel, options: TableReconcileOptions): void {
        // fetch the dfxs from styles
        model.columns.forEach((column) => {
            if (column.dxfId !== undefined) {
                // parsed: the attribute string (a number on write)
                const dxfId = typeof column.dxfId === "number" ? column.dxfId : parseInt(column.dxfId, 10);
                column.style = options.styles.getDxfStyle(dxfId);
            }
        });
    }
}

export default TableXform;
