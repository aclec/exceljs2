import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { FilterColumnModel, TableModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";

import FilterColumnXform from "./filter-column-xform";

/** The auto filter of a table as {@link AutoFilterXform} parses it: its range and its (sparse) filter columns. */
export interface TableAutoFilterModel {
    autoFilterRef?: string;
    columns: FilterColumnModel[];
}

/**
 * The auto filter of a table, `<autoFilter ref>` with a `<filterColumn>` per column of the table (rendered from the
 * table model). Parsing keeps the filter columns; the other children (sortState, extLst...) are skipped.
 */
class AutoFilterXform extends BaseXform<TableAutoFilterModel> {
    declare map: { readonly filterColumn: FilterColumnXform };

    /** While parsing: the depth inside an element that is skipped (0: none). */
    private skipDepth = 0;

    constructor() {
        super();

        this.map = {
            filterColumn: new FilterColumnXform(),
        };
    }

    get tag(): string {
        return "autoFilter";
    }

    /** Sets the `colId` of the columns of the table. */
    override prepare(model: TableModel): void {
        model.columns.forEach((column, index) => {
            this.map.filterColumn.prepare(column, { index });
        });
    }

    override render(xmlStream: XmlStream, model: TableModel): void {
        xmlStream.openNode(this.tag, { ref: model.autoFilterRef });

        for (const column of model.columns) {
            this.map.filterColumn.render(xmlStream, column);
        }

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<autoFilter>` is accepted (the unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth++;
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.skipDepth = 0;
                this.model = {
                    autoFilterRef: node.attributes.ref,
                    columns: [],
                };
                return true;

            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                    return true;
                }
                // sortState, extLst... are not modelled: skip the subtree
                this.skipDepth = 1;
                return true;
            }
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @throws Error on an end tag that does not match (malformed XML) */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the filter columns are inside <autoFilter>, which created the model
                this.model!.columns.push(this.map.filterColumn.model!);
                this.parser = undefined;
            }
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth--;
            return true;
        }
        switch (name) {
            case this.tag:
                return false;
            default:
                throw new Error(`Unexpected xml node in parseClose: ${name}`);
        }
    }
}

export default AutoFilterXform;
