import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { TableColumnProperties } from "../../../../index";
import type { ListItemOptions, TableColumnModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * A `<tableColumn id name totalsRowLabel totalsRowFunction dxfId>` of a table (see {@link TableColumnModel}). Its
 * children (calculatedColumnFormula, totalsRowFormula...) are skipped when parsing.
 */
class TableColumnXform extends BaseXform<TableColumnModel> {
    get tag(): string {
        return "tableColumn";
    }

    /** Sets `id`: the 1-based index of the column (`options.index` from ListXform). */
    override prepare(model: TableColumnModel, options: ListItemOptions): void {
        model.id = options.index! + 1;
    }

    /** @param model - prepared (it has its `id`) */
    override render(xmlStream: XmlStream, model: TableColumnModel): void {
        xmlStream.leafNode(this.tag, {
            id: model.id!.toString(),
            name: model.name,
            totalsRowLabel: model.totalsRowLabel,
            totalsRowFunction: model.totalsRowFunction,
            dxfId: model.dxfId,
        });
    }

    /** @returns `true` for `<tableColumn>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            const { attributes } = node;
            this.model = {
                // name is required (CT_TableColumn); the values are unvalidated
                name: attributes.name!,
                totalsRowLabel: attributes.totalsRowLabel,
                totalsRowFunction: attributes.totalsRowFunction as TableColumnProperties["totalsRowFunction"],
                dxfId: attributes.dxfId,
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
