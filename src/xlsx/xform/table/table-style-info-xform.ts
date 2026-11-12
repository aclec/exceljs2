import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { TableStyleProperties } from "../../../../index";
import BaseXform from "../base-xform";

/**
 * The style of a table, `<tableStyleInfo name showFirstColumn...>`: the table style is `theme` in the model (`null` when
 * absent).
 */
class TableStyleInfoXform extends BaseXform<TableStyleProperties> {
    get tag(): string {
        return "tableStyleInfo";
    }

    override render(xmlStream: XmlStream, model: TableStyleProperties): void {
        xmlStream.leafNode(this.tag, {
            name: model.theme ? model.theme : undefined,
            showFirstColumn: model.showFirstColumn ? "1" : "0",
            showLastColumn: model.showLastColumn ? "1" : "0",
            showRowStripes: model.showRowStripes ? "1" : "0",
            showColumnStripes: model.showColumnStripes ? "1" : "0",
        });
    }

    /** @returns `true` for `<tableStyleInfo>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            const { attributes } = node;
            this.model = {
                // unvalidated
                theme: attributes.name ? (attributes.name as TableStyleProperties["theme"]) : null,
                showFirstColumn: attributes.showFirstColumn === "1",
                showLastColumn: attributes.showLastColumn === "1",
                showRowStripes: attributes.showRowStripes === "1",
                showColumnStripes: attributes.showColumnStripes === "1",
            };
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default TableStyleInfoXform;
