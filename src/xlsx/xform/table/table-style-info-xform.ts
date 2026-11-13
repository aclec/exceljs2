import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { TableStyleProperties } from "../../../../index";
import { parseBoolean } from "../../../utils/utils";
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
                // a built-in or a custom table style (of the styles part)
                theme: attributes.name || null,
                showFirstColumn: parseBoolean(attributes.showFirstColumn),
                showLastColumn: parseBoolean(attributes.showLastColumn),
                showRowStripes: parseBoolean(attributes.showRowStripes),
                showColumnStripes: parseBoolean(attributes.showColumnStripes),
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
