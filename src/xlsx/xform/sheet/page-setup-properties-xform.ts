import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** The model of `<pageSetUpPr>`. */
export interface PageSetupPropertiesModel {
    /** Fit the printout to the number of pages of the page setup (`fitToWidth` / `fitToHeight`). */
    fitToPage?: boolean;
}

/** `<pageSetUpPr fitToPage="1"/>` of `<sheetPr>`: rendered only when `fitToPage` is set. */
class PageSetupPropertiesXform extends BaseXform<PageSetupPropertiesModel> {
    get tag(): string {
        return "pageSetUpPr";
    }

    /** @returns whether the element was written */
    override render(xmlStream: XmlStream, model: PageSetupPropertiesModel | undefined): boolean {
        if (model && model.fitToPage) {
            xmlStream.leafNode(this.tag, {
                fitToPage: model.fitToPage ? "1" : undefined,
            });
            return true;
        }
        return false;
    }

    /** @returns `true` for `<pageSetUpPr>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            this.model = {
                fitToPage: node.attributes.fitToPage === "1",
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

export default PageSetupPropertiesXform;
