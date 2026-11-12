import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { WorkbookProperties } from "../../../../index";
import BaseXform from "../base-xform";

/** `<workbookPr>` of `xl/workbook.xml`: the date system (`date1904`), with fixed theme and privacy settings. */
class WorksheetPropertiesXform extends BaseXform<Partial<WorkbookProperties>> {
    override render(xmlStream: XmlStream, model: Partial<WorkbookProperties>): void {
        xmlStream.leafNode("workbookPr", {
            date1904: model.date1904 ? 1 : undefined,
            defaultThemeVersion: 164011,
            filterPrivacy: 1,
        });
    }

    /** @returns `true` for `<workbookPr>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "workbookPr") {
            this.model = {
                date1904: node.attributes.date1904 === "1",
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

export default WorksheetPropertiesXform;
