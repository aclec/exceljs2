import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { WorkbookViewModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * A `<workbookView>` of `xl/workbook.xml` (see {@link WorkbookViewModel}). Render defaults: position 0, 0, size 12000 x
 * 24000; `visibility` is written unless `visible`. Parse defaults: position 0, 0, size 25000 x 10000, `visible`.
 */
class WorkbookViewXform extends BaseXform<WorkbookViewModel> {
    override render(xmlStream: XmlStream, model: Partial<WorkbookViewModel>): void {
        xmlStream.leafNode("workbookView", {
            xWindow: model.x || 0,
            yWindow: model.y || 0,
            windowWidth: model.width || 12000,
            windowHeight: model.height || 24000,
            firstSheet: model.firstSheet,
            activeTab: model.activeTab,
            visibility: model.visibility && model.visibility !== "visible" ? model.visibility : undefined,
        });
    }

    /** @returns `true` for `<workbookView>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "workbookView") {
            const { xWindow, yWindow, windowWidth, windowHeight, visibility, activeTab, firstSheet } = node.attributes;
            const model: WorkbookViewModel = {
                x: BaseXform.toIntValue(xWindow, 0),
                y: BaseXform.toIntValue(yWindow, 0),
                width: BaseXform.toIntValue(windowWidth, 25000),
                height: BaseXform.toIntValue(windowHeight, 10000),
                visibility: visibility !== undefined ? visibility : "visible",
            };
            if (activeTab !== undefined) {
                model.activeTab = parseInt(activeTab, 10);
            }
            if (firstSheet !== undefined) {
                model.firstSheet = parseInt(firstSheet, 10);
            }
            this.model = model;
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default WorkbookViewXform;
