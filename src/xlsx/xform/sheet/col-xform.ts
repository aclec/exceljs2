import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CellPrepareOptions, CellReconcileOptions, ColModel } from "../../../types";
import utils from "../../../utils/utils";
import colCache from "../../../utils/col-cache";
import BaseXform from "../base-xform";

// clamp to the worksheet grid: <col max="2000000000"/> must not allocate 2e9 columns
function clampColumn(value: string | undefined): number {
    return Math.min(Math.max(parseInt(value || "0", 10) || 0, 0), colCache.MAX_COLS);
}

/**
 * A `<col>` of `<cols>`: the definition (width, style, outline...) shared by the columns `min`..`max`. `customWidth` is
 * always written. Parsing clamps `min` / `max` to 0..{@link colCache.MAX_COLS}.
 */
class ColXform extends BaseXform<ColModel> {
    get tag(): string {
        return "col";
    }

    /** Registers the style of the columns (`styleId`, when not the default). */
    override prepare(model: ColModel, options: Pick<CellPrepareOptions, "styles">): void {
        const styleId = options.styles.addStyleModel(model.style || {});
        if (styleId) {
            model.styleId = styleId;
        }
    }

    override render(xmlStream: XmlStream, model: ColModel): void {
        xmlStream.openNode("col");
        xmlStream.addAttribute("min", model.min);
        xmlStream.addAttribute("max", model.max);
        if (model.width) {
            xmlStream.addAttribute("width", model.width);
        }
        if (model.styleId) {
            xmlStream.addAttribute("style", model.styleId);
        }
        if (model.hidden) {
            xmlStream.addAttribute("hidden", "1");
        }
        if (model.bestFit) {
            xmlStream.addAttribute("bestFit", "1");
        }
        if (model.outlineLevel) {
            xmlStream.addAttribute("outlineLevel", model.outlineLevel);
        }
        if (model.collapsed) {
            xmlStream.addAttribute("collapsed", "1");
        }
        xmlStream.addAttribute("customWidth", "1");
        xmlStream.closeNode();
    }

    /** @returns `true` for `<col>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "col") {
            const { attributes } = node;
            const model: ColModel = (this.model = {
                min: clampColumn(attributes.min),
                max: clampColumn(attributes.max),
                width: attributes.width === undefined ? undefined : parseFloat(attributes.width || "0"),
            });
            if (attributes.style) {
                model.styleId = parseInt(attributes.style, 10);
            }
            if (utils.parseBoolean(attributes.hidden)) {
                model.hidden = true;
            }
            if (utils.parseBoolean(attributes.bestFit)) {
                model.bestFit = true;
            }
            if (attributes.outlineLevel) {
                model.outlineLevel = parseInt(attributes.outlineLevel, 10);
            }
            if (utils.parseBoolean(attributes.collapsed)) {
                model.collapsed = true;
            }
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }

    /**
     * Sets the style of the columns from `styleId` (`null` for an unknown id; nothing when the workbook has no styles
     * part).
     */
    override reconcile(model: ColModel, options: Pick<CellReconcileOptions, "styles">): void {
        // reconcile column styles
        if (model.styleId && options.styles) {
            model.style = options.styles.getStyleModel(model.styleId);
        }
    }
}

export default ColXform;
