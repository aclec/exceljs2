import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetViewModel } from "../../../types";
import colCache from "../../../utils/col-cache";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";
import { isOneOf } from "../../../utils/utils";

/** A pane of a sheet view (ECMA-376 ST_Pane). */
type Pane = NonNullable<SheetViewModel["activePane"]>;

const PANES: Readonly<Record<Pane, true>> = { topLeft: true, topRight: true, bottomLeft: true, bottomRight: true };

/** The presentations of ST_SheetViewType other than `normal` (the default, read as no style). */
const VIEW_STYLES: Readonly<Record<NonNullable<SheetViewModel["style"]>, true>> = { pageBreakPreview: true, pageLayout: true };

/** The attributes of `<sheetView>`, kept until the element is closed. */
interface SheetViewAttributes {
    workbookViewId: number;
    rightToLeft: boolean;
    tabSelected: boolean;
    showRuler: boolean;
    showRowColHeaders: boolean;
    showGridLines: boolean;
    zoomScale: number;
    zoomScaleNormal: number;
    style: SheetViewModel["style"];
}

/** The attributes of `<pane>`. */
interface PaneAttributes {
    xSplit: number;
    ySplit: number;
    topLeftCell: string | undefined;
    activePane: Pane;
    state: string | undefined;
}

/** The state of the view of a `<pane>` state (ST_PaneState): split when absent or unknown. */
function viewState(paneState: string | undefined): "frozen" | "split" {
    switch (paneState) {
        case "frozen":
        case "frozenSplit":
            return "frozen";
        case undefined:
        default:
            return "split";
    }
}

/**
 * A `<sheetView>` (an item of `<sheetViews>`), with its `<pane>` and `<selection>`: a normal, frozen or split view of a
 * worksheet. The model is flat (see SheetViewModel); a view without `<pane>` is normal.
 */
class SheetViewXform extends BaseXform<SheetViewModel> {
    /** While parsing: the attributes of `<sheetView>`. */
    declare sheetView: SheetViewAttributes | undefined;
    /** While parsing: the attributes of `<pane>`, if any. */
    declare pane: PaneAttributes | undefined;
    /** While parsing: the active cell of each `<selection>`, by pane (a pane named by the file). */
    selections = new Map<string, { pane: string; activeCell: string | undefined }>();

    get tag(): string {
        return "sheetView";
    }

    /** Forces a `state` other than frozen and split to `"normal"` (mutates the model). */
    override prepare(model: SheetViewModel): void {
        switch (model.state) {
            case "frozen":
            case "split":
                break;
            case "normal":
            case undefined:
            default:
                model.state = "normal";
                break;
        }
    }

    /** Writes the view; a split view whose `activePane` is `topLeft` (the default) loses it (mutates the model). */
    override render(xmlStream: XmlStream, model: SheetViewModel): void {
        xmlStream.openNode("sheetView", {
            workbookViewId: model.workbookViewId || 0,
        });
        if (model.rightToLeft === true) {
            xmlStream.addAttribute("rightToLeft", "1");
        }
        if (model.tabSelected) {
            xmlStream.addAttribute("tabSelected", "1");
        }
        if (model.showRuler === false) {
            xmlStream.addAttribute("showRuler", "0");
        }
        if (model.showRowColHeaders === false) {
            xmlStream.addAttribute("showRowColHeaders", "0");
        }
        if (model.showGridLines === false) {
            xmlStream.addAttribute("showGridLines", "0");
        }
        if (model.zoomScale) {
            xmlStream.addAttribute("zoomScale", model.zoomScale);
        }
        if (model.zoomScaleNormal) {
            xmlStream.addAttribute("zoomScaleNormal", model.zoomScaleNormal);
        }
        if (model.style) {
            xmlStream.addAttribute("view", model.style);
        }

        switch (model.state) {
            case "frozen": {
                const xSplit = model.xSplit || 0;
                const ySplit = model.ySplit || 0;
                const topLeftCell = model.topLeftCell || colCache.getAddress(ySplit + 1, xSplit + 1).address;
                const activePane: Pane = (model.xSplit && model.ySplit && "bottomRight") || (model.xSplit && "topRight") || "bottomLeft";

                xmlStream.leafNode("pane", {
                    xSplit: model.xSplit || undefined,
                    ySplit: model.ySplit || undefined,
                    topLeftCell,
                    activePane,
                    state: "frozen",
                });
                xmlStream.leafNode("selection", {
                    pane: activePane,
                    activeCell: model.activeCell,
                    sqref: model.activeCell,
                });
                break;
            }
            case "split":
                if (model.activePane === "topLeft") {
                    model.activePane = undefined;
                }
                xmlStream.leafNode("pane", {
                    xSplit: model.xSplit || undefined,
                    ySplit: model.ySplit || undefined,
                    topLeftCell: model.topLeftCell,
                    activePane: model.activePane,
                });
                xmlStream.leafNode("selection", {
                    pane: model.activePane,
                    activeCell: model.activeCell,
                    sqref: model.activeCell,
                });
                break;
            case "normal":
                if (model.activeCell) {
                    xmlStream.leafNode("selection", {
                        activeCell: model.activeCell,
                        sqref: model.activeCell,
                    });
                }
                break;
            case undefined:
            default:
                break;
        }
        xmlStream.closeNode();
    }

    /** @returns `true` for `<sheetView>`, `<pane>` and `<selection>` */
    override parseOpen(node: SaxTag): boolean {
        const { attributes } = node;
        switch (node.name) {
            case "sheetView":
                this.sheetView = {
                    workbookViewId: BaseXform.toIntValue(attributes.workbookViewId, NaN),
                    rightToLeft: parseBoolean(attributes.rightToLeft),
                    tabSelected: parseBoolean(attributes.tabSelected),
                    showRuler: parseBoolean(attributes.showRuler, true),
                    showRowColHeaders: parseBoolean(attributes.showRowColHeaders, true),
                    showGridLines: parseBoolean(attributes.showGridLines, true),
                    zoomScale: parseInt(attributes.zoomScale || "100", 10),
                    zoomScaleNormal: parseInt(attributes.zoomScaleNormal || "100", 10),
                    // ST_SheetViewType: normal (the default) or an unknown value is no style
                    style: isOneOf(VIEW_STYLES, attributes.view) ? attributes.view : undefined,
                };
                this.pane = undefined;
                this.selections = new Map();
                return true;

            case "pane":
                this.pane = {
                    xSplit: parseInt(attributes.xSplit || "0", 10),
                    ySplit: parseInt(attributes.ySplit || "0", 10),
                    topLeftCell: attributes.topLeftCell,
                    // ST_Pane: topLeft (the default) when absent or unknown
                    activePane: isOneOf(PANES, attributes.activePane) ? attributes.activePane : "topLeft",
                    state: attributes.state,
                };
                return true;

            case "selection": {
                const name = attributes.pane || "topLeft";
                this.selections.set(name, {
                    pane: name,
                    activeCell: attributes.activeCell,
                });
                return true;
            }

            default:
                return false;
        }
    }

    override parseText(): void {}

    /** @returns `false` at the end of `<sheetView>` (the model is then built), else `true` */
    override parseClose(name: string): boolean {
        switch (name) {
            case "sheetView": {
                // set by the start tag of this element
                const { sheetView, pane } = this;
                if (!sheetView) {
                    return false;
                }
                let model: SheetViewModel;
                let selection;
                if (pane) {
                    model = this.model = {
                        workbookViewId: sheetView.workbookViewId,
                        rightToLeft: sheetView.rightToLeft,
                        state: viewState(pane.state),
                        xSplit: pane.xSplit,
                        ySplit: pane.ySplit,
                        topLeftCell: pane.topLeftCell,
                        showRuler: sheetView.showRuler,
                        showRowColHeaders: sheetView.showRowColHeaders,
                        showGridLines: sheetView.showGridLines,
                        zoomScale: sheetView.zoomScale,
                        zoomScaleNormal: sheetView.zoomScaleNormal,
                    };
                    if (model.state === "split") {
                        model.activePane = pane.activePane;
                    }
                    selection = this.selections.get(pane.activePane);
                } else {
                    model = this.model = {
                        workbookViewId: sheetView.workbookViewId,
                        rightToLeft: sheetView.rightToLeft,
                        state: "normal",
                        showRuler: sheetView.showRuler,
                        showRowColHeaders: sheetView.showRowColHeaders,
                        showGridLines: sheetView.showGridLines,
                        zoomScale: sheetView.zoomScale,
                        zoomScaleNormal: sheetView.zoomScaleNormal,
                    };
                    selection = this.selections.get("topLeft");
                }
                if (selection && selection.activeCell) {
                    model.activeCell = selection.activeCell;
                }
                if (sheetView.style) {
                    model.style = sheetView.style;
                }
                return false;
            }
            default:
                return true;
        }
    }

    override reconcile(): void {}
}

export default SheetViewXform;
