import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CellPrepareOptions, CellReconcileOptions, ParsedCellModel, RowModel } from "../../../types";
import BaseXform from "../base-xform";
import utils from "../../../utils/utils";
import colCache from "../../../utils/col-cache";

import CellXform from "./cell-xform";

/** The options of a {@link RowXform}. */
export interface RowXformOptions {
    /** Parsing fails when a row has more cells than this (a limit against hostile files). Default: no limit. */
    readonly maxItems?: number;
}

/**
 * A row, `<row r="1" spans="1:3" ht="20" …>` with its cells (an item of `<sheetData>`). Hot: it runs for every row
 * written and read, and delegates every cell to its CellXform.
 *
 * Parsing gives a `RowModel<ParsedCellModel>`: `number` is `NaN` when `r` is missing (the worksheet infers it), `min` /
 * `max` come from `spans`.
 */
class RowXform extends BaseXform<RowModel<ParsedCellModel> | RowModel> {
    /** The parsed row (prepare and render take the document `RowModel<CellModel>`: see the type parameter). */
    declare model: RowModel<ParsedCellModel> | null | undefined;

    /** See {@link RowXformOptions.maxItems}. */
    readonly maxItems: number | undefined;

    declare map: { readonly c: CellXform };

    constructor(options?: RowXformOptions) {
        super();

        this.maxItems = options && options.maxItems;
        this.map = {
            c: new CellXform(),
        };
    }

    get tag(): string {
        return "row";
    }

    /** Registers the style of the row (`styleId`, when not the default) and prepares its cells. */
    override prepare(model: RowModel, options: CellPrepareOptions): void {
        const styleId = options.styles.addStyleModel(model.style || {});
        if (styleId) {
            model.styleId = styleId;
        }
        const cellXform = this.map.c;
        const { cells } = model;
        for (let i = 0; i < cells.length; i++) {
            // a hole of a sparse array is undefined
            const cell = cells[i];
            if (cell !== undefined) {
                cellXform.prepare(cell, options);
            }
        }
    }

    /** Writes the row and its cells (ListXform passes the index of the row as third argument: not used). */
    override render(xmlStream: XmlStream, model: RowModel): void {
        xmlStream.openNode("row");
        xmlStream.addAttribute("r", model.number);
        if (model.height) {
            xmlStream.addAttribute("ht", model.height);
            xmlStream.addAttribute("customHeight", "1");
        }
        if (model.hidden) {
            xmlStream.addAttribute("hidden", "1");
        }
        const { min, max } = model;
        if (min !== undefined && max !== undefined && min > 0 && max > 0 && min <= max) {
            xmlStream.addAttribute("spans", `${min}:${max}`);
        }
        if (model.styleId) {
            xmlStream.addAttribute("s", model.styleId);
            xmlStream.addAttribute("customFormat", "1");
        }
        xmlStream.addAttribute("x14ac:dyDescent", "0.25");
        if (model.outlineLevel) {
            xmlStream.addAttribute("outlineLevel", model.outlineLevel);
        }
        if (model.collapsed) {
            xmlStream.addAttribute("collapsed", "1");
        }

        const cellXform = this.map.c;
        const { cells } = model;
        for (let i = 0; i < cells.length; i++) {
            // a hole of a sparse array is undefined
            const cell = cells[i];
            if (cell !== undefined) {
                cellXform.render(xmlStream, cell);
            }
        }

        xmlStream.closeNode();
    }

    /** @returns `true` for `<row>` and its cells */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (node.name === "row") {
            const { attributes } = node;
            const spans = attributes.spans ? attributes.spans.split(":").map((span) => parseInt(span, 10)) : [undefined, undefined];
            // through this.model, not a local alias of the new object (faster objects, see CellXform.parseOpen)
            this.model = {
                number: RowXform.parseRowNumber(attributes.r),
                min: spans[0],
                max: spans[1],
                cells: [],
            };
            if (attributes.s) {
                this.model.styleId = parseInt(attributes.s, 10);
            }
            if (utils.parseBoolean(attributes.hidden)) {
                this.model.hidden = true;
            }
            if (utils.parseBoolean(attributes.bestFit)) {
                this.model.bestFit = true;
            }
            if (attributes.ht) {
                this.model.height = parseFloat(attributes.ht);
            }
            if (attributes.outlineLevel) {
                this.model.outlineLevel = parseInt(attributes.outlineLevel, 10);
            }
            if (utils.parseBoolean(attributes.collapsed)) {
                this.model.collapsed = true;
            }
            return true;
        }

        // the only child: a direct test (never an inherited member of the map), faster than a map lookup for every cell
        if (node.name === "c") {
            this.parser = this.map.c;
            this.parser.parseOpen(node);
            return true;
        }
        return false;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /**
     * @returns `false` at the end of `<row>`, else `true`
     * @throws Error when the row has more than `maxItems` cells
     */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the row was opened before its cells; a cell model is set by its start tag
                const row = this.model;
                const cell = this.map.c.model;
                if (row && cell) {
                    const { cells } = row;
                    cells.push(cell);
                    if (this.maxItems && cells.length > this.maxItems) {
                        throw new Error(`Max column count (${this.maxItems}) exceeded`);
                    }
                }
                this.parser = undefined;
            }
            return true;
        }
        return false;
    }

    /**
     * Sets the style of the row from `styleId` (`{}` without one, `null` for an unknown id; nothing when the workbook has
     * no styles part) and reconciles its cells (mutating them into CellModels).
     */
    override reconcile(model: RowModel<ParsedCellModel>, options: CellReconcileOptions): void {
        model.style = model.styleId ? options.styles && options.styles.getStyleModel(model.styleId) : {};
        if (model.styleId !== undefined) {
            model.styleId = undefined;
        }

        const cellXform = this.map.c;
        const { cells } = model;
        for (let i = 0; i < cells.length; i++) {
            // a hole of a sparse array is undefined
            const cell = cells[i];
            if (cell !== undefined) {
                cellXform.reconcile(cell, options);
            }
        }
    }

    /**
     * The `r` attribute of `<row>` is optional: `NaN` when it is missing (the row then follows the previous one). A row
     * number outside the grid is rejected: even an empty `<row r="4294967294"/>` made every row traversal take minutes.
     *
     * @param r - the `r` attribute
     * @returns the 1-based row number, or `NaN`
     * @throws Error when the row number is outside 1..{@link colCache.MAX_ROWS}
     */
    static parseRowNumber(this: void, r: string | undefined): number {
        const number = BaseXform.toIntValue(r, NaN);
        if (number < 1 || number > colCache.MAX_ROWS) {
            throw new Error(`Out of bounds. Excel supports rows from 1 to ${colCache.MAX_ROWS}: <row r="${String(r)}">`);
        }
        return number;
    }
}

export default RowXform;
