import type { Fill } from "../../../../index";
import type { SaxInput, SaxTag } from "../../../utils/parse-sax";
import type { BorderModel, DxfModel, FontModel, NumFmtModel, StyleLookup, StyleModel, StyleRegistry, XfModel } from "../../../types";
import { ValueType } from "../../../doc/enums";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";

import BaseXform, { ownEntry } from "../base-xform";
import StaticXform from "../static-xform";
import ListXform from "../list-xform";
import FontXform from "./font-xform";
import FillXform, { type FillXformModel } from "./fill-xform";
import BorderXform from "./border-xform";
import NumFmtXform from "./numfmt-xform";
import StyleXform from "./style-xform";
import DxfXform from "./dxf-xform";

// custom numfmt ids start here
const NUMFMT_BASE = 164;

// the default (zero) font of a workbook
const DEFAULT_FONT: FontModel = { size: 11, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" };

/**
 * The model of the style manager role ({@link StylesXform.init}): the XML of each registered number format, font, fill,
 * border and cell format (deduplicated by their XML), in id order, and the differential formats.
 */
export interface StylesManagerModel {
    styles: string[];
    numFmts: string[];
    fonts: string[];
    borders: string[];
    fills: string[];
    dxfs: DxfModel[];
}

/**
 * The plain JSON model of a styles part: as parsed (a list is present only when it has items) or as rendered without the
 * style manager role (StylesXform.Mock, tests). `styles` are the cell formats (`cellXfs`).
 */
export interface StylesJsonModel {
    numFmts?: NumFmtModel[];
    fonts?: FontModel[];
    fills?: FillXformModel[];
    borders?: BorderModel[];
    styles?: XfModel[];
    dxfs?: DxfModel[];
}

/** The model of a {@link StylesXform}: its style manager state, or a plain JSON model. */
export type StylesModel = StylesManagerModel | StylesJsonModel;

/**
 * The index of the style manager role: the id of each registered part by its XML (number formats by format code).
 * The keys are built by the xforms (XML starting with `<`), the number format codes by the application: `numFmt` has no
 * prototype.
 */
export interface StylesManagerIndex {
    style: Record<string, number>;
    numFmt: Record<string, number>;
    /** Unused. */
    numFmtNextId: number;
    font: Record<string, number>;
    border: Record<string, number>;
    fill: Record<string, number>;
    /**
     * The id of the cell format of the styles that have only a number format (or nothing), by number format id: the
     * styles of most cells, registered without rendering their `<xf>` again.
     */
    numFmtStyle: Map<number | undefined, number>;
}

/**
 * The index of a parsed styles part: the style models built so far by {@link StylesXform.getStyleModel} (by cell format
 * id), and the custom number format codes by id.
 */
export interface StylesLookupIndex {
    model: (StyleModel | undefined)[];
    numFmt: (string | undefined)[];
}

// =============================================================================
/**
 * The styles part, `xl/styles.xml`, and the style manager of a workbook.
 *
 * - Writing (style manager role, `new StylesXform(true)` or {@link init}): the worksheets register the style of each
 *   cell, row and column with {@link addStyleModel} (and the differential formats with {@link addDxfStyle}); each part
 *   (number format, font, fill, border, cell format) is rendered to XML once and deduplicated by its XML. Then
 *   {@link render} (or `xml`) writes the part from these XML fragments.
 * - Reading: parsing builds a {@link StylesJsonModel}; {@link getStyleModel} and {@link getDxfStyle} give the style of a
 *   cell format id and of a differential format index.
 * - Without the style manager role, render writes a plain {@link StylesJsonModel}.
 */
class StylesXform extends BaseXform<StylesModel> implements StyleRegistry, StyleLookup {
    /** The namespaces of `<styleSheet>`. */
    static readonly STYLESHEET_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "x14ac x16r2",
        "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "xmlns:x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
    };

    /** The fixed parts of the styles part. */
    static readonly STATIC_XFORMS = {
        cellStyles: new StaticXform({
            tag: "cellStyles",
            $: { count: 1 },
            c: [{ tag: "cellStyle", $: { name: "Normal", xfId: 0, builtinId: 0 } }],
        }),
        dxfs: new StaticXform({ tag: "dxfs", $: { count: 0 } }),
        tableStyles: new StaticXform({
            tag: "tableStyles",
            $: { count: 0, defaultTableStyle: "TableStyleMedium2", defaultPivotStyle: "PivotStyleLight16" },
        }),
        extLst: new StaticXform({
            tag: "extLst",
            c: [
                {
                    tag: "ext",
                    $: {
                        uri: "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
                        "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                    },
                    c: [{ tag: "x14:slicerStyles", $: { defaultSlicerStyle: "SlicerStyleLight1" } }],
                },
                {
                    tag: "ext",
                    $: {
                        uri: "{9260A510-F301-46a8-8635-F512D64BE5F5}",
                        "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
                    },
                    c: [{ tag: "x15:timelineStyles", $: { defaultTimelineStyle: "TimeSlicerStyleLight1" } }],
                },
            ],
        }),
    };

    /** The style manager that writes no styles (assigned below the class). */
    declare static Mock: typeof StylesXformMock;

    declare map: {
        readonly numFmts: ListXform<NumFmtModel>;
        readonly fonts: ListXform<FontModel>;
        readonly fills: ListXform<FillXformModel>;
        readonly borders: ListXform<BorderModel>;
        readonly cellStyleXfs: ListXform<XfModel>;
        readonly cellXfs: ListXform<XfModel>;
        readonly dxfs: ListXform<DxfModel>;

        // for style manager
        readonly numFmt: NumFmtXform;
        readonly font: FontXform;
        readonly fill: FillXform;
        readonly border: BorderXform;
        readonly style: StyleXform;

        readonly cellStyles: StaticXform;
        readonly tableStyles: StaticXform;
        readonly extLst: StaticXform;
    };

    /** The index of the style manager role (set by {@link init}), or of a parsed part (set by the parse). */
    declare index?: StylesManagerIndex | StylesLookupIndex;

    /** Style manager role: the cell format id of each style object already registered. */
    declare weakMap?: WeakMap<StyleModel, number>;

    /** @param initialise - take the style manager role ({@link init}) */
    constructor(initialise?: boolean) {
        super();

        this.map = {
            numFmts: new ListXform({ tag: "numFmts", count: true, childXform: new NumFmtXform() }),
            fonts: new ListXform({
                tag: "fonts",
                count: true,
                childXform: new FontXform(),
                $: { "x14ac:knownFonts": 1 },
            }),
            fills: new ListXform({ tag: "fills", count: true, childXform: new FillXform() }),
            borders: new ListXform({ tag: "borders", count: true, childXform: new BorderXform() }),
            cellStyleXfs: new ListXform({ tag: "cellStyleXfs", count: true, childXform: new StyleXform() }),
            cellXfs: new ListXform({
                tag: "cellXfs",
                count: true,
                childXform: new StyleXform({ xfId: true }),
            }),
            dxfs: new ListXform({ tag: "dxfs", always: true, count: true, childXform: new DxfXform() }),

            // for style manager
            numFmt: new NumFmtXform(),
            font: new FontXform(),
            fill: new FillXform(),
            border: new BorderXform(),
            style: new StyleXform({ xfId: true }),

            cellStyles: StylesXform.STATIC_XFORMS.cellStyles,
            tableStyles: StylesXform.STATIC_XFORMS.tableStyles,
            extLst: StylesXform.STATIC_XFORMS.extLst,
        };

        if (initialise) {
            // StylesXform also acts as style manager and is used to build up styles-model during worksheet processing
            this.init();
        }
    }

    /** Resets the index of the style manager role. */
    initIndex(): void {
        this.index = {
            style: {},
            numFmt: Object.create(null) as Record<string, number>, // by format code, which may be "constructor"
            numFmtNextId: 164, // start custom format ids here
            font: {},
            border: {},
            fill: {},
            numFmtStyle: new Map(),
        };
    }

    /** Takes the style manager role: an empty model with the default border, cell format and fills. */
    init(): void {
        // Prepare for Style Manager role
        this.model = {
            styles: [],
            numFmts: [],
            fonts: [],
            borders: [],
            fills: [],
            dxfs: [],
        };

        this.initIndex();

        // default (zero) border
        this._addBorder({});

        // add default (all zero) style
        this._addStyle({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });

        // add default fills
        this._addFill({ type: "pattern", pattern: "none" });
        this._addFill({ type: "pattern", pattern: "gray125" });

        this.weakMap = new WeakMap();
    }

    /** The model in the style manager role (built by {@link init}). */
    private get managerModel(): StylesManagerModel {
        return this.model as StylesManagerModel;
    }

    /** The index in the style manager role (built by {@link init}). */
    private get managerIndex(): StylesManagerIndex {
        return this.index as StylesManagerIndex;
    }

    /**
     * Writes the styles part.
     *
     * @param model - default: the model of this xform. In the style manager role (or once parsed), it is read as a
     *   {@link StylesManagerModel}; otherwise as a {@link StylesJsonModel}.
     */
    override render(xmlStream: XmlStream, model?: StylesModel | null): void {
        const styles = model || this.model;
        //
        //   <fonts count="2" x14ac:knownFonts="1">
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("styleSheet", StylesXform.STYLESHEET_ATTRIBUTES);

        if (this.index) {
            // model has been built by style manager role (contains xml)
            const manager = styles as StylesManagerModel;
            if (manager.numFmts && manager.numFmts.length) {
                xmlStream.openNode("numFmts", { count: manager.numFmts.length });
                writeAllXml(xmlStream, manager.numFmts);
                xmlStream.closeNode();
            }

            if (!manager.fonts.length) {
                // default (zero) font
                this._addFont(DEFAULT_FONT);
            }
            xmlStream.openNode("fonts", { count: manager.fonts.length, "x14ac:knownFonts": 1 });
            writeAllXml(xmlStream, manager.fonts);
            xmlStream.closeNode();

            xmlStream.openNode("fills", { count: manager.fills.length });
            writeAllXml(xmlStream, manager.fills);
            xmlStream.closeNode();

            xmlStream.openNode("borders", { count: manager.borders.length });
            writeAllXml(xmlStream, manager.borders);
            xmlStream.closeNode();

            this.map.cellStyleXfs.render(xmlStream, [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }]);

            xmlStream.openNode("cellXfs", { count: manager.styles.length });
            writeAllXml(xmlStream, manager.styles);
            xmlStream.closeNode();
        } else {
            // model is plain JSON and needs to be xformed
            const json = styles as StylesJsonModel;
            this.map.numFmts.render(xmlStream, json.numFmts);
            this.map.fonts.render(xmlStream, json.fonts);
            this.map.fills.render(xmlStream, json.fills);
            this.map.borders.render(xmlStream, json.borders);
            this.map.cellStyleXfs.render(xmlStream, [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }]);
            this.map.cellXfs.render(xmlStream, json.styles);
        }

        StylesXform.STATIC_XFORMS.cellStyles.render(xmlStream);

        this.map.dxfs.render(xmlStream, styles!.dxfs);

        StylesXform.STATIC_XFORMS.tableStyles.render(xmlStream);
        StylesXform.STATIC_XFORMS.extLst.render(xmlStream);

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the part is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "styleSheet":
                this.initIndex();
                return true;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                return true;
            }
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "styleSheet": {
                const model: StylesJsonModel = {};
                // the lists that have items
                const { numFmts, fonts, fills, borders, cellXfs, dxfs } = this.map;
                if (numFmts.model && numFmts.model.length) model.numFmts = numFmts.model;
                if (fonts.model && fonts.model.length) model.fonts = fonts.model;
                if (fills.model && fills.model.length) model.fills = fills.model;
                if (borders.model && borders.model.length) model.borders = borders.model;
                if (cellXfs.model && cellXfs.model.length) model.styles = cellXfs.model;
                if (dxfs.model && dxfs.model.length) model.dxfs = dxfs.model;
                this.model = model;

                // index numFmts
                const index: StylesLookupIndex = {
                    model: [],
                    numFmt: [],
                };
                if (model.numFmts) {
                    const numFmtIndex = index.numFmt;
                    for (const numFmt of model.numFmts) {
                        numFmtIndex[numFmt.id] = numFmt.formatCode;
                    }
                }
                this.index = index;

                return false;
            }
            default:
                // not quite sure how we get here!
                return true;
        }
    }

    /**
     * Registers the style of a cell, row or column (style manager role): each part is processed and cross-referenced.
     * A style object already registered gets the same id.
     *
     * @param cellType - used when the style has no number format: Number cells get `General`, Date cells `mm-dd-yy`
     *   (default: Number)
     * @returns the cell format id (the `s` attribute); 0 for a falsy style
     */
    addStyleModel(model: StyleModel, cellType?: ValueType): number {
        if (!model) {
            return 0;
        }

        // if we have no default font, add it here now
        if (!this.managerModel.fonts.length) {
            // default (zero) font
            this._addFont(DEFAULT_FONT);
        }

        // if we have seen this style object before, assume it has the same styleId
        const { weakMap } = this;
        if (weakMap) {
            const id = weakMap.get(model);
            if (id !== undefined) {
                return id;
            }
        }

        const style: XfModel = {};
        cellType = cellType || ValueType.Number;

        if (model.numFmt) {
            style.numFmtId = this._addNumFmtStr(model.numFmt);
        } else if (cellType === ValueType.Number) {
            style.numFmtId = this._addNumFmtStr("General");
        } else if (cellType === ValueType.Date) {
            style.numFmtId = this._addNumFmtStr("mm-dd-yy");
        }

        const { font, border, fill, alignment, protection } = model;
        if (!font && !border && !fill && !alignment && !protection) {
            // the format of a style with only a number format depends on it only: skip rendering its <xf>
            const index = this.managerIndex.numFmtStyle;
            let styleId = index.get(style.numFmtId);
            if (styleId === undefined) {
                styleId = this._addStyle(style);
                index.set(style.numFmtId, styleId);
            }
            if (weakMap) {
                weakMap.set(model, styleId);
            }
            return styleId;
        }

        if (font) {
            style.fontId = this._addFont(font);
        }

        if (border) {
            style.borderId = this._addBorder(border);
        }

        if (fill) {
            style.fillId = this._addFill(fill);
        }

        if (alignment) {
            style.alignment = alignment;
        }

        if (protection) {
            style.protection = protection;
        }

        const styleId = this._addStyle(style);
        if (weakMap) {
            weakMap.set(model, styleId);
        }
        return styleId;
    }

    /**
     * The style of a cell format id (`s="n"`) of a parsed part. Built once per id: the style objects (and their font,
     * border and fill) are shared between the cells of a style.
     *
     * @returns `null` when the id is not a cell format of the part (or the part was not parsed)
     */
    getStyleModel(id: number): StyleModel | null {
        // only a parsed part has a lookup index (the style manager role has its own index)
        const { index } = this;
        if (!index || !("model" in index)) return null;
        // the model of a parsed part is its JSON model
        const json = this.model as StylesJsonModel;

        // if the style doesn't exist return null
        const style = json.styles?.[id];
        if (!style) return null;

        // have we built this model before?
        const built = index.model[id];
        if (built) return built;

        // build a new model
        const result: StyleModel = (index.model[id] = {});

        // -------------------------------------------------------
        // number format
        if (style.numFmtId) {
            const numFmt = index.numFmt[style.numFmtId] || NumFmtXform.getDefaultFmtCode(style.numFmtId);
            if (numFmt) {
                result.numFmt = numFmt;
            }
        }

        const font = part(json.fonts, style.fontId);
        if (font) result.font = font;
        const border = part(json.borders, style.borderId);
        if (border) result.border = border;
        const fill = part(json.fills, style.fillId);
        // a parsed fill may lack the parts that index.d.ts Fill requires (see FillModel)
        if (fill) result.fill = fill as Fill;

        // -------------------------------------------------------
        // alignment
        if (style.alignment) {
            result.alignment = style.alignment;
        }

        // -------------------------------------------------------
        // protection
        if (style.protection) {
            result.protection = style.protection;
        }

        return result;
    }

    /**
     * Registers a differential format (style manager role, also in the Mock): adds `numFmtId` to `style` when it has a
     * `numFmt` (mutation).
     *
     * @returns its index in `dxfs`
     */
    addDxfStyle(style: DxfModel): number {
        if (style.numFmt) {
            // register numFmtId to use it during dxf-xform rendering
            style.numFmtId = this._addNumFmtStr(style.numFmt);
        }

        // the model of the style manager role and the Mock's have a dxfs list
        const { dxfs } = this.model as { dxfs: DxfModel[] };
        dxfs.push(style);
        return dxfs.length - 1;
    }

    /** @returns the differential format at an index of a parsed part, `undefined` when there is none */
    getDxfStyle(id: number): DxfModel | undefined {
        // a parsed part has no dxfs when its <dxfs> is absent or empty
        return this.model?.dxfs?.[id];
    }

    // =========================================================================
    // Private Interface
    _addStyle(style: XfModel): number {
        const xml = this.map.style.toXml(style);
        const index = this.managerIndex.style;
        let id = index[xml];
        if (id === undefined) {
            const { styles } = this.managerModel;
            id = index[xml] = styles.length;
            styles.push(xml);
        }
        return id;
    }

    // =========================================================================
    // Number Formats
    /** @returns the id of a format code: built-in, or registered as a custom format */
    _addNumFmtStr(formatCode: string): number {
        // check if default format
        const defaultId = NumFmtXform.getDefaultFmtId(formatCode);
        if (defaultId !== undefined) return defaultId;

        // check if already in
        const index = this.managerIndex.numFmt;
        let id = index[formatCode];
        if (id !== undefined) return id;

        const { numFmts } = this.managerModel;
        id = index[formatCode] = NUMFMT_BASE + numFmts.length;
        const xml = this.map.numFmt.toXml({ id, formatCode });
        numFmts.push(xml);
        return id;
    }

    // =========================================================================
    // Fonts
    _addFont(font: FontModel): number {
        const xml = this.map.font.toXml(font);
        const index = this.managerIndex.font;
        let id = index[xml];
        if (id === undefined) {
            const { fonts } = this.managerModel;
            id = index[xml] = fonts.length;
            fonts.push(xml);
        }
        return id;
    }

    // =========================================================================
    // Borders
    _addBorder(border: BorderModel): number {
        const xml = this.map.border.toXml(border);
        const index = this.managerIndex.border;
        let id = index[xml];
        if (id === undefined) {
            const { borders } = this.managerModel;
            id = index[xml] = borders.length;
            borders.push(xml);
        }
        return id;
    }

    // =========================================================================
    // Fills
    _addFill(fill: FillXformModel): number {
        const xml = this.map.fill.toXml(fill);
        const index = this.managerIndex.fill;
        let id = index[xml];
        if (id === undefined) {
            const { fills } = this.managerModel;
            id = index[xml] = fills.length;
            fills.push(xml);
        }
        return id;
    }

    // =========================================================================
}

// writes XML fragments built by the xforms
function writeAllXml(xmlStream: XmlStream, fragments: readonly string[]): void {
    for (let i = 0; i < fragments.length; i++) {
        xmlStream.writeXml(fragments[i]!);
    }
}

// the part of a parsed list at an id, if any (ids come from the file)
function part<T>(group: readonly T[] | undefined, id: number | undefined): T | undefined {
    if (group && (id || id === 0)) {
        return group[id] || undefined;
    }
    return undefined;
}

/** The model of {@link StylesXformMock}: plain JSON, every list present. */
type StylesMockModel = Required<StylesJsonModel>;

/**
 * The style manager used when styles are not written (`useStyles: false`): every cell gets the default style (0), except
 * dates, which get one date style; differential formats are still registered (conditional formatting rules refer to
 * them by index). It renders a plain JSON model, and does not parse.
 */
class StylesXformMock extends StylesXform {
    declare model: StylesMockModel | null | undefined;

    /** The cell format id of the date style, once registered. */
    declare private _dateStyleId?: number;

    /** @param initialise - ignored */
    constructor(initialise?: boolean) {
        super();

        this.model = {
            styles: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }],
            numFmts: [],
            fonts: [{ size: 11, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" }],
            borders: [{}],
            fills: [
                { type: "pattern", pattern: "none" },
                { type: "pattern", pattern: "gray125" },
            ],
            // conditional formatting rules reference their differential formats by index: these are
            // written even without useStyles
            dxfs: [],
        };
    }

    // =========================================================================
    // Style Manager Interface

    /** Consumes and discards the stream (override of the normal behaviour). */
    override parseStream(stream: SaxInput & { autodrain(): void }): Promise<undefined> {
        stream.autodrain();
        return Promise.resolve(undefined);
    }

    /** @returns the date style for a Date cell, else 0 */
    override addStyleModel(model: StyleModel, cellType?: ValueType): number {
        return cellType === ValueType.Date ? this.dateStyleId : 0;
    }

    /** The cell format id of the date style (registered on first use). */
    get dateStyleId(): number {
        if (!this._dateStyleId) {
            const dateStyle: XfModel = {
                numFmtId: NumFmtXform.getDefaultFmtId("mm-dd-yy"),
            };
            // the mock has its model from its constructor
            const { styles } = this.model!;
            this._dateStyleId = styles.length;
            styles.push(dateStyle);
        }
        return this._dateStyleId;
    }

    /** @returns `{}`: the mock has no styles */
    override getStyleModel(id: number): StyleModel {
        return {};
    }

    // number formats of differential formats (see addDxfStyle). The mock has no index, so its
    // model holds plain numFmt models ({id, formatCode}), rendered by the numFmts list xform
    override _addNumFmtStr(formatCode: string): number {
        const defaultId = NumFmtXform.getDefaultFmtId(formatCode);
        if (defaultId !== undefined) return defaultId;

        const { numFmts } = this.model!;
        let numFmt = numFmts.find((candidate) => candidate.formatCode === formatCode);
        if (!numFmt) {
            numFmt = { id: NUMFMT_BASE + numFmts.length, formatCode };
            numFmts.push(numFmt);
        }
        return numFmt.id;
    }
}

StylesXform.Mock = StylesXformMock;

export default StylesXform;
