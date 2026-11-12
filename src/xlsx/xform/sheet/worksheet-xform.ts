import type { SaxTag } from "../../../utils/parse-sax";
import type {
    ColModel,
    ConditionalFormattingModel,
    CellPrepareOptions,
    CellReconcileOptions,
    ConditionalFormattingRuleModel,
    DrawingAnchorModel,
    DrawingModel,
    FormulaCellModel,
    MediaModel,
    NoteModel,
    PageSetupModel,
    ParsedCellModel,
    ParsedWorksheetModel,
    RelationshipModel,
    RowModel,
    SheetHyperlinkModel,
    SheetPartRefModel,
    SheetViewModel,
    TableModel,
    WorkbookPrepareOptions,
    WorkbookReconcileOptions,
    WorksheetLoadBudget,
    WorksheetModel,
    WorksheetPrepareOptions,
    WorksheetPropertiesModel,
    WorksheetReconcileOptions,
} from "../../../types";

import colCache from "../../../utils/col-cache";
import Range from "../../../doc/range";
import { ValueType } from "../../../doc/enums";
import { DEFAULT_MAX_MERGED_CELLS, checkMergedCells } from "../../../utils/merge-limits";
import XmlStream from "../../../utils/xml-stream";
import { runSteps } from "../../../utils/pacer";

import RelType from "../../rel-type";

import Merges from "./merges";

import BaseXform, { ownEntry } from "../base-xform";
import ListXform from "../list-xform";
import RowXform from "./row-xform";
import ColXform from "./col-xform";
import DimensionXform from "./dimension-xform";
import HyperlinkXform from "./hyperlink-xform";
import MergeCellXform from "./merge-cell-xform";
import DataValidationsXform from "./data-validations-xform";
import SheetPropertiesXform, { type SheetPropertiesModel } from "./sheet-properties-xform";
import SheetFormatPropertiesXform, { type SheetFormatPropertiesModel } from "./sheet-format-properties-xform";
import SheetViewXform from "./sheet-view-xform";
import SheetProtectionXform from "./sheet-protection-xform";
import PageMarginsXform from "./page-margins-xform";
import PageSetupXform from "./page-setup-xform";
import PrintOptionsXform, { type PrintOptionsModel } from "./print-options-xform";
import AutoFilterXform from "./auto-filter-xform";
import PictureXform from "./picture-xform";
import DrawingXform from "./drawing-xform";
import TablePartXform from "./table-part-xform";
import RowBreaksXform from "./row-breaks-xform";
import HeaderFooterXform, { defaultHeaderFooter } from "./header-footer-xform";
import ConditionalFormattingsXform from "./cf/conditional-formattings-xform";
import CfRuleXform from "./cf/cf-rule-xform";
import DatabarExtXform from "./cf-ext/databar-ext-xform";
import ExtListXform from "./ext-lst-xform";
import type { ConditionalFormattingExtModel } from "./cf-ext/conditional-formatting-ext-xform";

/** The steps of prepare / render / reconcile: one per row (the caller may pause between them). */
type Steps = Generator<void, void, undefined>;

/** `<legacyDrawing r:id="…"/>`: the VML drawing of the notes of a worksheet. */
class LegacyDrawingXform extends DrawingXform {
    override get tag(): string {
        return "legacyDrawing";
    }
}

/**
 * An empty object without prototype, for keys that come from the file (cell and part references, ids): a lookup never
 * finds an inherited member, and a "__proto__" key is an ordinary key. Object.create(null), not setPrototypeOf on a
 * literal: its lookups (once per cell) are faster.
 */
function emptyRecord<T>(): Record<string, T> {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Object.create is typed any: an empty object (without prototype) is a Record of any value type
    return Object.create(null) as Record<string, T>;
}

/**
 * The own entry of an object keyed by data of the file (a part name, a media name): never an inherited member.
 *
 * @param key - converted to a string, as for an object key (a missing key is `"undefined"`)
 */
function getOwn<T>(object: Readonly<Record<string, T>> | null | undefined, key: string | undefined): T | undefined {
    const name = String(key);
    return object && Object.hasOwn(object, name) ? object[name] : undefined;
}

// Parts are keyed by their zip entry name. A worksheet relationship target is relative to
// xl/worksheets/ or absolute from the package root (openpyxl writes /xl/comments/comment1.xml).
function getPart<T>(parts: Readonly<Record<string, T>> | undefined, rel: RelationshipModel | undefined): T | undefined {
    if (!rel || !rel.Target || rel.TargetMode === "External") {
        return undefined;
    }
    const segments = rel.Target.startsWith("/") ? [] : ["xl", "worksheets"];
    for (const segment of rel.Target.split("/")) {
        if (segment === "..") {
            segments.pop();
        } else if (segment && segment !== ".") {
            segments.push(segment);
        }
    }
    return getOwn(parts, segments.join("/"));
}

// A <hyperlink> has a relationship (external target), a location (place in the workbook, or
// anchor in the external target), or both. exceljs writes internal links with both, the
// target being the location itself ("Sheet2!A1" or "#'Sheet2'!A1"): don't append it twice.
function getHyperlinkTarget(relTarget: string | undefined, location: string | undefined): string | undefined {
    if (!location) {
        return relTarget;
    }
    const anchor = location.startsWith("#") ? location : `#${location}`;
    if (!relTarget) {
        return anchor;
    }
    if (relTarget === location || relTarget.includes("#")) {
        return relTarget;
    }
    return `${relTarget}${anchor}`;
}

// The r attribute of <row> and <c> is optional (ECMA-376 CT_Row, CT_Cell): a row without it
// follows the previous row, a cell without it follows the previous cell of its row.
function inferMissingReferences(row: RowModel<ParsedCellModel>, previousRowNumber: number): void {
    if (!(row.number > 0)) {
        const addressed = row.cells.find((cell) => cell.address);
        row.number = addressed && addressed.address ? colCache.decodeAddress(addressed.address).row : previousRowNumber + 1;
    }
    let previousAddress: string | undefined;
    for (const cell of row.cells) {
        if (!cell.address) {
            const previousCol = previousAddress ? colCache.decodeAddress(previousAddress).col : 0;
            cell.address = colCache.encodeAddress(row.number, previousCol + 1);
        }
        previousAddress = cell.address;
    }
}

const cellRefRegex = /^([A-Z]{1,3})([1-9]\d{0,6})$/;

// Excel writes no <c> element for an empty cell without a style, even when it has a note: create the cells (and rows)
// that notes refer to, so that the notes are not lost. References outside the grid are ignored.
function addNoteCells(rows: RowModel<ParsedCellModel>[], comments: readonly NoteModel[] | undefined): void {
    const addresses = new Set<string | undefined>();
    for (const row of rows) {
        for (const cell of row.cells) {
            addresses.add(cell.address);
        }
    }
    const rowsByNumber = new Map(rows.map((row) => [row.number, row]));
    for (const { ref } of comments || []) {
        const match = typeof ref === "string" && !addresses.has(ref) && cellRefRegex.exec(ref);
        if (!match) {
            continue;
        }
        const [, letters = "", digits = ""] = match;
        if ((letters.length === 3 && letters > "XFD") || Number(digits) > colCache.MAX_ROWS) {
            continue;
        }
        const number = Number(digits);
        let row = rowsByNumber.get(number);
        if (!row) {
            row = { number, cells: [] };
            rows.push(row);
            rowsByNumber.set(number, row);
        }
        row.cells.push({ address: ref, type: ValueType.Null });
        addresses.add(ref);
    }
}

// Every cell of a merge is materialised when a worksheet is loaded: the merges of the worksheet are checked against the caps
// of utils/merge-limits before any cell is created. The load budget is shared by all the worksheets of the workbook.
function spendMergedCells(loadBudget: WorksheetLoadBudget, mergeCells: readonly string[] | null | undefined): void {
    let sheetCount = 0;
    for (const ref of mergeCells || []) {
        const { count } = new Range(ref);
        sheetCount += count;
        loadBudget.mergedCells += count;
        checkMergedCells(`Cannot merge ${ref}`, sheetCount, loadBudget.mergedCells, loadBudget.maxMergedCells);
    }
}

/** Completes a rule with the settings of its x14 version (of the same type) that it lacks. */
function fillUndefined<T extends ConditionalFormattingRuleModel>(rule: T, extRule: T, keys: readonly (keyof T)[]): void {
    for (const key of keys) {
        if (rule[key] === undefined && extRule[key] !== undefined) {
            rule[key] = extRule[key];
        }
    }
}

/** The settings of every rule that its x14 version may give. */
const COMMON_KEYS = ["priority", "stopIfTrue", "dxfId", "x14Id", "formulae", "style"] as const;

/**
 * Merges the x14 version of a rule into the main rule: the settings that the main rule lacks (by type: those its xforms
 * read). A version of another type (malformed) is ignored.
 */
const mergeRule = (rule: ConditionalFormattingRuleModel, extRule: ConditionalFormattingRuleModel): void => {
    switch (rule.type) {
        case "expression":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "x14"]);
            }
            break;
        case "cellIs":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "operator", "x14"]);
            }
            break;
        case "top10":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "rank", "percent", "bottom"]);
            }
            break;
        case "aboveAverage":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "aboveAverage"]);
            }
            break;
        case "colorScale":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "cfvo", "color"]);
            }
            break;
        case "iconSet":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "iconSet", "reverse", "showValue", "custom", "cfvo", "icons"]);
            }
            break;
        case "containsText":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "operator", "text", "x14"]);
            }
            break;
        case "timePeriod":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, "timePeriod", "x14"]);
            }
            break;
        case "dataBar":
            if (extRule.type === rule.type) {
                fillUndefined(rule, extRule, [...COMMON_KEYS, ...DatabarExtXform.X14_SETTINGS, "gradient", "showValue", "cfvo", "color"]);
                // Excel 2010+ reads the thresholds of a data bar from its extension, where they may be of the
                // types autoMin / autoMax, unknown to the main <cfvo> (which has min / max instead)
                if (extRule.cfvo && extRule.cfvo.length) {
                    rule.cfvo = extRule.cfvo;
                }
            }
            break;
    }
};

const mergeConditionalFormattings = (
    model: ConditionalFormattingModel[],
    extModel: readonly ConditionalFormattingExtModel[] | null | undefined
): ConditionalFormattingModel[] => {
    // conditional formattings are rendered in worksheet.conditionalFormatting and also in
    // worksheet.extLst.ext.x14:conditionalFormattings
    // some (e.g. dataBar) are even spread across both!
    if (!extModel || !extModel.length) {
        return model;
    }
    const cfs = model || [];

    // index model rules by x14Id. Both keys come from the file: a Map, so that an id such as "__proto__" is an ordinary key
    const cfMap = new Map<string | undefined, ConditionalFormattingModel>();
    const ruleMap = new Map<string, ConditionalFormattingRuleModel>();
    for (const cf of cfs) {
        cfMap.set(cf.ref, cf);
        for (const rule of cf.rules) {
            const { x14Id } = rule;
            if (x14Id) {
                ruleMap.set(x14Id, rule);
            }
        }
    }

    for (const extCf of extModel) {
        const extOnlyRules: ConditionalFormattingRuleModel[] = [];
        for (const extRule of extCf.rules) {
            const rule = extRule.x14Id && ruleMap.get(extRule.x14Id);
            if (rule) {
                // merge with matching rule
                mergeRule(rule, extRule);
                continue;
            }
            // a formula rule found only in the extension (e.g. one referencing another sheet) is
            // written back there
            if (CfRuleXform.isFormulaRule(extRule)) {
                extRule.x14 = true;
            }
            extOnlyRules.push(extRule);
        }
        if (!extOnlyRules.length) {
            continue;
        }
        const { ref } = extCf;
        const existing = cfMap.get(ref);
        if (existing) {
            // reuse existing cf ref
            existing.rules.push(...extOnlyRules);
        } else if (ref !== undefined) {
            // create new cf (the rules of an <x14:conditionalFormatting> without <xm:sqref> (malformed) apply to no cell:
            // they are not read)
            const cf: ConditionalFormattingModel = { ref, rules: extOnlyRules };
            cfMap.set(ref, cf);
            cfs.push(cf);
        }
    }

    return cfs;
};

/**
 * An index of an array of media (`imageId`): a number, or its string (array index semantics: the canonical string of an
 * index, `"1"`, not `"01"`).
 */
function mediumAt(media: readonly MediaModel[], imageId: number | string): MediaModel {
    const index = typeof imageId === "number" ? imageId : Number(imageId);
    const medium = typeof imageId === "number" || String(index) === imageId ? media[index] : undefined;
    if (!medium) {
        throw new Error(`Unknown image ${imageId}`);
    }
    return medium;
}

/** The options of a {@link WorkSheetXform}. */
export interface WorksheetXformOptions {
    /** Parsing fails when the worksheet has more rows than this. Default: no limit. */
    maxRows?: number;
    /** Parsing fails when a row has more cells than this. Default: no limit. */
    maxCols?: number;
    /** The elements of the worksheet not to parse (e.g. `["dataValidations"]`). */
    ignoreNodes?: readonly string[];
    /** The budgets of the workbook being loaded (see WorkSheetXform.createLoadBudget). */
    loadBudget?: WorksheetLoadBudget;
}

/**
 * A worksheet part, `xl/worksheets/sheetN.xml` (ECMA-376 CT_Worksheet): its elements are rendered in the order of the
 * schema (Excel reports a file as corrupt otherwise).
 *
 * Write: {@link prepare} sets up the per-sheet options of the cells (merges, hyperlinks, notes, shared formulae) and the
 * relationships of the worksheet (hyperlinks, notes, images, tables, pivot tables), then {@link render} writes it. Read:
 * the parse gives a ParsedWorksheetModel (the merged cells are checked against the caps of the load budget), which
 * {@link reconcile} resolves with the other parts (notes, hyperlinks, images, tables). The three phases also run row by
 * row (`prepareSteps`, `renderSteps`, `reconcileSteps`), so that the caller may pause between rows (utils/pacer).
 */
class WorkSheetXform extends BaseXform<ParsedWorksheetModel | WorksheetModel> {
    /** The parsed worksheet (the xform renders a WorksheetModel, and parses a ParsedWorksheetModel). */
    declare model: ParsedWorksheetModel | null | undefined;

    /** The namespaces of `<worksheet>`. */
    static readonly WORKSHEET_ATTRIBUTES = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "x14ac",
        "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    } as const;

    /**
     * The budgets shared by all the worksheets of one load.
     *
     * @param maxMergedCells - the cap of merged cells of the workbook
     */
    static createLoadBudget(this: void, maxMergedCells: number = DEFAULT_MAX_MERGED_CELLS): WorksheetLoadBudget {
        return {
            mergedCells: 0,
            maxMergedCells,
            dataValidationCells: DataValidationsXform.MAX_EXPANDED_CELLS,
        };
    }

    /** See {@link WorksheetXformOptions.ignoreNodes}. */
    readonly ignoreNodes: readonly string[];
    /** See {@link WorksheetXformOptions.loadBudget}. */
    readonly loadBudget: WorksheetLoadBudget | undefined;

    declare map: {
        readonly sheetPr: SheetPropertiesXform;
        readonly dimension: DimensionXform;
        readonly sheetViews: ListXform<SheetViewModel>;
        readonly sheetFormatPr: SheetFormatPropertiesXform;
        readonly cols: ListXform<ColModel>;
        readonly sheetData: ListXform<RowModel<ParsedCellModel> | RowModel>;
        readonly autoFilter: AutoFilterXform;
        readonly mergeCells: ListXform<string>;
        readonly rowBreaks: RowBreaksXform;
        readonly hyperlinks: ListXform<SheetHyperlinkModel>;
        readonly pageMargins: PageMarginsXform;
        readonly dataValidations: DataValidationsXform;
        readonly pageSetup: PageSetupXform;
        readonly headerFooter: HeaderFooterXform;
        readonly printOptions: PrintOptionsXform;
        readonly picture: PictureXform;
        readonly drawing: DrawingXform;
        readonly legacyDrawing: LegacyDrawingXform;
        readonly sheetProtection: SheetProtectionXform;
        readonly tableParts: ListXform<SheetPartRefModel>;
        readonly conditionalFormatting: ConditionalFormattingsXform;
        readonly extLst: ExtListXform;
    };

    constructor(options?: WorksheetXformOptions) {
        super();

        const { maxRows, maxCols, ignoreNodes, loadBudget } = options || {};

        this.ignoreNodes = ignoreNodes || [];
        this.loadBudget = loadBudget;

        this.map = {
            sheetPr: new SheetPropertiesXform(),
            dimension: new DimensionXform(),
            sheetViews: new ListXform({
                tag: "sheetViews",
                count: false,
                childXform: new SheetViewXform(),
            }),
            sheetFormatPr: new SheetFormatPropertiesXform(),
            cols: new ListXform({ tag: "cols", count: false, childXform: new ColXform() }),
            sheetData: new ListXform({
                tag: "sheetData",
                count: false,
                empty: true,
                childXform: new RowXform({ maxItems: maxCols }),
                maxItems: maxRows,
            }),
            autoFilter: new AutoFilterXform(),
            mergeCells: new ListXform({ tag: "mergeCells", count: true, childXform: new MergeCellXform() }),
            rowBreaks: new RowBreaksXform(),
            hyperlinks: new ListXform({
                tag: "hyperlinks",
                count: false,
                childXform: new HyperlinkXform(),
            }),
            pageMargins: new PageMarginsXform(),
            dataValidations: new DataValidationsXform({ loadBudget }),
            pageSetup: new PageSetupXform(),
            headerFooter: new HeaderFooterXform(),
            printOptions: new PrintOptionsXform(),
            picture: new PictureXform(),
            drawing: new DrawingXform(),
            legacyDrawing: new LegacyDrawingXform(),
            sheetProtection: new SheetProtectionXform(),
            tableParts: new ListXform({ tag: "tableParts", count: true, childXform: new TablePartXform() }),
            conditionalFormatting: new ConditionalFormattingsXform(),
            extLst: new ExtListXform(),
        };
    }

    /**
     * Prepares the worksheet (mutates the model and `options`), see {@link prepareSteps}.
     *
     * @throws Error for a shared formula clone without its master, or an image of an unknown medium
     */
    override prepare(model: WorksheetModel, options: WorkbookPrepareOptions & Partial<CellPrepareOptions>): void {
        runSteps(this.prepareSteps(model, options));
    }

    // prepare(), render() and reconcile() row by row: the caller may pause between rows (utils/pacer)

    /**
     * {@link prepare}, one row per step: sets the per-sheet options of the cells, prepares the columns, rows and
     * conditional formattings, and sets `mergeCells`, `hyperlinks`, `comments`, `rels`, `drawing`, `background`, `image`
     * on the model (and the dxf ids of the table columns, the `rId` of the tables).
     */
    *prepareSteps(model: WorksheetModel, workbookOptions: WorkbookPrepareOptions & Partial<CellPrepareOptions>): Steps {
        const merges = new Merges();
        const hyperlinks: SheetHyperlinkModel[] = [];
        model.hyperlinks = hyperlinks;
        const comments: CellPrepareOptions["comments"] = [];
        // the per-sheet options of the cells, set on the workbook-level object (mutated)
        const options: WorksheetPrepareOptions = Object.assign(workbookOptions, {
            merges,
            hyperlinks,
            comments,
            // shared formula masters by cell address (no prototype: a sharedFormula never finds an inherited member)
            formulae: emptyRecord<FormulaCellModel>(),
            siFormulae: 0,
        });
        if (options.date1904) {
            // the serial dates of the data validations
            model.date1904 = true;
        }
        this.map.cols.prepare(model.cols, options);
        yield* this.map.sheetData.prepareSteps(model.rows, options);
        this.map.conditionalFormatting.prepare(model.conditionalFormattings, options);

        model.mergeCells = merges.mergeCells;

        // prepare relationships
        const rels: RelationshipModel[] = (model.rels = []);

        const nextRid = (r: readonly RelationshipModel[]): string => `rId${r.length + 1}`;

        for (const hyperlink of hyperlinks) {
            if (HyperlinkXform.isInternalLink(hyperlink.target)) {
                continue;
            }
            const rId = nextRid(rels);
            hyperlink.rId = rId;
            rels.push({
                Id: rId,
                Type: RelType.Hyperlink,
                Target: String(hyperlink.target),
                TargetMode: "External",
            });
        }

        // prepare comment relationships
        if (comments.length > 0) {
            const comment = {
                Id: nextRid(rels),
                Type: RelType.Comments,
                Target: `../comments${model.id}.xml`,
            };
            rels.push(comment);
            const vmlDrawing = {
                Id: nextRid(rels),
                Type: RelType.VmlDrawing,
                Target: `../drawings/vmlDrawing${model.id}.vml`,
            };
            rels.push(vmlDrawing);

            options.commentRefs.push({
                commentName: `comments${model.id}`,
                vmlDrawing: `vmlDrawing${model.id}`,
            });
        }
        // the notes with their decoded reference (the note objects are mutated)
        model.comments = comments.map((item) => Object.assign(item, { refAddress: colCache.decodeAddress(item.ref) }));

        // one image relationship per distinct image of this sheet's drawing
        const imageRIds = new Map<number | string, string>();
        for (const medium of model.media) {
            if (medium.type === "background") {
                const rId = nextRid(rels);
                const bookImage = mediumAt(options.media, medium.imageId);
                rels.push({
                    Id: rId,
                    Type: RelType.Image,
                    Target: `../media/${String(bookImage.name)}.${bookImage.extension}`,
                });
                model.background = {
                    rId,
                };
                model.image = bookImage;
            } else if (medium.type === "image") {
                let { drawing } = model;
                const bookImage = mediumAt(options.media, medium.imageId);
                if (!drawing) {
                    const name = `drawing${++options.drawingsCount}`;
                    const rId = nextRid(rels);
                    drawing = model.drawing = {
                        rId,
                        name,
                        anchors: [],
                        rels: [],
                    };
                    options.drawings.push(drawing);
                    rels.push({
                        Id: rId,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                        Target: `../drawings/${name}.xml`,
                    });
                }
                const drawingRels = (drawing.rels = drawing.rels || []);
                let rIdImage = imageRIds.get(medium.imageId);
                if (!rIdImage) {
                    rIdImage = nextRid(drawingRels);
                    imageRIds.set(medium.imageId, rIdImage);
                    drawingRels.push({
                        Id: rIdImage,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        Target: `../media/${String(bookImage.name)}.${bookImage.extension}`,
                    });
                }

                const anchor: DrawingAnchorModel & { picture: NonNullable<DrawingAnchorModel["picture"]> } = {
                    picture: {
                        rId: rIdImage,
                    },
                    range: medium.range,
                };
                if (medium.hyperlinks && medium.hyperlinks.hyperlink) {
                    const rIdHyperLink = nextRid(drawingRels);
                    anchor.picture.hyperlinks = {
                        tooltip: medium.hyperlinks.tooltip,
                        rId: rIdHyperLink,
                    };
                    drawingRels.push({
                        Id: rIdHyperLink,
                        Type: RelType.Hyperlink,
                        Target: medium.hyperlinks.hyperlink,
                        TargetMode: "External",
                    });
                }
                drawing.anchors.push(anchor);
            }
        }

        // prepare tables
        for (const table of model.tables) {
            // relationships
            const rId = nextRid(rels);
            table.rId = rId;
            rels.push({
                Id: rId,
                Type: RelType.Table,
                Target: `../tables/${String(table.target)}`,
            });

            // dynamic styles
            for (const column of table.columns) {
                const { style } = column;
                if (style) {
                    column.dxfId = options.styles.addDxfStyle(style);
                }
            }
        }

        // prepare pivot tables
        if ((model.pivotTables || []).length) {
            rels.push({
                Id: nextRid(rels),
                Type: RelType.PivotTable,
                Target: "../pivotTables/pivotTable1.xml",
            });
        }

        // prepare ext items
        this.map.extLst.prepare(model);
    }

    /** Writes the worksheet part (with its XML declaration), see {@link renderSteps}. */
    override render(xmlStream: XmlStream, model: WorksheetModel): void {
        runSteps(this.renderSteps(xmlStream, model));
    }

    /** {@link render}, one row per step. The elements follow the CT_Worksheet sequence. */
    *renderSteps(xmlStream: XmlStream, model: WorksheetModel): Steps {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("worksheet", WorkSheetXform.WORKSHEET_ATTRIBUTES);

        const { properties, pageSetup } = model;
        const sheetFormatPropertiesModel: SheetFormatPropertiesModel | undefined = properties
            ? {
                  defaultRowHeight: properties.defaultRowHeight,
                  dyDescent: properties.dyDescent,
                  outlineLevelCol: properties.outlineLevelCol,
                  outlineLevelRow: properties.outlineLevelRow,
              }
            : undefined;
        if (properties && properties.defaultColWidth && sheetFormatPropertiesModel) {
            sheetFormatPropertiesModel.defaultColWidth = properties.defaultColWidth;
        }
        const sheetPropertiesModel: SheetPropertiesModel = {
            outlineProperties: properties && properties.outlineProperties,
            tabColor: properties && properties.tabColor,
            pageSetup:
                pageSetup && pageSetup.fitToPage
                    ? {
                          fitToPage: pageSetup.fitToPage,
                      }
                    : undefined,
        };
        const pageMarginsModel = pageSetup && pageSetup.margins;
        const printOptionsModel: Partial<PrintOptionsModel> = {
            showRowColHeaders: pageSetup && pageSetup.showRowColHeaders,
            showGridLines: pageSetup && pageSetup.showGridLines,
            horizontalCentered: pageSetup && pageSetup.horizontalCentered,
            verticalCentered: pageSetup && pageSetup.verticalCentered,
        };
        const sheetProtectionModel = model.sheetProtection;
        // the document gives the used range as a Range (rendered as its range string)
        const { dimensions } = model;
        const dimensionsRef = dimensions ? (typeof dimensions === "string" ? dimensions : dimensions.range) : undefined;

        this.map.sheetPr.render(xmlStream, sheetPropertiesModel);
        this.map.dimension.render(xmlStream, dimensionsRef);
        this.map.sheetViews.render(xmlStream, model.views);
        this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel);
        this.map.cols.render(xmlStream, model.cols);
        yield* this.map.sheetData.renderSteps(xmlStream, model.rows);
        this.map.sheetProtection.render(xmlStream, sheetProtectionModel); // Note: must be after sheetData and before autoFilter
        this.map.autoFilter.render(xmlStream, model.autoFilter);
        this.map.mergeCells.render(xmlStream, model.mergeCells);
        this.map.conditionalFormatting.render(xmlStream, model.conditionalFormattings); // Note: must be before dataValidations
        this.map.dataValidations.render(xmlStream, model.dataValidations, model.date1904);

        // For some reason hyperlinks have to be after the data validations
        this.map.hyperlinks.render(xmlStream, model.hyperlinks);

        this.map.printOptions.render(xmlStream, printOptionsModel); // Note: must be before pageMargins
        this.map.pageMargins.render(xmlStream, pageMarginsModel);
        this.map.pageSetup.render(xmlStream, pageSetup);
        this.map.headerFooter.render(xmlStream, model.headerFooter);
        this.map.rowBreaks.render(xmlStream, model.rowBreaks);
        // CT_Worksheet order: drawing, legacyDrawing, picture, tableParts, extLst
        this.map.drawing.render(xmlStream, model.drawing); // Note: must be after rowBreaks
        if (model.rels) {
            // add a <legacyDrawing /> node for the comments (vml drawing)
            for (const rel of model.rels) {
                if (rel.Type === RelType.VmlDrawing) {
                    xmlStream.leafNode("legacyDrawing", { "r:id": rel.Id });
                }
            }
        }
        this.map.picture.render(xmlStream, model.background);
        this.map.tableParts.render(xmlStream, model.tables);

        this.map.extLst.render(xmlStream, model);

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the part is accepted (the unknown and ignored ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        if (node.name === "worksheet") {
            for (const xform of Object.values(this.map)) {
                xform.reset();
            }
            return true;
        }

        // element names come from the file: own entries of the map only
        const parser = ownEntry(this.map, node.name);
        if (parser && !this.ignoreNodes.includes(node.name)) {
            this.parser = parser;
            parser.parseOpen(node);
        }
        return true;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /**
     * @returns `false` at the end of `<worksheet>` (the model is then built), else `true`
     * @throws Error when the merged cells exceed the caps of the load budget
     */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "worksheet": {
                const { map } = this;
                const sheetPr = map.sheetPr.model;
                const properties: WorksheetPropertiesModel = map.sheetFormatPr.model || {};
                if (sheetPr && sheetPr.tabColor) {
                    properties.tabColor = sheetPr.tabColor;
                }
                if (sheetPr && sheetPr.outlineProperties) {
                    properties.outlineProperties = sheetPr.outlineProperties;
                }
                const sheetProperties: PageSetupModel = {
                    fitToPage: (sheetPr && sheetPr.pageSetup && sheetPr.pageSetup.fitToPage) || false,
                    margins: map.pageMargins.model,
                };
                const pageSetup: PageSetupModel = Object.assign(sheetProperties, map.pageSetup.model, map.printOptions.model);
                const extLst = map.extLst.model;
                const conditionalFormattings = mergeConditionalFormattings(
                    // reset() set it to []
                    map.conditionalFormatting.model ?? [],
                    extLst && extLst["x14:conditionalFormattings"]
                );
                if (this.loadBudget) {
                    spendMergedCells(this.loadBudget, map.mergeCells.model);
                }
                const model: ParsedWorksheetModel = (this.model = {
                    dimensions: map.dimension.model,
                    cols: map.cols.model,
                    // parsed items are RowXform parse models (the list type also covers the rendered document rows)
                    // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- the items of <sheetData> are parsed by RowXform, whose parse model is RowModel<ParsedCellModel> (ListXform has one item type for render and parse)
                    rows: map.sheetData.model as RowModel<ParsedCellModel>[] | null | undefined,
                    mergeCells: map.mergeCells.model,
                    rowBreaks: map.rowBreaks.model ?? [],
                    hyperlinks: map.hyperlinks.model,
                    dataValidations: map.dataValidations.model,
                    properties,
                    // without <sheetViews> / <headerFooter>: as a new worksheet has them (rendered as nothing)
                    views: map.sheetViews.model ?? [],
                    pageSetup,
                    headerFooter: map.headerFooter.model ?? defaultHeaderFooter(),
                    background: map.picture.model,
                    drawing: map.drawing.model,
                    tables: map.tableParts.model,
                    conditionalFormattings,
                });

                if (map.autoFilter.model) {
                    model.autoFilter = map.autoFilter.model;
                }
                if (map.sheetProtection.model) {
                    model.sheetProtection = map.sheetProtection.model;
                }
                if (map.legacyDrawing.model) {
                    model.legacyDrawing = map.legacyDrawing.model;
                }

                return false;
            }

            default:
                // not quite sure how we get here!
                return true;
        }
    }

    /** Resolves the parsed worksheet with the other parts, see {@link reconcileSteps}. */
    override reconcile(model: ParsedWorksheetModel, options: WorkbookReconcileOptions & Partial<CellReconcileOptions>): void {
        runSteps(this.reconcileSteps(model, options));
    }

    /**
     * {@link reconcile}, one row per step: attaches the notes (comments and VML parts), the hyperlinks, the images
     * (drawing, background) and the tables of the worksheet, compacts the rows and cells (inferring missing references,
     * creating the cells of the notes), and reconciles the columns, rows and conditional formattings. A part claimed by
     * another worksheet is ignored. Mutates the model (deletes `relationships`, `legacyDrawing`, `hyperlinks`, `comments`)
     * and `options` (the per-sheet maps).
     */
    *reconcileSteps(model: ParsedWorksheetModel, workbookOptions: WorkbookReconcileOptions & Partial<CellReconcileOptions>): Steps {
        // options.merges = new Merges();
        // options.merges.reconcile(model.mergeCells, model.rows);
        const relationships = model.relationships || [];
        // keyed by the relationship ids of the file
        const rels = emptyRecord<RelationshipModel>();
        for (const rel of relationships) {
            rels[rel.Id] = rel;
        }
        // an object key: a missing id is "undefined"
        const relOf = (rId: string | undefined): RelationshipModel | undefined => rels[String(rId)];
        // A comments, VML, drawing or table part belongs to one worksheet. Relationships of other worksheets to the same
        // part are ignored: they would materialise it once per worksheet (a small file could then hold millions of notes).
        const claimedParts = workbookOptions.claimedParts || new Set();
        const claim = <T extends object>(part: T | null | undefined): T | undefined => {
            if (!part || claimedParts.has(part)) {
                return undefined;
            }
            claimedParts.add(part);
            return part;
        };
        // a relationship may point at a part that is missing or was not recognised: ignore it
        const findPart = <T extends object>(type: string, parts: Readonly<Record<string, T>> | undefined): T | undefined =>
            relationships
                .filter((rel) => rel.Type === type)
                .map((rel) => getPart(parts, rel))
                .find((part) => part && !claimedParts.has(part));
        const commentsPart = claim(findPart(RelType.Comments, workbookOptions.comments));
        if (commentsPart) {
            const { comments } = commentsPart;
            model.comments = comments;
            // the notes vml is the one named by <legacyDrawing> (<legacyDrawingHF> has header/footer
            // images); without that element, fall back to the first vml drawing relationship
            const vmlRel = model.legacyDrawing
                ? relOf(model.legacyDrawing.rId)
                : relationships.find((rel) => rel.Type === RelType.VmlDrawing);
            const vmlPart = claim(getPart(workbookOptions.vmlDrawings, vmlRel));
            if (vmlPart) {
                comments.forEach((comment, index) => {
                    comment.note = Object.assign({}, comment.note, vmlPart.comments[index]);
                });
            }
        }
        // keyed by the cell references of the file: no prototype, so that a lookup never finds an
        // inherited member and a "__proto__" reference is an ordinary key
        const commentsMap = emptyRecord<NoteModel>();
        for (const comment of model.comments || []) {
            if (comment.ref) {
                commentsMap[comment.ref] = comment;
            }
        }
        const hyperlinkMap: WorksheetReconcileOptions["hyperlinkMap"] = emptyRecord();
        for (const hyperlink of model.hyperlinks || []) {
            const rel = hyperlink.rId && relOf(hyperlink.rId);
            const target = getHyperlinkTarget(rel ? rel.Target : undefined, hyperlink.target);
            if (target) {
                hyperlinkMap[String(hyperlink.address)] = { target, tooltip: hyperlink.tooltip };
            }
        }
        // the per-sheet options of the cells, set on the workbook-level object (mutated)
        const options: WorksheetReconcileOptions = Object.assign(workbookOptions, {
            commentsMap,
            hyperlinkMap,
            // shared formula masters by si (from the file)
            formulae: emptyRecord<string | undefined>(),
        });

        // compact the rows and cells
        const rows = (model.rows && model.rows.filter(Boolean)) || [];
        model.rows = rows;
        let previousRowNumber = 0;
        for (const row of rows) {
            row.cells = (row.cells && row.cells.filter(Boolean)) || [];
            inferMissingReferences(row, previousRowNumber);
            previousRowNumber = row.number;
        }
        addNoteCells(rows, model.comments);

        this.map.cols.reconcile(model.cols || undefined, options);
        yield* this.map.sheetData.reconcileSteps(rows, options);
        this.map.conditionalFormatting.reconcile(model.conditionalFormattings, options);
        this.map.dataValidations.reconcile(model.dataValidations, options);

        const media: NonNullable<ParsedWorksheetModel["media"]> = (model.media = []);
        if (model.drawing) {
            const drawingRel = relOf(model.drawing.rId);
            const match =
                drawingRel &&
                typeof drawingRel.Target === "string" &&
                drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
            const drawing: DrawingModel | undefined = match ? claim(getOwn(options.drawings, match[1])) : undefined;
            if (drawing) {
                for (const anchor of drawing.anchors) {
                    // the index is set by WorkbookXform.reconcile for every medium of the workbook
                    const imageId = anchor.medium && anchor.medium.index;
                    if (imageId !== undefined) {
                        media.push({
                            type: "image",
                            imageId,
                            range: anchor.range,
                            hyperlinks: anchor.picture && anchor.picture.hyperlinks,
                        });
                    }
                }
            }
        }

        const backgroundRel = model.background && relOf(model.background.rId);
        if (backgroundRel && typeof backgroundRel.Target === "string") {
            const target = backgroundRel.Target.split("/media/")[1];
            const imageId = getOwn(options.mediaIndex, target);
            if (imageId !== undefined) {
                media.push({
                    type: "background",
                    imageId,
                });
            }
        }

        const tables: TableModel[] = [];
        for (const tablePart of model.tables || []) {
            const table = claim(getPart(options.tables, relOf(tablePart.rId)));
            if (table) {
                tables.push(table);
            }
        }
        model.tables = tables;

        delete model.relationships;
        delete model.legacyDrawing;
        delete model.hyperlinks;
        delete model.comments;
    }
}

export default WorkSheetXform;
