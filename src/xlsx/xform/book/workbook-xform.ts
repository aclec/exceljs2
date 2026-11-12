import type { CalculationProperties, WorkbookProperties } from "../../../../index";
import type { SaxInput, SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type {
    DefinedNameModel,
    MediaModel,
    ParsedWorksheetModel,
    PageSetupModel,
    RelationshipModel,
    SheetDefModel,
    WorkbookPartModel,
    WorkbookPivotCacheModel,
    WorkbookReadModel,
    WorkbookViewModel,
} from "../../../types";

import BaseXform, { ownEntry } from "../base-xform";
import StaticXform from "../static-xform";
import ListXform from "../list-xform";
import DefinedNameXform from "./defined-name-xform";
import SheetXform from "./sheet-xform";
import WorkbookViewXform from "./workbook-view-xform";
import WorkbookPropertiesXform from "./workbook-properties-xform";
import WorkbookCalcPropertiesXform from "./workbook-calc-properties-xform";
import WorkbookPivotCacheXform from "./workbook-pivot-cache-xform";

// Worksheets are stored in an array indexed by sheetId. The file format allows any unsigned int, but Excel
// numbers sheets from 1 as they are added, so a huge id is hostile: sheetId="4294967294" made an array whose
// every traversal (workbook.worksheets, getWorksheet(name), addWorksheet, write) took minutes.
const MAX_SHEET_ID = 65535;

// The absolute reference ($A$1, $A or $1) to a cell, a column or a row given with or without $ (A1, $A1, A$1...): a print
// area or print title is a defined name, in which a relative reference would depend on the active cell. Anything else is
// kept as given.
const CELL_REF = /^\$?([A-Za-z]{1,3})?\$?(\d+)?$/;
function absoluteRef(ref: string): string {
    const match = CELL_REF.exec(ref);
    if (!match || (!match[1] && !match[2])) {
        return ref;
    }
    return (match[1] ? `$${match[1]}` : "") + (match[2] ? `$${match[2]}` : "");
}

// the absolute form of a range (A1:B2, A:C, 1:3) or of a cell
const absoluteRange = (range: string): string => range.split(":").map(absoluteRef).join(":");

// Keeps valid ids; gives the lowest free ids to sheets whose id is not an integer in 1..MAX_SHEET_ID or is
// already used by an earlier sheet (which used to replace that sheet).
function renumberSheets(sheets: readonly SheetDefModel[]): void {
    const used = new Set<number>();
    const invalid = sheets.filter((sheet) => {
        const valid = Number.isInteger(sheet.id) && sheet.id >= 1 && sheet.id <= MAX_SHEET_ID && !used.has(sheet.id);
        if (valid) {
            used.add(sheet.id);
        }
        return !valid;
    });
    let id = 1;
    for (const sheet of invalid) {
        while (used.has(id)) {
            id++;
        }
        sheet.id = id;
        used.add(id);
    }
}

/** A worksheet as {@link WorkbookXform} writes it: its `<sheet>` and its print areas and titles. */
export interface WorkbookSheetWriteModel extends SheetDefModel {
    readonly pageSetup?: Pick<PageSetupModel, "printArea" | "printTitlesRow" | "printTitlesColumn"> | null;
}

/**
 * What {@link WorkbookXform} prepares and renders: the workbook model of xlsx.ts or of the stream WorkbookWriter.
 * prepare sets `sheets` (the worksheets), adds the print areas and titles to `definedNames` and names the media.
 */
export interface WorkbookXformWriteModel {
    worksheets: WorkbookSheetWriteModel[];
    sheets?: readonly WorkbookSheetWriteModel[];
    definedNames: DefinedNameModel[];
    views: readonly WorkbookViewModel[];
    properties: Partial<WorkbookProperties>;
    calcProperties: Partial<CalculationProperties>;
    media?: readonly MediaModel[];
    /** a `<pivotCache>` per pivot table */
    pivotTables?: readonly WorkbookPivotCacheModel[];
}

/**
 * The workbook part, `xl/workbook.xml`: the worksheets (`<sheets>`), the defined names (with the print areas and titles
 * of the worksheets), the views and the properties.
 *
 * - Write: {@link prepare} then {@link render} a {@link WorkbookXformWriteModel}.
 * - Read: parsing gives a {@link WorkbookPartModel} (the model of this xform); {@link reconcile} then completes the read
 *   model of xlsx.ts (worksheet names / ids / states, print areas and titles, media indexes).
 */
class WorkbookXform extends BaseXform<WorkbookPartModel | WorkbookXformWriteModel | WorkbookReadModel> {
    /** The namespaces of `<workbook>`. */
    static readonly WORKBOOK_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "x15",
        "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    };

    /** The fixed parts of the workbook part. */
    static readonly STATIC_XFORMS = {
        fileVersion: new StaticXform({
            tag: "fileVersion",
            $: { appName: "xl", lastEdited: 5, lowestEdited: 5, rupBuild: 9303 },
        }),
    };

    /** The parsed part (prepare / render / reconcile take other models: see the type parameter). */
    declare model: WorkbookPartModel | null | undefined;

    declare map: {
        readonly fileVersion: StaticXform;
        readonly workbookPr: WorkbookPropertiesXform;
        readonly bookViews: ListXform<WorkbookViewModel>;
        readonly sheets: ListXform<SheetDefModel>;
        readonly definedNames: ListXform<DefinedNameModel>;
        readonly calcPr: WorkbookCalcPropertiesXform;
        readonly pivotCaches: ListXform<WorkbookPivotCacheModel>;
    };

    constructor() {
        super();

        this.map = {
            fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
            workbookPr: new WorkbookPropertiesXform(),
            bookViews: new ListXform({
                tag: "bookViews",
                count: false,
                childXform: new WorkbookViewXform(),
            }),
            sheets: new ListXform({ tag: "sheets", count: false, childXform: new SheetXform() }),
            definedNames: new ListXform({
                tag: "definedNames",
                count: false,
                childXform: new DefinedNameXform(),
            }),
            calcPr: new WorkbookCalcPropertiesXform(),
            pivotCaches: new ListXform({
                tag: "pivotCaches",
                count: false,
                childXform: new WorkbookPivotCacheXform(),
            }),
        };
    }

    /** Sets `sheets`, adds the print areas and titles of the worksheets to the defined names, names the media. */
    override prepare(model: WorkbookXformWriteModel): void {
        const sheets = model.worksheets;
        model.sheets = sheets;

        // collate all the print areas from all of the sheets and add them to the defined names
        const printAreas: DefinedNameModel[] = [];
        let index = 0; // sheets is sparse array - calc index manually
        sheets.forEach((sheet) => {
            // quoted sheet name for a formula reference: embedded quotes are doubled
            const sheetRef = `'${sheet.name.replace(/'/g, "''")}'`;
            if (sheet.pageSetup && sheet.pageSetup.printArea) {
                // a sheet has at most one _xlnm.Print_Area: several areas are comma-separated ranges
                const ranges = sheet.pageSetup.printArea.split("&&").map((printArea) => `${sheetRef}!${absoluteRange(printArea)}`);
                printAreas.push({
                    name: "_xlnm.Print_Area",
                    ranges,
                    localSheetId: index,
                });
            }

            if (sheet.pageSetup && (sheet.pageSetup.printTitlesRow || sheet.pageSetup.printTitlesColumn)) {
                const ranges: string[] = [];

                if (sheet.pageSetup.printTitlesColumn) {
                    ranges.push(`${sheetRef}!${absoluteRange(sheet.pageSetup.printTitlesColumn)}`);
                }

                if (sheet.pageSetup.printTitlesRow) {
                    ranges.push(`${sheetRef}!${absoluteRange(sheet.pageSetup.printTitlesRow)}`);
                }

                const definedName: DefinedNameModel = {
                    name: "_xlnm.Print_Titles",
                    ranges,
                    localSheetId: index,
                };

                printAreas.push(definedName);
            }
            index++;
        });
        if (printAreas.length) {
            model.definedNames = model.definedNames.concat(printAreas);
        }

        (model.media || []).forEach((medium, i) => {
            // assign name
            medium.name = medium.type + (i + 1);
        });
    }

    override render(xmlStream: XmlStream, model: WorkbookXformWriteModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("workbook", WorkbookXform.WORKBOOK_ATTRIBUTES);

        this.map.fileVersion.render(xmlStream);
        this.map.workbookPr.render(xmlStream, model.properties);
        this.map.bookViews.render(xmlStream, model.views);
        this.map.sheets.render(xmlStream, model.sheets);
        this.map.definedNames.render(xmlStream, model.definedNames);
        this.map.calcPr.render(xmlStream, model.calcProperties);
        this.map.pivotCaches.render(xmlStream, model.pivotTables);

        xmlStream.closeNode();
    }

    /** @returns the parsed part */
    override async parseStream(stream: SaxInput): Promise<WorkbookPartModel | null | undefined> {
        await super.parseStream(stream);
        return this.model;
    }

    /** @returns `true`: every element of the part is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "workbook":
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
            case "workbook": {
                // (`?? undefined`: the lists are null only after a reset, which this xform never does)
                const model: WorkbookPartModel = {
                    sheets: this.map.sheets.model ?? undefined,
                    properties: this.map.workbookPr.model || {},
                    views: this.map.bookViews.model ?? [],
                    calcProperties: {},
                };
                if (this.map.definedNames.model) {
                    model.definedNames = this.map.definedNames.model;
                }
                this.model = model;

                return false;
            }
            default:
                // not quite sure how we get here!
                return true;
        }
    }

    /**
     * Completes the read model of xlsx.ts: the name, id and state of each worksheet (renumbering invalid or duplicate
     * ids), its print areas and titles (the other defined names stay in `definedNames`), and the index of each medium.
     */
    override reconcile(model: WorkbookReadModel): void {
        // keyed by ids from the file: a Map, so that an id such as "constructor" is unknown
        const rels = new Map<string, RelationshipModel>();
        for (const rel of model.workbookRels || []) {
            rels.set(rel.Id, rel);
        }

        renumberSheets(model.sheets || []);

        // reconcile sheet ids, rIds and names
        const worksheets: ParsedWorksheetModel[] = [];
        let worksheet: ParsedWorksheetModel | undefined;
        let index = 0;

        (model.sheets || []).forEach((sheet) => {
            const rel = sheet.rId === undefined ? undefined : rels.get(sheet.rId);
            if (!rel || typeof rel.Target !== "string") {
                return;
            }
            // if rel.Target start with `[space]/xl/` or `/xl/` , then it will be replaced with `''` and spliced behind `xl/`,
            // otherwise it will be spliced directly behind `xl/`. i.g.
            worksheet = ownEntry(model.worksheetHash, `xl/${rel.Target.replace(/^(\s|\/xl\/)+/, "")}`);
            // If there are "chartsheets" in the file, rel.Target will
            // come out as chartsheets/sheet1.xml or similar here, and
            // that won't be in model.worksheetHash.
            // As we don't have the infrastructure to support chartsheets,
            // we will ignore them for now:
            if (worksheet) {
                worksheet.name = sheet.name;
                worksheet.id = sheet.id;
                worksheet.state = sheet.state;
                worksheets[index++] = worksheet;
            }
        });

        // reconcile print areas
        const definedNames: DefinedNameModel[] = [];
        // the areas of each worksheet, joined once at the end: legacy files may hold several Print_Area names
        // for the same sheet, and re-joining the accumulated area for each name was quadratic
        const printAreas = new Map<ParsedWorksheetModel, string[]>();
        for (const definedName of model.definedNames || []) {
            if (definedName.name === "_xlnm.Print_Area") {
                worksheet = sheetAt(worksheets, definedName.localSheetId);
                if (worksheet) {
                    if (!worksheet.pageSetup) {
                        worksheet.pageSetup = {};
                    }
                    // keep the reference as written (minus sheet and $): an area may be a cell, a range,
                    // or whole columns/rows (A:C, 1:3); #REF! areas are dropped
                    let areas = printAreas.get(worksheet);
                    if (!areas) {
                        areas = worksheet.pageSetup.printArea ? [worksheet.pageSetup.printArea] : [];
                        printAreas.set(worksheet, areas);
                    }
                    definedName.ranges.forEach((range) => {
                        const area = range.slice(range.lastIndexOf("!") + 1).replace(/\$/g, "");
                        if (area && !area.startsWith("#")) {
                            areas.push(area);
                        }
                    });
                }
            } else if (definedName.name === "_xlnm.Print_Titles") {
                worksheet = sheetAt(worksheets, definedName.localSheetId);
                if (worksheet) {
                    if (!worksheet.pageSetup) {
                        worksheet.pageSetup = {};
                    }

                    const rangeString = definedName.ranges.join(",");

                    const dollarRegex = /\$/g;

                    const rowRangeRegex = /\$\d+:\$\d+/;
                    const rowRangeMatches = rangeString.match(rowRangeRegex);

                    if (rowRangeMatches && rowRangeMatches.length) {
                        const range = rowRangeMatches[0];
                        worksheet.pageSetup.printTitlesRow = range.replace(dollarRegex, "");
                    }

                    const columnRangeRegex = /\$[A-Z]+:\$[A-Z]+/;
                    const columnRangeMatches = rangeString.match(columnRangeRegex);

                    if (columnRangeMatches && columnRangeMatches.length) {
                        const range = columnRangeMatches[0];
                        worksheet.pageSetup.printTitlesColumn = range.replace(dollarRegex, "");
                    }
                }
            } else {
                definedNames.push(definedName);
            }
        }
        printAreas.forEach((areas, sheet) => {
            sheet.pageSetup.printArea = areas.join("&&");
        });
        model.definedNames = definedNames;

        // used by sheets to build their image models
        model.media.forEach((media, i) => {
            media.index = i;
        });
    }
}

// the worksheet at a localSheetId of a defined name (absent or not an index: none)
function sheetAt(worksheets: readonly ParsedWorksheetModel[], localSheetId: number | undefined): ParsedWorksheetModel | undefined {
    return localSheetId === undefined ? undefined : worksheets[localSheetId];
}

export default WorkbookXform;
