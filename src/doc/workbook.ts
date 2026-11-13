// A workbook document: worksheets, defined names, images, properties; read and written by `xlsx` and `csv`.
//
// Workbook requirements
//  Load and Save from file and stream
//  Access/Add/Delete individual worksheets
//  Manage String table, Hyperlink table, etc.
//  Manage scaffolding for contained objects to write to/read from

import type { AddChartsheetOptions, CalculationProperties, Chartsheet, ChartsheetChartOptions, WorkbookProperties } from "../../index";
import Worksheet, { type DocumentWorksheetModel, type WorksheetOptions } from "./worksheet";
import DefinedNames from "./defined-names";
import XLSX from "../xlsx/xlsx";
import CSV from "../csv/csv";
import { DEFAULT_MAX_MERGED_CELLS } from "../utils/merge-limits";
import type { PivotTable } from "./pivot-table";
import { createdChartsheetPart, mergeSheets, withTabSelected } from "../xlsx/chartsheet-parts";
import { ChartsheetChart } from "./chart";
import { checkSheetName, chartsheetState } from "./sheet-name";
import type { ChartsheetModel, ColorModel, MediaModel, WorkbookModel, WorkbookViewModel } from "../types";

/** The options of `addWorksheet` (index.d.ts `AddWorksheetOptions`). */
export type AddWorksheetOptions = Partial<Omit<WorksheetOptions, "workbook" | "id" | "name" | "orderNo">>;

/** Whether the options of `addWorksheet` are a color (deprecated: the tab color), told by its `argb`, `theme` or `indexed`. */
function isTabColor(options: AddWorksheetOptions | ColorModel): options is ColorModel {
    return ("argb" in options && !!options.argb) || ("theme" in options && !!options.theme) || ("indexed" in options && !!options.indexed);
}

/** The model of a Workbook (its `model` getter): a {@link WorkbookModel} whose worksheet models come from Worksheet. */
export type DocumentWorkbookModel = Omit<WorkbookModel, "worksheets"> & { worksheets: DocumentWorksheetModel[] };

/** An image added to the workbook (`workbook.addImage`): its extension and its data (a file name, a buffer or base64). */
export type ImageInput = Omit<MediaModel, "type" | "name" | "index">;

/**
 * A workbook: its worksheets (by id, in their order), defined names, images and document properties. `xlsx` and `csv`
 * read and write it.
 */
class Workbook {
    /** Document property: the category ("" by default). */
    category: string | undefined;
    /** Document property: the company ("" by default). */
    company: string | undefined;
    /** Document property: the creation date (now by default). */
    created: Date | undefined;
    /** Document property: the description ("" by default). */
    description: string | undefined;
    /** Document property: the keywords ("" by default). */
    keywords: string | undefined;
    /** Document property: the manager ("" by default). */
    manager: string | undefined;
    /** Document property: the date of the last change (now by default). */
    modified: Date | undefined;
    /** Workbook properties (e.g. `date1904`). */
    properties: Partial<WorkbookProperties>;
    /** Calculation properties (e.g. `fullCalcOnLoad`). */
    calcProperties: Partial<CalculationProperties>;
    /** The worksheets, indexed by id (sparse: index 0 and removed ids are empty). */
    _worksheets: (Worksheet | undefined)[];
    /**
     * The chartsheets read from a file, carried to be written back (src/xlsx/chartsheet-parts.ts), and those created by
     * {@link addChartsheet} (with their `chart`, generated when the workbook is written), each with its place among the
     * sheets: `orderNo` is compared with the `orderNo` of the worksheets.
     */
    _chartsheets: { readonly model: ChartsheetModel; orderNo: number; readonly chart?: ChartsheetChart }[];
    /** Document property: the subject ("" by default). */
    subject: string | undefined;
    /** Document property: the title ("" by default). */
    title: string | undefined;
    /** The workbook views (windows): `[]` by default, also once a file without `<bookViews>` is read. */
    views: WorkbookViewModel[];
    /** The images (`addImage`), indexed by image id. */
    media: MediaModel[];
    /** The pivot tables of the worksheets (`worksheet.addPivotTable`, one at most). */
    pivotTables: PivotTable[];
    _definedNames: DefinedNames;
    /**
     * Cap on the cells covered by merges in all the worksheets (utils/merge-limits); also the default `maxMergedCells`
     * of the xlsx reader.
     */
    maxMergedCells: number;
    /** The number of merged cells of all the worksheets, kept up to date by the worksheets. */
    _mergedCellCount: number;
    // set by the application or the model setter
    /** Document property: the author. */
    declare creator?: string | undefined;
    /** Document property: the author of the last change. */
    declare lastModifiedBy?: string | undefined;
    /** Document property: the date of the last print. */
    declare lastPrinted?: Date | undefined;
    /** Document property: the language. */
    declare language?: string | undefined;
    /** Document property: the revision number. */
    declare revision?: number | undefined;
    /** Document property: the content status (`"Draft"`, `"Final"`…). */
    declare contentStatus?: string | undefined;
    /** The themes read from a file (theme part name → XML), written back as they are. */
    declare _themes?: Record<string, string> | undefined;
    declare _xlsx?: XLSX<Workbook>;
    declare _csv?: CSV;

    constructor() {
        this.category = "";
        this.company = "";
        this.created = new Date();
        this.description = "";
        this.keywords = "";
        this.manager = "";
        this.modified = this.created;
        this.properties = {};
        this.calcProperties = {};
        this._worksheets = [];
        this._chartsheets = [];
        this.subject = "";
        this.title = "";
        this.views = [];
        this.media = [];
        this.pivotTables = [];
        this._definedNames = new DefinedNames();
        // cap on the cells covered by merges in all the worksheets (utils/merge-limits), and their number, kept up to date by
        // the worksheets
        this.maxMergedCells = DEFAULT_MAX_MERGED_CELLS;
        this._mergedCellCount = 0;
    }

    /** The xlsx reader / writer (created on first use). */
    get xlsx(): XLSX<Workbook> {
        if (!this._xlsx) this._xlsx = new XLSX(this);
        return this._xlsx;
    }

    /** The csv reader / writer (created on first use). */
    get csv(): CSV {
        if (!this._csv) this._csv = new CSV(this);
        return this._csv;
    }

    /** The id of the next worksheet: the first id (from 1) of no worksheet and no chartsheet. */
    get nextId(): number {
        const chartsheetIds = new Set(this._chartsheets.map(({ model }) => model.id));
        // find the next unique spot to add worksheet
        let i = 1;
        while (this._worksheets[i] || chartsheetIds.has(i)) {
            i++;
        }
        return i;
    }

    /** The highest `orderNo` of the sheets (0 without any). */
    private lastOrderNo(): number {
        return this._chartsheets.reduce(
            (acc, chartsheet) => (chartsheet.orderNo > acc ? chartsheet.orderNo : acc),
            this._worksheets.reduce((acc, ws) => (ws && ws.orderNo !== undefined && ws.orderNo > acc ? ws.orderNo : acc), 0)
        );
    }

    /**
     * Adds a worksheet after the others.
     * @param options - the options; a color (or an ARGB string) is the tab color (deprecated)
     * @throws Error for an invalid name (see the Worksheet name setter)
     */
    addWorksheet(name?: string, options?: AddWorksheetOptions | ColorModel | string): Worksheet {
        const id = this.nextId;

        // if options is a color, call it tabColor (and signal deprecated message)
        let worksheetOptions: AddWorksheetOptions | undefined;
        if (options) {
            if (typeof options === "string") {
                // eslint-disable-next-line no-console -- a deprecation notice with its stack trace, as it always was
                console.trace(
                    'tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { argb: "rbg value" } }'
                );
                worksheetOptions = {
                    properties: {
                        tabColor: { argb: options },
                    },
                };
            } else if (isTabColor(options)) {
                // eslint-disable-next-line no-console -- a deprecation notice with its stack trace, as it always was
                console.trace(
                    "tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }"
                );
                worksheetOptions = {
                    properties: {
                        tabColor: options,
                    },
                };
            } else {
                worksheetOptions = options;
            }
        }

        const fullOptions: WorksheetOptions = Object.assign({}, worksheetOptions, {
            id,
            name,
            orderNo: this.lastOrderNo() + 1,
            workbook: this,
        });

        const worksheet = new Worksheet(fullOptions);

        this._worksheets[id] = worksheet;
        return worksheet;
    }

    /**
     * Removes a worksheet (a worksheet removed twice does not remove the worksheet that took its id). The `activeTab` /
     * `firstSheet` of the workbook views keep their sheets (the next one for the sheet removed, the last one at the end).
     */
    removeWorksheetEx(worksheet: Worksheet): void {
        // a worksheet removed twice must not remove the worksheet that took its id
        if (this._worksheets[worksheet.id] === worksheet) {
            const position = this._sheetOrder().indexOf(worksheet);
            this._mergedCellCount -= worksheet._mergedCellCount;
            // eslint-disable-next-line @typescript-eslint/no-array-delete -- deliberate: no splice, the indexes are the worksheet ids
            delete this._worksheets[worksheet.id];
            this._sheetRemoved(position);
        }
    }

    /** The worksheets and the chartsheets (their entries), in their order among the sheets (a worksheet of the same orderNo first). */
    _sheetOrder(): (Worksheet | Workbook["_chartsheets"][number])[] {
        const { worksheets } = this;
        const sheets: (Worksheet | Workbook["_chartsheets"][number])[] = [];
        let w = 0;
        for (const entry of this._chartsheets) {
            for (let worksheet = worksheets[w]; worksheet && !(entry.orderNo < Number(worksheet.orderNo)); worksheet = worksheets[++w]) {
                sheets.push(worksheet);
            }
            sheets.push(entry);
        }
        return sheets.concat(worksheets.slice(w));
    }

    /**
     * Moves the `activeTab` / `firstSheet` of the workbook views once the sheet at a position (among all the sheets) is
     * removed: those after it move back by one, those on it stay (the next sheet), all within the sheets left.
     */
    _sheetRemoved(position: number): void {
        if (position < 0) {
            return;
        }
        const last = Math.max(this._worksheets.filter(Boolean).length + this._chartsheets.length - 1, 0);
        const move = (index: number | undefined): number | undefined =>
            index === undefined || !Number.isInteger(index) ? index : Math.min(index > position ? index - 1 : index, last);
        for (const view of this.views) {
            if (view.activeTab !== undefined) {
                view.activeTab = move(view.activeTab);
            }
            if (view.firstSheet !== undefined) {
                view.firstSheet = move(view.firstSheet);
            }
        }
    }

    /** Removes a worksheet by id or name. */
    removeWorksheet(id: number | string): void {
        const worksheet = this.getWorksheet(id);
        if (worksheet) {
            worksheet.destroy();
        }
    }

    /** A worksheet by id or name; without argument, the first worksheet (by id). */
    getWorksheet(id?: number | string): Worksheet | undefined {
        if (id === undefined) {
            return this._worksheets.find(Boolean);
        }
        if (typeof id === "number") {
            return this._worksheets[id];
        }
        if (typeof id === "string") {
            return this._worksheets.find((worksheet) => worksheet && worksheet.name === id);
        }
        return undefined;
    }

    /** The worksheets, in their order (a new array). */
    get worksheets(): Worksheet[] {
        // return a clone of _worksheets
        return (
            this._worksheets
                .slice(1)
                .filter((worksheet): worksheet is Worksheet => !!worksheet)
                // an undefined orderNo gives NaN
                .sort((a, b) => Number(a.orderNo) - Number(b.orderNo))
        );
    }

    /**
     * The chartsheets (sheets that are a single chart): read from a file (written back as they were read) or created by
     * {@link addChartsheet}, in their order among the sheets: their id, name and state (copies).
     */
    get chartsheets(): Chartsheet[] {
        return this._chartsheets.map(({ model }) => ({ id: model.id, name: model.name, state: model.state ?? "visible" }));
    }

    /**
     * Adds a chartsheet after the other sheets: a sheet that is a single chart. Its ranges name their worksheet
     * (`"Data!B2:B10"`); the values shown before Excel recalculates the chart are read from the cells when the workbook
     * is written.
     * @param name - the name of the sheet tab: the rules of a worksheet name (see the Worksheet name setter)
     * @throws Error for an invalid or duplicate name, invalid chart options (see `worksheet.addChart`; a range without a
     *   worksheet name, an `anchor`), or an invalid `state`
     */
    addChartsheet(name: string, chart: ChartsheetChartOptions, options?: AddChartsheetOptions): Chartsheet {
        const sheetName = checkSheetName(
            name,
            (lowerName) =>
                this._worksheets.some((ws) => ws && ws.name.toLowerCase() === lowerName) ||
                this._chartsheets.some(({ model }) => model.name.toLowerCase() === lowerName)
        );
        const state = chartsheetState(options);
        const chartsheetChart = new ChartsheetChart(this, chart);
        const id = this.nextId;
        const model: ChartsheetModel = { ...createdChartsheetPart(), id, name: sheetName, state, index: 0, charts: [] };
        this._chartsheets.push({ model, orderNo: this.lastOrderNo() + 1, chart: chartsheetChart });
        return { id, name: sheetName, state };
    }

    /**
     * Removes a chartsheet by its exact name, as `getWorksheet` / `removeWorksheet` take a name (nothing for an unknown
     * name). The `activeTab` / `firstSheet` of the workbook views keep their sheets (the next one for the sheet removed).
     */
    removeChartsheet(name: string): void {
        const entry = this._chartsheets.find(({ model }) => model.name === name);
        if (entry) {
            const position = this._sheetOrder().indexOf(entry);
            this._chartsheets = this._chartsheets.filter((chartsheet) => chartsheet !== entry);
            this._sheetRemoved(position);
        }
    }

    /** Calls `iteratee` with each worksheet (in their order) and its id. */
    eachSheet(iteratee: (sheet: Worksheet, id: number) => void): void {
        this.worksheets.forEach((sheet) => {
            iteratee(sheet, sheet.id);
        });
    }

    /** The defined names. */
    get definedNames(): DefinedNames {
        return this._definedNames;
    }

    /** Drops the themes read from a file (the default theme is written). Note: themes are not an exposed feature. */
    clearThemes(): void {
        // Note: themes are not an exposed feature, meddle at your peril!
        this._themes = undefined;
    }

    /**
     * Adds an image (a copy of `image`, without validation).
     * @returns its id, for `worksheet.addImage` / `addBackgroundImage`
     */
    addImage(image: ImageInput): number {
        // TODO:  validation?
        const id = this.media.length;
        this.media.push(Object.assign({}, image, { type: "image" as const }));
        return id;
    }

    /** An image by id, if any. */
    getImage(id: number): MediaModel | undefined {
        return this.media[id];
    }

    /**
     * The model of the workbook (for the xlsx writer): the worksheet models and the settings (by reference); the
     * chartsheets (copies, sharing their parts) with their index among the sheets.
     */
    get model(): DocumentWorkbookModel {
        const { worksheets } = this;
        // the chartsheets in their places among the worksheets (a worksheet of the same orderNo first)
        const chartsheets: ChartsheetModel[] = [];
        let position = 0;
        let w = 0;
        for (const { model, orderNo, chart } of this._chartsheets) {
            while (w < worksheets.length && !(orderNo < Number(worksheets[w]?.orderNo))) {
                w++;
                position++;
            }
            // its tab selected only when it is the active tab of a workbook view (the active tab of the first one when there
            // are no views)
            const index = position++;
            const sheet = withTabSelected(model.sheet, (viewId) => index === (this.views[viewId]?.activeTab ?? 0));
            // a chartsheet created: its chart generated now
            chartsheets.push(chart ? { ...model, sheet, charts: [chart.toModel(model.name)], index } : { ...model, sheet, index });
        }
        return {
            creator: this.creator || "Unknown",
            lastModifiedBy: this.lastModifiedBy || "Unknown",
            lastPrinted: this.lastPrinted,
            created: this.created,
            modified: this.modified,
            properties: this.properties,
            worksheets: worksheets.map((worksheet) => worksheet.model),
            sheets: worksheets.map((ws) => ws.model).filter(Boolean),
            definedNames: this._definedNames.model,
            views: this.views,
            company: this.company,
            manager: this.manager,
            title: this.title,
            subject: this.subject,
            keywords: this.keywords,
            category: this.category,
            description: this.description,
            language: this.language,
            revision: this.revision,
            contentStatus: this.contentStatus,
            themes: this._themes,
            media: this.media,
            pivotTables: this.pivotTables,
            calcProperties: this.calcProperties,
            // only a workbook with chartsheets has the member
            ...(chartsheets.length ? { chartsheets } : {}),
        };
    }

    /**
     * Sets the chartsheets of a model (after its worksheets): a chartsheet whose name is the name of a worksheet or of an
     * earlier chartsheet (case-insensitive) is left out. The worksheets and the chartsheets are numbered (`orderNo`) by
     * their positions among all the sheets.
     */
    _setChartsheets(chartsheets: readonly ChartsheetModel[]): void {
        const { worksheets } = this;
        const names = new Set(worksheets.map((worksheet) => worksheet.name.toLowerCase()));
        const kept = chartsheets.filter((chartsheet) => {
            const name = chartsheet.name.toLowerCase();
            if (names.has(name)) {
                return false;
            }
            names.add(name);
            return true;
        });
        const entries = kept.map((model) => ({ model, orderNo: 0 }));
        const indexed = entries.map((entry) => ({ entry, index: entry.model.index }));
        mergeSheets<Worksheet, (typeof indexed)[number]>(worksheets, indexed).forEach((sheet, orderNo) => {
            if (sheet instanceof Worksheet) {
                sheet.orderNo = orderNo;
            } else {
                sheet.entry.orderNo = orderNo;
            }
        });
        // in their order among the sheets
        this._chartsheets = entries.sort((a, b) => a.orderNo - b.orderNo);
    }

    /**
     * Replaces the workbook with a model (the worksheets are created from their models, ordered as `sheets`).
     * @throws Error for an invalid worksheet name, or when the merges exceed the cap
     */
    set model(value: WorkbookModel) {
        this.creator = value.creator;
        this.lastModifiedBy = value.lastModifiedBy;
        this.lastPrinted = value.lastPrinted;
        this.created = value.created;
        this.modified = value.modified;
        this.company = value.company;
        this.manager = value.manager;
        this.title = value.title;
        this.subject = value.subject;
        this.keywords = value.keywords;
        this.category = value.category;
        this.description = value.description;
        this.language = value.language;
        this.revision = value.revision;
        this.contentStatus = value.contentStatus;

        this.properties = value.properties;
        this.calcProperties = value.calcProperties;
        // the chartsheets first: those of the workbook replaced must not clash with the names of the new worksheets
        this._chartsheets = [];
        this._worksheets = [];
        this._mergedCellCount = 0;
        // index of the first <sheet> of each id, looked up once per worksheet (a findIndex per worksheet was
        // quadratic on a file with many worksheet parts and <sheet> elements)
        const sheetIndexes = new Map<number, number>();
        (value.sheets || []).forEach((sheet, index) => {
            if (!sheetIndexes.has(sheet.id)) {
                sheetIndexes.set(sheet.id, index);
            }
        });
        value.worksheets.forEach((worksheetModel) => {
            const { id, name, state } = worksheetModel;
            const orderNo = value.sheets && (sheetIndexes.has(id) ? sheetIndexes.get(id) : -1);
            const worksheet = (this._worksheets[id] = new Worksheet({
                id,
                name,
                orderNo,
                state,
                workbook: this,
            }));
            worksheet.model = worksheetModel;
        });

        if (value.chartsheets && value.chartsheets.length) {
            this._setChartsheets(value.chartsheets);
        }

        this._definedNames.model = value.definedNames;
        this.views = value.views;
        this._themes = value.themes;
        this.media = value.media || [];
        this.pivotTables = value.pivotTables || [];
    }
}

export default Workbook;
