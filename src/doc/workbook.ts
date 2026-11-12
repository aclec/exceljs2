// A workbook document: worksheets, defined names, images, properties; read and written by `xlsx` and `csv`.
//
// Workbook requirements
//  Load and Save from file and stream
//  Access/Add/Delete individual worksheets
//  Manage String table, Hyperlink table, etc.
//  Manage scaffolding for contained objects to write to/read from

import type { CalculationProperties, WorkbookProperties } from "../../index";
import Worksheet, { type DocumentWorksheetModel, type WorksheetOptions } from "./worksheet";
import DefinedNames from "./defined-names";
import XLSX from "../xlsx/xlsx";
import CSV from "../csv/csv";
import { DEFAULT_MAX_MERGED_CELLS } from "../utils/merge-limits";
import type { PivotTable } from "./pivot-table";
import type { ColorModel, MediaModel, WorkbookModel, WorkbookViewModel } from "../types";

/** The options of `addWorksheet` (index.d.ts `AddWorksheetOptions`). */
export type AddWorksheetOptions = Partial<Omit<WorksheetOptions, "workbook" | "id" | "name" | "orderNo">>;

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

    /** The id of the next worksheet: the first free id (from 1). */
    get nextId(): number {
        // find the next unique spot to add worksheet
        for (let i = 1; i < this._worksheets.length; i++) {
            if (!this._worksheets[i]) {
                return i;
            }
        }
        return this._worksheets.length || 1;
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
            const color = options as ColorModel;
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
            } else if (color.argb || color.theme || color.indexed) {
                // eslint-disable-next-line no-console -- a deprecation notice with its stack trace, as it always was
                console.trace(
                    "tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }"
                );
                worksheetOptions = {
                    properties: {
                        tabColor: color,
                    },
                };
            } else {
                worksheetOptions = options as AddWorksheetOptions;
            }
        }

        const lastOrderNo = this._worksheets.reduce(
            (acc, ws) => (ws && ws.orderNo !== undefined && ws.orderNo > acc ? ws.orderNo : acc),
            0
        );
        const fullOptions: WorksheetOptions = Object.assign({}, worksheetOptions, {
            id,
            name,
            orderNo: lastOrderNo + 1,
            workbook: this,
        });

        const worksheet = new Worksheet(fullOptions);

        this._worksheets[id] = worksheet;
        return worksheet;
    }

    /** Removes a worksheet (a worksheet removed twice does not remove the worksheet that took its id). */
    removeWorksheetEx(worksheet: Worksheet): void {
        // a worksheet removed twice must not remove the worksheet that took its id
        if (this._worksheets[worksheet.id] === worksheet) {
            this._mergedCellCount -= worksheet._mergedCellCount;
            // eslint-disable-next-line @typescript-eslint/no-array-delete -- deliberate: no splice, the indexes are the worksheet ids
            delete this._worksheets[worksheet.id];
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
                // sort never passes the empty entries to the comparator (they go last); an undefined orderNo gives NaN
                .sort((a, b) => Number((a as Worksheet).orderNo) - Number((b as Worksheet).orderNo))
                .filter((worksheet): worksheet is Worksheet => !!worksheet)
        );
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

    /** The model of the workbook (for the xlsx writer): the worksheet models and the settings (by reference). */
    get model(): DocumentWorkbookModel {
        return {
            creator: this.creator || "Unknown",
            lastModifiedBy: this.lastModifiedBy || "Unknown",
            lastPrinted: this.lastPrinted,
            created: this.created,
            modified: this.modified,
            properties: this.properties,
            worksheets: this.worksheets.map((worksheet) => worksheet.model),
            sheets: this.worksheets.map((ws) => ws.model).filter(Boolean),
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
        };
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

        this._definedNames.model = value.definedNames;
        this.views = value.views;
        this._themes = value.themes;
        this.media = value.media || [];
        this.pivotTables = value.pivotTables || [];
    }
}

export default Workbook;
