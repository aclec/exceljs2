import * as fs from "fs";

import StreamBuf from "../../utils/stream-buf";
import { ZipWriter, type ZipEntry, type ZipWriterOptions } from "../../utils/zip-stream";

import RelType from "../../xlsx/rel-type";
import StylesXform from "../../xlsx/xform/style/styles-xform";
import SharedStrings from "../../utils/shared-strings";
import DefinedNames from "../../doc/defined-names";
import { DEFAULT_MAX_MERGED_CELLS, checkMaxMergedCells } from "../../utils/merge-limits";

import CoreXform from "../../xlsx/xform/core/core-xform";
import RelationshipsXform from "../../xlsx/xform/core/relationships-xform";
import ContentTypesXform from "../../xlsx/xform/core/content-types-xform";
import AppXform from "../../xlsx/xform/core/app-xform";
import WorkbookXform from "../../xlsx/xform/book/workbook-xform";
import SharedStringsXform from "../../xlsx/xform/strings/shared-strings-xform";

import WorksheetWriter from "./worksheet-writer";
import { CapturedChart, type CapturedSheet } from "./stream-chart";
import { validateChart, type ChartContext } from "../../doc/chart";
import { checkSheetName, chartsheetState } from "../../doc/sheet-name";
import { ChartPartWriter, prepareChartAnchors, writeDrawingParts } from "../../xlsx/chart-parts";
import {
    createdChartsheetPart,
    prepareChartsheetDrawings,
    prepareChartsheets,
    writeChartsheetParts,
    CHARTSHEET_REL_TYPE,
} from "../../xlsx/chartsheet-parts";

import theme1Xml from "../../xlsx/xml/theme1";
import type {
    AddChartsheetOptions,
    AddWorksheetOptions,
    AutoFilter,
    Chartsheet,
    ChartsheetChartOptions,
    Color,
    Image,
    WorkbookView,
    WorksheetState,
} from "../../../index";
import type { ChartsheetWriteModel, CommentRefsEntry, ContentTypeOverride, DrawingModel, RelationshipModel } from "../../types";

/** The styles manager of the writer: StylesXform (`useStyles`) or its mock (no styles), and the styles part. */
export type WriterStyles = StylesXform;

// the extension of an image ends up in a zip entry name and in XML: alphanumeric only
const IMAGE_EXTENSION = /^[a-z0-9]{1,5}$/i;

/** The options of the {@link WorkbookWriter}. */
export interface WorkbookWriterOptions {
    /** where the archive is written (else `filename`, else an in-memory StreamBuf) */
    stream?: NodeJS.WritableStream;
    /** the path of the file to write, when there is no `stream` */
    filename?: string;
    /** use a shared strings table (smaller file, more memory); default false */
    useSharedStrings?: boolean;
    /** write the styles of the cells; default false */
    useStyles?: boolean;
    /** the zip options (compression level...) */
    zip?: ZipWriterOptions;
    /** the cap on the cells covered by merges in all the worksheets (utils/merge-limits); default 1,048,576 */
    maxMergedCells?: number;
    created?: Date;
    modified?: Date;
    creator?: string;
    lastModifiedBy?: string;
    lastPrinted?: Date;
}

/** The options of {@link WorkbookWriter.addWorksheet}. */
export type AddStreamWorksheetOptions = Partial<AddWorksheetOptions> & {
    /** overrides the workbook's `useSharedStrings` for this worksheet */
    useSharedStrings?: boolean;
    autoFilter?: AutoFilter | null;
    /** @deprecated use `properties.tabColor` */
    tabColor?: Partial<Color>;
};

/** A medium of the writer: the image given to addImage, with its part name (`image0.png`). */
export type WriterMedium = Image & { type: "image"; name: string };

/** A chartsheet added by {@link WorkbookWriter.addChartsheet}: written when the workbook is committed. */
interface WriterChartsheet {
    readonly id: number;
    readonly name: string;
    readonly state: WorksheetState;
    /** Its position among the sheets (worksheets and chartsheets) when it was added. */
    readonly index: number;
    readonly chart: CapturedChart;
}

/** The parts of the charts written (WorkbookWriter commit), for the workbook parts written after them. */
interface WrittenCharts {
    readonly drawings: DrawingModel[];
    readonly chartParts: ContentTypeOverride[];
    readonly writtenChartsheets: ChartsheetWriteModel[];
}

/**
 * The streaming xlsx writer: the worksheets are written to the archive as their rows are committed, so that a large
 * workbook is never held in memory. Once every worksheet is committed (or when {@link commit} commits them), the
 * workbook parts are added and the archive is completed.
 */
class WorkbookWriter {
    /** Document property: the creation date (now by default). */
    created: Date;
    /** Document property: the date of the last change (now by default). */
    modified: Date;
    /** Document property: the author ("ExcelJS" by default). */
    creator: string;
    /** Document property: the author of the last change ("ExcelJS" by default). */
    lastModifiedBy: string;
    /** Document property: the date of the last print. */
    lastPrinted: Date | undefined;
    /** Use shared strings by default in the worksheets (smaller file, more memory). */
    useSharedStrings: boolean;
    /** The shared strings table (used when `useSharedStrings`). */
    sharedStrings: SharedStrings;
    /** The styles registry: StylesXform with `useStyles`, else its mock (no styles). */
    styles: WriterStyles;
    /** The workbook views (windows). */
    views: WorkbookView[];
    /** The cap on the cells covered by merges in all the worksheets (utils/merge-limits). */
    maxMergedCells: number;
    /** The cells covered by merges in all the worksheets so far. */
    _mergedCellCount: number;
    /** The zip options (compression level…). */
    zipOptions: ZipWriterOptions | undefined;
    /** The images added with `addImage`, by image id. */
    media: WriterMedium[];
    /** The comments parts of the worksheets (for the content types). */
    commentRefs: CommentRefsEntry[];
    /** The next block of VML shape ids that the notes of a worksheet can claim (VmlNotesXform). */
    _nextShapeBlock: number;
    /** The archive being written. */
    zip: ZipWriter;
    /** Where the archive is written. */
    stream: NodeJS.WritableStream;
    /** The parts added at once (theme, package relationships). */
    promise: Promise<unknown>;

    private _definedNames: DefinedNames;
    private _worksheets: (WorksheetWriter | undefined)[];
    // the drawings of the worksheets with charts (in the order they are numbered), the chartsheets added, and what is
    // written of them when the workbook is committed
    private _chartDrawings: { readonly name: string; readonly charts: readonly CapturedChart[] }[];
    private _chartsheets: WriterChartsheet[];
    private _writtenCharts: WrittenCharts | undefined;

    /** @throws TypeError for an invalid `maxMergedCells` */
    constructor(options?: WorkbookWriterOptions) {
        options = options || {};

        this.created = options.created || new Date();
        this.modified = options.modified || this.created;
        this.creator = options.creator || "ExcelJS";
        this.lastModifiedBy = options.lastModifiedBy || "ExcelJS";
        this.lastPrinted = options.lastPrinted;

        // using shared strings creates a smaller xlsx file but may use more memory
        this.useSharedStrings = options.useSharedStrings || false;
        this.sharedStrings = new SharedStrings();

        // style manager
        this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock(true);

        // defined names
        this._definedNames = new DefinedNames();

        this._worksheets = [];
        this._chartDrawings = [];
        this._chartsheets = [];
        this._writtenCharts = undefined;
        this.views = [];

        // cap on the cells covered by merges in all the worksheets (utils/merge-limits)
        this.maxMergedCells = options.maxMergedCells !== undefined ? checkMaxMergedCells(options.maxMergedCells) : DEFAULT_MAX_MERGED_CELLS;
        this._mergedCellCount = 0;

        this.zipOptions = options.zip;

        this.media = [];
        this.commentRefs = [];
        // the next block of VML shape ids that the notes of a worksheet can claim (VmlNotesXform)
        this._nextShapeBlock = 1;

        this.zip = new ZipWriter(this.zipOptions);
        if (options.stream) {
            this.stream = options.stream;
        } else if (options.filename) {
            this.stream = fs.createWriteStream(options.filename);
        } else {
            this.stream = new StreamBuf();
        }
        this.zip.pipe(this.stream);

        // these bits can be added right now
        this.promise = Promise.all([this.addThemes(), this.addOfficeRels()]);
    }

    /** The defined names of the workbook. */
    get definedNames(): DefinedNames {
        return this._definedNames;
    }

    /** A zip entry written as a stream (write / end); it emits 'zipped' once all of it is in the archive. */
    _openStream(path: string): ZipEntry {
        return this.zip.openStream(path);
    }

    private _commitWorksheets(): Promise<unknown> {
        const commitWorksheet = function (worksheet: WorksheetWriter | undefined): Promise<void> {
            if (worksheet && !worksheet.committed) {
                return new Promise((resolve) => {
                    worksheet.stream.on("zipped", () => {
                        resolve();
                    });
                    void worksheet.commit();
                });
            }
            return Promise.resolve();
        };
        // if there are any uncommitted worksheets, commit them now and wait
        const promises = this._worksheets.map(commitWorksheet);
        if (promises.length) {
            return Promise.all(promises);
        }
        return Promise.resolve();
    }

    /**
     * Commits the worksheets not committed yet, adds the workbook parts and completes the archive.
     *
     * @returns resolves (with this writer) once the output stream has finished
     * @throws (rejects) on a failure of the zip writer or of the output stream, or an unsupported medium
     */
    async commit(): Promise<this> {
        // commit all worksheets, then add suplimentary files
        await this.promise;
        await this.addMedia();
        // a failing zip (e.g. the output stream failed) would never finish the worksheet entries
        await Promise.race([this._commitWorksheets(), this.zip.finished]);
        // every worksheet is committed: the values of the charts are captured
        this._addCharts();
        await Promise.all([
            this.addContentTypes(),
            this.addApp(),
            this.addCore(),
            this.addSharedStrings(),
            this.addStyles(),
            this.addWorkbookRels(),
        ]);
        await this.addWorkbook();
        return this._finalize();
    }

    /** The next free worksheet id: the first id (from 1) of no worksheet and no chartsheet. */
    get nextId(): number {
        const chartsheetIds = new Set(this._chartsheets.map(({ id }) => id));
        // find the next unique spot to add worksheet
        let i = 1;
        while (this._worksheets[i] || chartsheetIds.has(i)) {
            i++;
        }
        return i;
    }

    /**
     * Adds an image to the workbook (`buffer`, `base64` or `filename`), written by {@link commit}.
     *
     * @returns its id, for `addBackgroundImage`
     * @throws Error for an extension that is not 1 to 5 letters or digits
     */
    addImage(image: Image): number {
        // the extension ends up in a zip entry name and in XML: alphanumeric only
        if (!IMAGE_EXTENSION.test(image.extension)) {
            throw new Error(`Invalid image extension: ${JSON.stringify(image.extension)}`);
        }
        const id = this.media.length;
        const medium: WriterMedium = Object.assign({}, image, { type: "image" as const, name: `image${id}.${image.extension}` });
        this.media.push(medium);
        return id;
    }

    /** An image added with {@link addImage}. */
    getImage(id: number): WriterMedium | undefined {
        return this.media[id];
    }

    /**
     * Adds a worksheet, written as its rows are committed. It's possible to add a worksheet with different than default
     * shared string handling; in fact, it's even possible to switch it mid-sheet.
     *
     * @param name - default `sheet<id>`
     */
    addWorksheet(name?: string, options?: AddStreamWorksheetOptions): WorksheetWriter {
        options = options || {};
        const useSharedStrings = options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

        if (options.tabColor) {
            // eslint-disable-next-line no-console -- deprecation notice (with the caller's stack), kept from ExcelJS
            console.trace("tabColor option has moved to { properties: tabColor: {...} }");
            options.properties = Object.assign(
                {
                    tabColor: options.tabColor,
                },
                options.properties
            );
        }

        const id = this.nextId;
        name = name || `sheet${id}`;
        // a worksheet may not have the name of a chartsheet (case-insensitive)
        const lowerName = String(name).toLowerCase();
        if (this._chartsheets.some((chartsheet) => chartsheet.name.toLowerCase() === lowerName)) {
            throw new Error(`Worksheet name already exists: ${name}`);
        }

        const worksheet = new WorksheetWriter({
            id,
            name,
            workbook: this,
            useSharedStrings,
            properties: options.properties,
            state: options.state,
            pageSetup: options.pageSetup,
            views: options.views,
            autoFilter: options.autoFilter,
            headerFooter: options.headerFooter,
        });

        this._worksheets[id] = worksheet;
        return worksheet;
    }

    /**
     * Adds a chartsheet after the sheets added so far: a sheet that is a single chart, with the options of the document
     * `addChartsheet` (see index.d.ts `ChartsheetChartOptions`). The values shown before Excel recalculates the chart
     * are captured from the cells as their rows are written; a series with a row committed before is written without
     * cached values. The chartsheet is written when the workbook is committed.
     *
     * @param name - the name of the sheet tab: the rules of a worksheet name (unique among all the sheets,
     *   case-insensitive)
     * @returns the chartsheet added
     * @throws Error for an invalid or duplicate name, invalid chart options (a range without a worksheet name, an
     *   `anchor`...), or an invalid `state`
     */
    addChartsheet(name: string, chart: ChartsheetChartOptions, options?: AddChartsheetOptions): Chartsheet {
        const sheetName = checkSheetName(
            name,
            (lowerName) =>
                this._worksheets.some((worksheet) => worksheet && worksheet.name.toLowerCase() === lowerName) ||
                this._chartsheets.some((chartsheet) => chartsheet.name.toLowerCase() === lowerName)
        );
        const state = chartsheetState(options);
        const definition = validateChart<CapturedSheet>(chart, this._chartContext(undefined));
        const id = this.nextId;
        const index = this._worksheets.filter(Boolean).length + this._chartsheets.length;
        this._chartsheets.push({ id, name: sheetName, state, index, chart: new CapturedChart(definition) });
        return { id, name: sheetName, state };
    }

    /**
     * The context of the validation of a chart (src/doc/chart.ts): the worksheets of this writer.
     * @param sheet - the worksheet of the chart, undefined for the chart of a chartsheet
     */
    _chartContext(sheet: WorksheetWriter | undefined): ChartContext<CapturedSheet> {
        return {
            sheet,
            cellSizes: sheet && {
                _findColumn: (col) => sheet.columns?.[col - 1],
                findRow: (row) => sheet.findRow(row),
            },
            findSheet: (name) => {
                const lowerName = name.toLowerCase();
                return this._worksheets.find((worksheet) => worksheet && worksheet.name.toLowerCase() === lowerName);
            },
        };
    }

    /**
     * Registers the charts of a worksheet being committed (written when the workbook is committed).
     * @returns the name of their drawing (`drawing1`)
     */
    _addChartDrawing(charts: readonly CapturedChart[]): string {
        const name = `drawing${this._chartDrawings.length + 1}`;
        this._chartDrawings.push({ name, charts });
        return name;
    }

    /**
     * Writes the drawings of the worksheets, the chartsheets and the chart parts (every worksheet committed: the values
     * of the charts are captured), through the code of the document writer (chart-parts.ts, chartsheet-parts.ts).
     */
    private _addCharts(): void {
        if (!this._chartDrawings.length && !this._chartsheets.length) {
            return;
        }
        const registry = new ChartPartWriter();
        const drawings: DrawingModel[] = this._chartDrawings.map(({ name, charts }) => {
            const rels: RelationshipModel[] = [];
            const models = charts.map((chart, index) => chart.toModel(`Chart ${index + 1}`));
            return { name, anchors: [], rels, chartAnchors: prepareChartAnchors(models, name, rels, registry) };
        });
        const worksheets = this._worksheets.filter((worksheet) => worksheet !== undefined);
        const writtenChartsheets = prepareChartsheets(
            worksheets,
            this._chartsheets.map(({ id, name, state, index, chart }) => ({
                ...createdChartsheetPart(),
                id,
                name,
                state,
                index,
                charts: [chart.toModel("Chart 1")],
            }))
        );
        // the drawings of the chartsheets are numbered after those of the worksheets
        prepareChartsheetDrawings(writtenChartsheets, { drawingsCount: drawings.length, drawings }, registry);
        const chartParts = registry.complete();
        writeChartsheetParts(this.zip, writtenChartsheets);
        writeDrawingParts(this.zip, drawings);
        registry.write(this.zip);
        this._writtenCharts = { drawings, chartParts, writtenChartsheets };
        this._chartDrawings = [];
    }

    /** The first worksheet (no argument), the worksheet with this id, or the one with this name. */
    getWorksheet(id?: number | string): WorksheetWriter | undefined {
        if (id === undefined) {
            return this._worksheets.find(() => true);
        }
        if (typeof id === "number") {
            return this._worksheets[id];
        }
        if (typeof id === "string") {
            return this._worksheets.find((worksheet) => worksheet && worksheet.name === id);
        }
        return undefined;
    }

    /** Adds the styles part. */
    addStyles(): Promise<void> {
        return new Promise((resolve) => {
            this.zip.append(this.styles.xml, { name: "xl/styles.xml" });
            resolve();
        });
    }

    /** Adds the theme part. */
    addThemes(): Promise<void> {
        return new Promise((resolve) => {
            this.zip.append(theme1Xml, { name: "xl/theme/theme1.xml" });
            resolve();
        });
    }

    /** Adds the package relationships (`_rels/.rels`). */
    addOfficeRels(): Promise<void> {
        return new Promise((resolve) => {
            const xform = new RelationshipsXform();
            const xml = xform.toXml([
                { Id: "rId1", Type: RelType.OfficeDocument, Target: "xl/workbook.xml" },
                { Id: "rId2", Type: RelType.CoreProperties, Target: "docProps/core.xml" },
                { Id: "rId3", Type: RelType.ExtenderProperties, Target: "docProps/app.xml" },
            ]);
            this.zip.append(xml, { name: "/_rels/.rels" });
            resolve();
        });
    }

    /** Adds `[Content_Types].xml`. */
    addContentTypes(): Promise<void> {
        return new Promise((resolve) => {
            const model = {
                worksheets: this._worksheets.filter((worksheet) => worksheet !== undefined),
                sharedStrings: this.sharedStrings,
                commentRefs: this.commentRefs,
                media: this.media,
                ...this._writtenCharts,
            };
            const xform = new ContentTypesXform();
            const xml = xform.toXml(model);
            this.zip.append(xml, { name: "[Content_Types].xml" });
            resolve();
        });
    }

    /**
     * Adds the media parts (an image file is read now, its entry keeping its place in the archive).
     *
     * @throws Error for a medium without data (synchronously)
     */
    addMedia(): Promise<void[]> {
        // the files being read
        const reads: Promise<void>[] = [];
        for (const medium of this.media) {
            const filename = `xl/media/${medium.name}`;
            if (medium.type !== "image") {
                throw new Error("Unsupported media");
            } else if (medium.filename) {
                // created now to keep its place in the archive, written once read
                const entry = this.zip.openStream(filename);
                reads.push(fs.promises.readFile(medium.filename).then((data) => entry.end(data)));
            } else if (medium.buffer) {
                this.zip.append(medium.buffer, { name: filename });
            } else if (medium.base64) {
                const dataimg64 = medium.base64;
                const content = dataimg64.substring(dataimg64.indexOf(",") + 1);
                this.zip.append(content, { name: filename, base64: true });
            } else {
                throw new Error("Unsupported media");
            }
        }
        return Promise.all(reads);
    }

    /** Adds the extended properties (`docProps/app.xml`). */
    addApp(): Promise<void> {
        return new Promise((resolve) => {
            const model = {
                worksheets: this._worksheets.filter((worksheet) => worksheet !== undefined),
                writtenChartsheets: this._writtenCharts?.writtenChartsheets,
            };
            const xform = new AppXform();
            const xml = xform.toXml(model);
            this.zip.append(xml, { name: "docProps/app.xml" });
            resolve();
        });
    }

    /** Adds the core properties (`docProps/core.xml`), from the properties of this writer. */
    addCore(): Promise<void> {
        return new Promise((resolve) => {
            const coreXform = new CoreXform();
            const xml = coreXform.toXml(this);
            this.zip.append(xml, { name: "docProps/core.xml" });
            resolve();
        });
    }

    /** Adds the shared strings part, when a worksheet used it. */
    addSharedStrings(): Promise<void> {
        if (this.sharedStrings.count) {
            return new Promise((resolve) => {
                const sharedStringsXform = new SharedStringsXform();
                // count is the number of references to the shared strings (ECMA-376 18.4.9)
                const { values, totalRefs } = this.sharedStrings;
                const xml = sharedStringsXform.toXml({ values, count: totalRefs });
                this.zip.append(xml, { name: "/xl/sharedStrings.xml" });
                resolve();
            });
        }
        return Promise.resolve();
    }

    /** Adds the workbook relationships; sets the `rId` of each worksheet. */
    addWorkbookRels(): Promise<void> {
        let count = 1;
        const relationships: RelationshipModel[] = [
            { Id: `rId${count++}`, Type: RelType.Styles, Target: "styles.xml" },
            { Id: `rId${count++}`, Type: RelType.Theme, Target: "theme/theme1.xml" },
        ];
        if (this.sharedStrings.count) {
            relationships.push({
                Id: `rId${count++}`,
                Type: RelType.SharedStrings,
                Target: "sharedStrings.xml",
            });
        }
        for (const worksheet of this._worksheets) {
            if (worksheet) {
                const rId = (worksheet.rId = `rId${count++}`);
                relationships.push({
                    Id: rId,
                    Type: RelType.Worksheet,
                    Target: `worksheets/sheet${worksheet.id}.xml`,
                });
            }
        }
        for (const chartsheet of this._writtenCharts?.writtenChartsheets ?? []) {
            const rId = (chartsheet.rId = `rId${count++}`);
            relationships.push({ Id: rId, Type: CHARTSHEET_REL_TYPE, Target: `chartsheets/sheet${chartsheet.partNo}.xml` });
        }
        return new Promise((resolve) => {
            const xform = new RelationshipsXform();
            const xml = xform.toXml(relationships);
            this.zip.append(xml, { name: "/xl/_rels/workbook.xml.rels" });
            resolve();
        });
    }

    /** Adds the workbook part (`xl/workbook.xml`). */
    addWorkbook(): Promise<void> {
        const { zip } = this;
        const model = {
            worksheets: this._worksheets.filter((worksheet) => worksheet !== undefined),
            writtenChartsheets: this._writtenCharts?.writtenChartsheets,
            definedNames: this._definedNames.model,
            views: this.views,
            properties: {},
            calcProperties: {},
        };

        return new Promise((resolve) => {
            const xform = new WorkbookXform();
            xform.prepare(model);
            zip.append(xform.toXml(model), { name: "/xl/workbook.xml" });
            resolve();
        });
    }

    private _finalize(): Promise<this> {
        return new Promise((resolve, reject) => {
            this.stream.on("error", reject);
            this.stream.on("finish", () => {
                resolve(this);
            });
            // rejects on any failure of the zip writer or of the output stream
            this.zip.finalize().catch(reject);
        });
    }
}

export default WorkbookWriter;
