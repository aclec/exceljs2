import RelType from "../../xlsx/rel-type";

import colCache from "../../utils/col-cache";
import Encryptor from "../../utils/encryptor";
import { checkMergedCells } from "../../utils/merge-limits";
import { eachEntry, someEntry } from "../../utils/sparse-array";
import Range, { type RangeArgs } from "../../doc/range";
import StringBuf from "../../utils/string-buf";

import Row from "../../doc/row";
import Column from "../../doc/column";

import SheetRelsWriter from "./sheet-rels-writer";
import SheetCommentsWriter from "./sheet-comments-writer";
import DataValidations from "../../doc/data-validations";
import type Cell from "../../doc/cell";
import type { ZipEntry } from "../../utils/zip-stream";
import type {
    AutoFilter,
    ConditionalFormattingOptions,
    HeaderFooter,
    PageSetup,
    RowBreak,
    RowValues,
    WorksheetProperties,
    WorksheetProtection,
    WorksheetState,
    WorksheetView,
} from "../../../index";
import type {
    CellPrepareOptions,
    CommentRefModel,
    FormulaCellModel,
    HeaderFooterModel,
    ListItemOptions,
    MergeCellModel,
    PageSetupModel,
    SheetProtectionModel,
    StyleRegistry,
} from "../../types";
import type WorkbookWriter from "./workbook-writer";

// ============================================================================================
// Xforms
import ListXform from "../../xlsx/xform/list-xform";
import DataValidationsXform from "../../xlsx/xform/sheet/data-validations-xform";
import SheetPropertiesXform from "../../xlsx/xform/sheet/sheet-properties-xform";
import SheetFormatPropertiesXform from "../../xlsx/xform/sheet/sheet-format-properties-xform";
import ColXform from "../../xlsx/xform/sheet/col-xform";
import RowXform from "../../xlsx/xform/sheet/row-xform";
import HyperlinkXform from "../../xlsx/xform/sheet/hyperlink-xform";
import SheetViewXform from "../../xlsx/xform/sheet/sheet-view-xform";
import SheetProtectionXform from "../../xlsx/xform/sheet/sheet-protection-xform";
import PageMarginsXform from "../../xlsx/xform/sheet/page-margins-xform";
import PageSetupXform from "../../xlsx/xform/sheet/page-setup-xform";
import PrintOptionsXform from "../../xlsx/xform/sheet/print-options-xform";
import AutoFilterXform from "../../xlsx/xform/sheet/auto-filter-xform";
import PictureXform from "../../xlsx/xform/sheet/picture-xform";
import ConditionalFormattingsXform from "../../xlsx/xform/sheet/cf/conditional-formattings-xform";
import HeaderFooterXform from "../../xlsx/xform/sheet/header-footer-xform";
import RowBreaksXform from "../../xlsx/xform/sheet/row-breaks-xform";
import ExtListXform from "../../xlsx/xform/sheet/ext-lst-xform";

const xmlBuffer = new StringBuf();

// since prepare and render are functional, we can use singletons
const xform = {
    dataValidations: new DataValidationsXform(),
    sheetProperties: new SheetPropertiesXform(),
    sheetFormatProperties: new SheetFormatPropertiesXform(),
    columns: new ListXform({ tag: "cols", childXform: new ColXform() }),
    row: new RowXform(),
    hyperlinks: new ListXform({ tag: "hyperlinks", childXform: new HyperlinkXform() }),
    sheetViews: new ListXform({ tag: "sheetViews", childXform: new SheetViewXform() }),
    sheetProtection: new SheetProtectionXform(),
    pageMargins: new PageMarginsXform(),
    pageSeteup: new PageSetupXform(),
    printOptions: new PrintOptionsXform(),
    autoFilter: new AutoFilterXform(),
    picture: new PictureXform(),
    conditionalFormattings: new ConditionalFormattingsXform(),
    headerFooter: new HeaderFooterXform(),
    rowBreaks: new RowBreaksXform(),
    extLst: new ExtListXform(),
};

/** A column definition (`worksheet.columns = [...]`): index.d.ts `Column` properties, and multi-row `headers`. */
export type ColumnDefinition = Partial<Column> & { headers?: readonly unknown[] };

/** The sheet properties of a streamed worksheet: the defaults below, and those given. */
export type WorksheetWriterProperties = Partial<WorksheetProperties> &
    Pick<WorksheetProperties, "defaultRowHeight" | "dyDescent" | "outlineLevelCol" | "outlineLevelRow">;

/** The header and footer of a streamed worksheet: `null` by default. */
export type WorksheetWriterHeaderFooter = HeaderFooterModel;

/** The page setup of a streamed worksheet: the defaults below (`rowBreaks` / `colBreaks` null), and those given. */
export type WorksheetWriterPageSetup = PageSetupModel & { rowBreaks?: unknown; colBreaks?: unknown };

/** The arguments of the {@link WorksheetWriter} constructor (WorkbookWriter.addWorksheet). */
export interface WorksheetWriterOptions {
    /** the number of the worksheet in the workbook */
    id: number;
    name?: string;
    workbook: WorkbookWriter;
    useSharedStrings?: boolean;
    properties?: Partial<WorksheetProperties>;
    state?: WorksheetState;
    pageSetup?: Partial<PageSetup>;
    views?: Partial<WorksheetView>[];
    autoFilter?: AutoFilter | null;
    headerFooter?: Partial<HeaderFooter>;
}

// the error of an operation on the rows of a committed worksheet
function committedError(): Error {
    return new Error("Invalid Operation: the worksheet has been committed");
}

/**
 * A worksheet of the streaming {@link WorkbookWriter}: the rows are written (and dropped) as they are committed, in
 * order; the rest of the worksheet (merges, validations, hyperlinks, page setup...) when the worksheet is committed. The
 * XML elements follow the ECMA-376 CT_Worksheet sequence, as WorkSheetXform.render.
 */
class WorksheetWriter {
    /** The number of the worksheet in the workbook. */
    readonly id: number;
    /** The worksheet name. */
    name: string;
    /** The visibility (`"visible"`, `"hidden"`, `"veryHidden"`). */
    state: WorksheetState;
    /** True once {@link commit} was called: no row can be added. */
    committed: boolean;
    /** The data validations of the cells. */
    dataValidations: DataValidations;
    /** The conditional formattings (`addConditionalFormatting`). */
    conditionalFormatting: ConditionalFormattingOptions[];
    /** The manual page breaks (`row.addPageBreak`). */
    rowBreaks: RowBreak[];
    /** The worksheet properties (tab colour, outline, default row height…). */
    properties: WorksheetWriterProperties;
    /** Headers and footers of the printed pages (`null` by default). */
    headerFooter: WorksheetWriterHeaderFooter;
    /** For all things printing (margins, orientation, fit to page…). */
    pageSetup: WorksheetWriterPageSetup;
    /** Use the shared strings table (smaller file, more memory). */
    useSharedStrings: boolean;
    /** True once a row with notes was written. */
    hasComments: boolean;
    /** The auto filter of the worksheet, or null. */
    autoFilter: AutoFilter | null;
    /** The protection of the worksheet, or null when it is not protected. */
    sheetProtection: SheetProtectionModel | null;
    /** True once `<sheetData>` was opened. */
    startedData: boolean;
    /** Set to `[]` by the comments writer when the first note is written. */
    comments?: CommentRefModel[];
    /** The relationship id of the worksheet in the workbook (WorkbookWriter.addWorkbookRels). */
    rId?: string;

    // rows are stored here while they need to be worked on, indexed by number - this._rowsBase.
    // when they are committed, they will be deleted; null once the worksheet is committed
    private _rows: (Row | undefined)[] | null;
    private _rowsBase: number;
    // the number of rows in this._rows
    private _rowsCount: number;
    // column definitions
    private _columns: Column[] | null;
    private _headerRowCount?: number;
    // column keys (addRow convenience): key ==> this._columns index
    // no prototype: keys come from data (e.g. CSV headers) and may be '__proto__' or 'constructor'
    private _keys: Record<string, Column>;
    // the merged ranges; `add` is the merges collector of the cells' prepare (the merges are written from this list)
    private _merges: Range[] & { add(cell: MergeCellModel): void };
    private _mergedCellCount: number;
    private _sheetRelsWriter: SheetRelsWriter;
    private _sheetCommentsWriter: SheetCommentsWriter;
    private _dimensions: Range;
    // first uncommitted row
    private _rowZero: number;
    // shared formula masters by cell address, and the next shared formula index (cell-xform prepare)
    private _formulae: Record<string, FormulaCellModel>;
    private _siFormulae: number;
    private _workbook: WorkbookWriter;
    private _views: Partial<WorksheetView>[];
    private _media: unknown[];
    private _background?: { imageId?: number; rId?: string };
    private _stream?: ZipEntry;

    constructor(options: WorksheetWriterOptions) {
        // in a workbook, each sheet will have a number
        this.id = options.id;

        // and a name
        this.name = options.name || `Sheet${this.id}`;

        // add a state
        this.state = options.state || "visible";

        this._rows = [];
        this._rowsBase = 1;
        this._rowsCount = 0;

        this._columns = null;

        this._keys = Object.create(null) as Record<string, Column>;

        // keep a record of all merges
        this._merges = Object.assign([] as Range[], {
            add() {}, // ignore cell instruction
        });
        this._mergedCellCount = 0;

        // keep record of all hyperlinks
        this._sheetRelsWriter = new SheetRelsWriter(options);

        this._sheetCommentsWriter = new SheetCommentsWriter(this, this._sheetRelsWriter, options);

        // keep a record of dimensions
        this._dimensions = new Range();

        this._rowZero = 1;

        // committed flag
        this.committed = false;

        // for data validations
        this.dataValidations = new DataValidations();

        // for sharing formulae; keyed by cell addresses (no prototype: a sharedFormula "constructor" finds no master)
        this._formulae = Object.create(null) as Record<string, FormulaCellModel>;
        this._siFormulae = 0;

        // keep a record of conditionalFormattings
        this.conditionalFormatting = [];

        // keep a record of all row and column pageBreaks
        this.rowBreaks = [];

        // for default row height, outline levels, etc
        this.properties = Object.assign(
            {},
            {
                defaultRowHeight: 15,
                dyDescent: 55,
                outlineLevelCol: 0,
                outlineLevelRow: 0,
            },
            options.properties
        );

        this.headerFooter = Object.assign(
            {},
            {
                differentFirst: false,
                differentOddEven: false,
                oddHeader: null,
                oddFooter: null,
                evenHeader: null,
                evenFooter: null,
                firstHeader: null,
                firstFooter: null,
            },
            options.headerFooter
        );

        // for all things printing
        this.pageSetup = Object.assign(
            {},
            {
                margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                orientation: "portrait" as const,
                horizontalDpi: 4294967295,
                verticalDpi: 4294967295,
                fitToPage: !!(
                    options.pageSetup &&
                    (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
                    !options.pageSetup.scale
                ),
                pageOrder: "downThenOver" as const,
                blackAndWhite: false,
                draft: false,
                cellComments: "None" as const,
                errors: "displayed" as const,
                scale: 100,
                fitToWidth: 1,
                fitToHeight: 1,
                paperSize: undefined,
                showRowColHeaders: false,
                showGridLines: false,
                horizontalCentered: false,
                verticalCentered: false,
                rowBreaks: null,
                colBreaks: null,
            },
            options.pageSetup
        );

        // using shared strings creates a smaller xlsx file but may use more memory
        this.useSharedStrings = options.useSharedStrings || false;

        this._workbook = options.workbook;

        this.hasComments = false;

        // views
        this._views = options.views || [];

        // auto filter
        this.autoFilter = options.autoFilter || null;

        this._media = [];

        // worksheet protection
        this.sheetProtection = null;

        // start writing to stream now
        this._writeOpenWorksheet();

        this.startedData = false;
    }

    /** The workbook writer. */
    get workbook(): WorkbookWriter {
        return this._workbook;
    }

    /** The zip entry of the worksheet part, opened on first use. */
    get stream(): ZipEntry {
        if (!this._stream) {
            this._stream = this._workbook._openStream(`/xl/worksheets/sheet${this.id}.xml`);
        }
        return this._stream;
    }

    /**
     * Not a valid operation for a streaming writer (even though some streamers might be able to, it's a bad idea).
     *
     * @throws Error always
     */
    destroy(): never {
        throw new Error("Invalid Operation: destroy");
    }

    /**
     * Writes the rest of the worksheet: the uncommitted rows, then the elements after `<sheetData>`. Rows can no longer
     * be added.
     *
     * @returns a promise to await before writing more (backpressure, see {@link _commitRow})
     */
    commit(): Promise<void> {
        if (this.committed) {
            return this._workbook.zip.drained();
        }
        // commit all rows
        eachEntry(
            this._liveRows,
            (cRow) => {
                // write the row to the stream
                this._writeRow(cRow);
            },
            this._rowZero - this._rowsBase
        );

        // we _cannot_ accept new rows from now on
        this._rows = null;

        if (!this.startedData) {
            this._writeOpenSheetData();
        }
        this._writeCloseSheetData();

        // for some reason, Excel can't handle dimensions at the bottom of the file
        // this._writeDimensions();

        // the rest must follow the ECMA-376 CT_Worksheet sequence, like WorkSheetXform.render
        // (legacyDrawing before picture and extLst), otherwise Excel reports the file as corrupt
        this._writeSheetProtection();
        this._writeAutoFilter();
        this._writeMergeCells();
        this._writeConditionalFormatting();
        this._writeDataValidations();
        this._writeHyperlinks();
        this._writePrintOptions();
        this._writePageMargins();
        this._writePageSetup();
        this._writeHeaderFooter();
        this._writeRowBreaks();
        // Legacy Data tag for comments
        this._writeLegacyData();
        this._writeBackground();
        this._writeExtLst();

        this._writeCloseWorksheet();
        // signal end of stream to workbook
        this.stream.end();

        this._sheetCommentsWriter.commit();
        // also commit the hyperlinks if any
        this._sheetRelsWriter.commit();

        this.committed = true;
        return this._workbook.zip.drained();
    }

    /** The dimensions of the rows written so far. */
    get dimensions(): Range {
        return this._dimensions;
    }

    /** The sheet views (`views` option). */
    get views(): Partial<WorksheetView>[] {
        return this._views;
    }

    // =========================================================================
    // Columns

    /** The columns defined or used, or `null` when there is none. */
    get columns(): Column[] | null {
        return this._columns;
    }

    /** Defines the columns. Note: any headers defined will overwrite existing values. */
    set columns(value: readonly ColumnDefinition[]) {
        // calculate max header row count
        this._headerRowCount = value.reduce((pv, cv) => {
            const headerCount = (cv.header && 1) || (cv.headers && cv.headers.length) || 0;
            return Math.max(pv, headerCount);
        }, 0);

        // construct Column objects
        let count = 1;
        const columns: Column[] = (this._columns = []);
        value.forEach((defn) => {
            const column = new Column(this, count++, false);
            columns.push(column);
            column.defn = defn;
        });
    }

    /** The column with this key (Column.key), if any. */
    getColumnKey(key: string): Column | undefined {
        return this._keys[key];
    }

    /** Registers the column of a key (Column.key setter). */
    setColumnKey(key: string, value: Column): void {
        this._keys[key] = value;
    }

    /** Unregisters a key (Column.key setter). */
    deleteColumnKey(key: string): void {
        delete this._keys[key];
    }

    /** Calls `f` for each keyed column. */
    eachColumnKey(f: (column: Column, key: string) => void): void {
        for (const key of Object.keys(this._keys)) {
            f(this._keys[key]!, key);
        }
    }

    /**
     * A column by key, letter or number (1-based); it and any gap before it are created if needed.
     *
     * @returns `undefined` for a number below 1
     * @throws Error for letters that are not a column
     */
    getColumn(c: number | string): Column | undefined {
        if (typeof c === "string") {
            // if it matches a key'd column, return that
            const col = this._keys[c];
            if (col) return col;

            // otherwise, assume letter
            c = colCache.l2n(c);
        }
        if (!this._columns) {
            this._columns = [];
        }
        if (c > this._columns.length) {
            let n = this._columns.length + 1;
            while (n <= c) {
                this._columns.push(new Column(this, n++));
            }
        }
        return this._columns[c - 1];
    }

    // =========================================================================
    // Rows

    // the uncommitted rows
    private get _liveRows(): (Row | undefined)[] {
        if (!this._rows) {
            throw committedError();
        }
        return this._rows;
    }

    private get _nextRow(): number {
        return Math.max(this._rowZero, this._rowsBase + this._liveRows.length);
    }

    /**
     * Iterates over the uncommitted rows with values (with `includeEmpty`, over every uncommitted row, created if
     * needed).
     */
    eachRow(iteratee: (row: Row, rowNumber: number) => void): void;
    eachRow(options: { includeEmpty?: boolean } | undefined, iteratee: (row: Row, rowNumber: number) => void): void;
    eachRow(
        options: { includeEmpty?: boolean } | ((row: Row, rowNumber: number) => void) | undefined,
        iteratee?: (row: Row, rowNumber: number) => void
    ): void {
        if (!iteratee) {
            iteratee = options as (row: Row, rowNumber: number) => void;
            options = undefined;
        }
        if (options && typeof options === "object" && options.includeEmpty) {
            const n = this._nextRow;
            for (let i = this._rowZero; i < n; i++) {
                iteratee(this.getRow(i), i);
            }
        } else {
            const callback = iteratee;
            eachEntry(
                this._liveRows,
                (row) => {
                    if (row.hasValues) {
                        callback(row, row.number);
                    }
                },
                this._rowZero - this._rowsBase
            );
        }
    }

    /**
     * Writes the rows up to `cRow` (rows are written in order) and drops them (Row.commit).
     *
     * @returns a promise that resolves once the output can take more data: awaiting it keeps memory bounded. Without
     * waiting, written data piles up in memory until the event loop runs (the zip compresses asynchronously).
     */
    _commitRow(cRow: Row): Promise<void> {
        // since rows must be written in order, we commit all rows up till and including cRow.
        // The committed rows are deleted, not shifted out: shift takes time proportional to the length of a
        // sparse array, so a row after a hole of 10,000 rows took 2 s to commit.
        const rows = this._liveRows;
        someEntry(
            rows,
            (row, index) => {
                // removed before it is written: a row whose writing fails is not written again. A hole, not undefined:
                // the iteration of a sparse array visits the keys of the array (sparse-array.ts)
                // eslint-disable-next-line @typescript-eslint/no-array-delete -- see above
                delete rows[index];
                this._rowsCount--;
                this._rowZero = row.number + 1;
                this._writeRow(row);
                return row.number === cRow.number;
            },
            this._rowZero - this._rowsBase
        );
        // drop the committed indexes once they outnumber the rows: time proportional to the rows committed
        const committed = this._rowZero - this._rowsBase;
        if (committed > this._rowsCount + 32) {
            const remaining: (Row | undefined)[] = [];
            eachEntry(
                rows,
                (row, index) => {
                    remaining[index - committed] = row;
                },
                committed
            );
            this._rows = remaining;
            this._rowsBase = this._rowZero;
        }
        return this._workbook.zip.drained();
    }

    /** The last uncommitted row, if any. */
    get lastRow(): Row | undefined {
        const rows = this._liveRows;
        const index = rows.length - 1;
        return index >= this._rowZero - this._rowsBase ? rows[index] : undefined;
    }

    /** A row (1-based number) if it exists and is not committed. */
    findRow(rowNumber: number): Row | undefined {
        return rowNumber < this._rowZero ? undefined : this._liveRows[rowNumber - this._rowsBase];
    }

    /**
     * A row (1-based number), created if needed.
     *
     * @throws Error when the row was committed
     */
    getRow(rowNumber: number): Row {
        // may fail if rows have been comitted
        if (rowNumber < this._rowZero) {
            throw new Error("Out of bounds: this row has been committed");
        }
        const rows = this._liveRows;
        const index = rowNumber - this._rowsBase;
        let row = rows[index];
        if (!row) {
            rows[index] = row = new Row(this, rowNumber);
            this._rowsCount++;
        }
        return row;
    }

    /** Adds a row after the last one, with values (an array by column number, or an object by column key). */
    addRow(value: RowValues): Row {
        const rows = this._liveRows;
        const row = new Row(this, this._nextRow);
        rows[row.number - this._rowsBase] = row;
        this._rowsCount++;
        row.values = value;
        return row;
    }

    // ================================================================================
    // Cells

    /** The cell at an address (`B3`) or at (row, col), if it exists. */
    findCell(address: string): Cell | undefined;
    findCell(row: number, col: number): Cell | undefined;
    findCell(r: number | string, c?: number): Cell | undefined {
        const address = typeof c === "number" ? colCache.getAddress(r as number, c) : colCache.getAddress(r as string);
        const row = this.findRow(address.row);
        return row ? row.findCell(address.col) : undefined;
    }

    /**
     * The cell at an address (`B3`) or at (row, col), created if needed.
     *
     * @throws Error when its row was committed
     */
    getCell(address: string): Cell;
    getCell(row: number, col: number): Cell;
    getCell(r: number | string, c?: number): Cell {
        const address = typeof c === "number" ? colCache.getAddress(r as number, c) : colCache.getAddress(r as string);
        const row = this.getRow(address.row);
        return row.getCellEx(address);
    }

    /**
     * Merges a range of cells (same arguments as `new Range(...)`): every cell but the top-left one is merged into it.
     *
     * @throws Error when a cell is already merged, a row was committed, or the merged cells exceed the limits
     * (utils/merge-limits)
     */
    mergeCells(...cells: RangeArgs): void {
        // may fail if rows have been comitted
        const dimensions = new Range(cells);

        // every cell of a merge is materialised: same caps as the document model (utils/merge-limits)
        const workbook = this._workbook;
        const mergedCellCount = this._mergedCellCount + dimensions.count;
        checkMergedCells(
            `Cannot merge ${dimensions.range}`,
            mergedCellCount,
            workbook._mergedCellCount + dimensions.count,
            workbook.maxMergedCells
        );

        // check cells aren't already merged
        for (const merge of this._merges) {
            if (merge.intersects(dimensions)) {
                throw new Error("Cannot merge already merged cells");
            }
        }

        // apply merge
        const master = this.getCell(dimensions.top, dimensions.left);
        for (let i = dimensions.top; i <= dimensions.bottom; i++) {
            for (let j = dimensions.left; j <= dimensions.right; j++) {
                if (i > dimensions.top || j > dimensions.left) {
                    this.getCell(i, j).merge(master);
                }
            }
        }

        // index merge
        this._merges.push(dimensions);
        this._mergedCellCount = mergedCellCount;
        workbook._mergedCellCount += dimensions.count;
    }

    // ===========================================================================
    // Conditional Formatting

    /** Adds conditional formatting rules, written when the worksheet is committed. */
    addConditionalFormatting(cf: ConditionalFormattingOptions): void {
        this.conditionalFormatting.push(cf);
    }

    /** Removes the conditional formattings at an index, those a filter rejects, or all of them. */
    removeConditionalFormatting(filter?: number | ((value: ConditionalFormattingOptions, index: number) => unknown)): void {
        if (typeof filter === "number") {
            this.conditionalFormatting.splice(filter, 1);
        } else if (filter instanceof Function) {
            this.conditionalFormatting = this.conditionalFormatting.filter(filter);
        } else {
            this.conditionalFormatting = [];
        }
    }

    // =========================================================================

    /** Sets the background image (an id of WorkbookWriter.addImage). */
    addBackgroundImage(imageId: number): void {
        this._background = {
            imageId,
        };
    }

    /** The id of the background image, if any. */
    getBackgroundImageId(): number | undefined {
        return this._background && this._background.imageId;
    }

    // =========================================================================
    // Worksheet Protection

    /**
     * Protects the worksheet, with a password hashed with SHA-512 (`spinCount` iterations, 100000 by default). Mutates
     * `options`: `spinCount` is made an integer >= 0.
     */
    protect(password?: string, options?: Partial<WorksheetProtection>): Promise<void> {
        // TODO: make this function truly async
        // perhaps marshal to worker thread or something
        return new Promise((resolve) => {
            const sheetProtection: SheetProtectionModel = {
                sheet: true,
            };
            this.sheetProtection = sheetProtection;
            let spinCount = 100000;
            if (options && "spinCount" in options) {
                // force spinCount to be integer >= 0 (allow user specified spinCount)
                const requested = options.spinCount;
                spinCount = options.spinCount =
                    typeof requested === "number" && Number.isFinite(requested) ? Math.round(Math.max(0, requested)) : 100000;
            }
            if (password) {
                sheetProtection.algorithmName = "SHA-512";
                sheetProtection.saltValue = Encryptor.randomBytes(16).toString("base64");
                sheetProtection.spinCount = spinCount;
                sheetProtection.hashValue = Encryptor.convertPasswordToHash(password, "SHA512", sheetProtection.saltValue, spinCount);
            }
            if (options) {
                this.sheetProtection = Object.assign(sheetProtection, options);
                if (!password && "spinCount" in options) {
                    delete this.sheetProtection.spinCount;
                }
            }
            resolve();
        });
    }

    /** Removes the protection. */
    unprotect(): void {
        this.sheetProtection = null;
    }

    // ================================================================================

    private _write(text: string): void {
        xmlBuffer.reset();
        xmlBuffer.addText(text);
        this.stream.write(xmlBuffer);
    }

    private _writeSheetProperties(xmlBuf: StringBuf, properties: WorksheetWriterProperties, pageSetup: WorksheetWriterPageSetup): void {
        const sheetPropertiesModel = {
            outlineProperties: properties && properties.outlineProperties,
            tabColor: properties && properties.tabColor,
            pageSetup:
                pageSetup && pageSetup.fitToPage
                    ? {
                          fitToPage: pageSetup.fitToPage,
                      }
                    : undefined,
        };

        xmlBuf.addText(xform.sheetProperties.toXml(sheetPropertiesModel));
    }

    private _writeSheetFormatProperties(xmlBuf: StringBuf, properties: WorksheetWriterProperties): void {
        const sheetFormatPropertiesModel: Partial<WorksheetProperties> = {
            defaultRowHeight: properties.defaultRowHeight,
            dyDescent: properties.dyDescent,
            outlineLevelCol: properties.outlineLevelCol,
            outlineLevelRow: properties.outlineLevelRow,
        };
        if (properties.defaultColWidth) {
            sheetFormatPropertiesModel.defaultColWidth = properties.defaultColWidth;
        }

        xmlBuf.addText(xform.sheetFormatProperties.toXml(sheetFormatPropertiesModel));
    }

    private _writeOpenWorksheet(): void {
        xmlBuffer.reset();

        xmlBuffer.addText('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        xmlBuffer.addText(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
                ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
                ' mc:Ignorable="x14ac"' +
                ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        );

        this._writeSheetProperties(xmlBuffer, this.properties, this.pageSetup);

        xmlBuffer.addText(xform.sheetViews.toXml(this.views));

        this._writeSheetFormatProperties(xmlBuffer, this.properties);

        this.stream.write(xmlBuffer);
    }

    private _writeColumns(): void {
        const cols = Column.toModel(this.columns);
        if (cols) {
            const options: ListItemOptions & { styles: StyleRegistry } = { styles: this._workbook.styles };
            xform.columns.prepare(cols, options);
            this.stream.write(xform.columns.toXml(cols));
        }
    }

    private _writeOpenSheetData(): void {
        this._write("<sheetData>");
    }

    private _writeRow(row: Row): void {
        if (!this.startedData) {
            this._writeColumns();
            this._writeOpenSheetData();
            this.startedData = true;
        }

        // a row of empty cells is still written for the notes of its cells
        if (row.hasValues || row.height || someEntry(row._cells, (cell) => cell._comment)) {
            // not null: the row has a height, or cells (a cell always has a model)
            const model = row.model!;
            const options: CellPrepareOptions = {
                styles: this._workbook.styles,
                sharedStrings: this.useSharedStrings ? this._workbook.sharedStrings : undefined,
                hyperlinks: this._sheetRelsWriter.hyperlinksProxy,
                merges: this._merges,
                formulae: this._formulae,
                siFormulae: this._siFormulae,
                comments: [],
            };
            xform.row.prepare(model, options);
            // the next shared formula index: the prepare of the cells increments it
            this._siFormulae = options.siFormulae;
            this.stream.write(xform.row.toXml(model));

            if (options.comments.length) {
                this.hasComments = true;
                this._sheetCommentsWriter.addComments(options.comments);
            }
        }
    }

    private _writeCloseSheetData(): void {
        this._write("</sheetData>");
    }

    private _writeMergeCells(): void {
        if (this._merges.length) {
            xmlBuffer.reset();
            xmlBuffer.addText(`<mergeCells count="${this._merges.length}">`);
            for (const merge of this._merges) {
                xmlBuffer.addText(`<mergeCell ref="${merge.toString()}"/>`);
            }
            xmlBuffer.addText("</mergeCells>");

            this.stream.write(xmlBuffer);
        }
    }

    private _writeHyperlinks(): void {
        this.stream.write(xform.hyperlinks.toXml(this._sheetRelsWriter._hyperlinks));
    }

    private _writeConditionalFormatting(): void {
        const options = {
            styles: this._workbook.styles,
        };
        xform.conditionalFormattings.prepare(this.conditionalFormatting, options);
        // assigns the x14 ids that the cfRules reference, before they are rendered
        xform.extLst.prepare({ conditionalFormattings: this.conditionalFormatting });
        this.stream.write(xform.conditionalFormattings.toXml(this.conditionalFormatting));
    }

    private _writeRowBreaks(): void {
        this.stream.write(xform.rowBreaks.toXml(this.rowBreaks));
    }

    private _writeDataValidations(): void {
        this.stream.write(xform.dataValidations.toXml(this.dataValidations.model));
    }

    private _writeSheetProtection(): void {
        this.stream.write(xform.sheetProtection.toXml(this.sheetProtection));
    }

    private _writePrintOptions(): void {
        this.stream.write(xform.printOptions.toXml(this.pageSetup));
    }

    private _writePageMargins(): void {
        this.stream.write(xform.pageMargins.toXml(this.pageSetup.margins));
    }

    private _writePageSetup(): void {
        this.stream.write(xform.pageSeteup.toXml(this.pageSetup));
    }

    private _writeHeaderFooter(): void {
        this.stream.write(xform.headerFooter.toXml(this.headerFooter));
    }

    private _writeExtLst(): void {
        // x14 conditional formattings (e.g. dataBar), prepared in _writeConditionalFormatting
        this.stream.write(xform.extLst.toXml({ conditionalFormattings: this.conditionalFormatting }));
    }

    private _writeAutoFilter(): void {
        this.stream.write(xform.autoFilter.toXml(this.autoFilter));
    }

    private _writeBackground(): void {
        if (this._background) {
            if (this._background.imageId !== undefined) {
                const image = this._workbook.getImage(this._background.imageId);
                if (!image) {
                    throw new Error(`Unknown background image id: ${this._background.imageId}`);
                }
                const pictureId = this._sheetRelsWriter.addMedia({
                    Target: `../media/${image.name}`,
                    Type: RelType.Image,
                });

                this._background = {
                    ...this._background,
                    rId: pictureId,
                };
            }
            this.stream.write(xform.picture.toXml({ rId: this._background.rId }));
        }
    }

    private _writeLegacyData(): void {
        if (this.hasComments) {
            xmlBuffer.reset();
            xmlBuffer.addText(`<legacyDrawing r:id="${this._sheetCommentsWriter.vmlRelId}"/>`);
            this.stream.write(xmlBuffer);
        }
    }

    /** Writes nothing: Excel can't handle dimensions at the bottom of the file, and they are known only at the commit. */
    _writeDimensions(): void {
        // this._write('<dimension ref="' + this._dimensions + '"/>');
    }

    private _writeCloseWorksheet(): void {
        this._write("</worksheet>");
    }
}

export default WorksheetWriter;
