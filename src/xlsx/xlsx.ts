import * as fs from "fs";
import { PassThrough, type Stream } from "stream";
import { ZipWriter, type ZipOutput, type ZipWriterOptions } from "../utils/zip-stream";
import { readZip, ZIP_INPUT_ERROR, type ZipInput, type ZipReadOptions } from "../utils/zip-reader";
import RelType from "./rel-type";

import { createDictionary, exists } from "../utils/utils";
import XmlStream from "../utils/xml-stream";
import { bufferToString } from "../utils/browser-buffer-decode";
import { checkMaxMergedCells } from "../utils/merge-limits";
import Pacer, { nextTurn, runSteps } from "../utils/pacer";
import type { SaxInput } from "../utils/parse-sax";

import StylesXform from "./xform/style/styles-xform";

import CoreXform from "./xform/core/core-xform";
import SharedStringsXform from "./xform/strings/shared-strings-xform";
import RelationshipsXform from "./xform/core/relationships-xform";
import ContentTypesXform from "./xform/core/content-types-xform";
import AppXform from "./xform/core/app-xform";
import WorkbookXform from "./xform/book/workbook-xform";
import WorksheetXform from "./xform/sheet/worksheet-xform";
import DrawingXform from "./xform/drawing/drawing-xform";
import TableXform from "./xform/table/table-xform";
import PivotCacheRecordsXform from "./xform/pivot-table/pivot-cache-records-xform";
import PivotCacheDefinitionXform from "./xform/pivot-table/pivot-cache-definition-xform";
import PivotTableXform from "./xform/pivot-table/pivot-table-xform";
import CommentsXform from "./xform/comment/comments-xform";
import VmlNotesXform from "./xform/comment/vml-notes-xform";

import theme1Xml from "./xml/theme1";
import type { XlsxReadOptions } from "../../index";
import type {
    DrawingModel,
    DrawingReconcileOptions,
    MediaModel,
    RelationshipModel,
    TableModel,
    WorkbookModel,
    WorkbookPartModel,
    WorkbookPrepareOptions,
    WorkbookReadModel,
    WorkbookReconcileOptions,
    WorkbookWriteModel,
    WorksheetLoadBudget,
} from "../types";

// ---------------------------------------------------------------------------------------------------------------------

/** The styles manager: StylesXform (a parsed or a new style sheet) or its mock (no styles). */
export type XlsxStyles = StylesXform;

/** What xlsx.ts needs from the workbook (src/doc/workbook.ts). */
export interface XlsxWorkbook {
    /** the cap on the merged cells of the workbook, kept by a loaded workbook (a ModelContainer has none: the default cap applies) */
    maxMergedCells?: number;
    get model(): WorkbookModel;
    set model(value: WorkbookModel | XlsxReadModel);
}

// ---------------------------------------------------------------------------------------------------------------------

/** The read model of {@link XLSX.load}: the parts read, the managers of the styles and shared strings, the load budget. */
export type XlsxReadModel = Omit<WorkbookReadModel, "sharedStrings"> & {
    styles?: XlsxStyles;
    sharedStrings?: SharedStringsXform;
    loadBudget?: WorksheetLoadBudget;
};

/** The write model of {@link XLSX.write}: the document model prepared, with the managers of the styles and shared strings. */
export type XlsxWriteModel = WorkbookWriteModel & {
    styles: XlsxStyles;
    sharedStrings: SharedStringsXform;
};

/** The options of {@link XLSX.load}: index.d.ts `XlsxReadOptions`, and `base64` (the data is a base64 string). */
export type XlsxLoadOptions = Partial<XlsxReadOptions> & ZipReadOptions & { base64?: boolean };

/** The options of {@link XLSX.write}. */
export interface XlsxWriteOptions {
    /** use a shared strings table; default true */
    useSharedStrings?: boolean;
    /** write the styles of the cells; default true */
    useStyles?: boolean;
    /** the zip options (compression level...) */
    zip?: ZipWriterOptions;
}

/** The data of a workbook to load: the bytes, a Blob (e.g. a File), or a base64 string (`base64` option). */
export type XlsxLoadData = ZipInput | Blob | string;

// the parts that xlsx.ts deletes from the read model once reconciled
type ReadState = Partial<
    Pick<
        XlsxReadModel,
        | "worksheetHash"
        | "worksheetRels"
        | "globalRels"
        | "sharedStrings"
        | "workbookRels"
        | "styles"
        | "mediaIndex"
        | "drawings"
        | "drawingRels"
        | "vmlDrawings"
        | "loadBudget"
    > & { sheetDefs: unknown }
>;

function fsReadFileAsync(filename: string): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        fs.readFile(filename, (error, data) => {
            if (error) {
                reject(error);
            } else {
                resolve(data);
            }
        });
    });
}

// a relationships part: RelationshipsXform throws on another root element; a part without any element (empty) gives
// undefined
async function parseRelationships(stream: SaxInput): Promise<RelationshipModel[] | undefined> {
    const xform = new RelationshipsXform();
    return (await xform.parseStream(stream)) ?? undefined;
}

// the first group of a part name pattern (they all have one, not optional), or undefined when the name does not match
function firstGroup(pattern: RegExp, name: string): string | undefined {
    const match = pattern.exec(name);
    return match ? match[1] : undefined;
}

// a stream that has an async iterator (a Node.js Readable, an async generator...)
function isAsyncIterable(stream: object): stream is AsyncIterable<unknown> {
    return Symbol.asyncIterator in stream && typeof stream[Symbol.asyncIterator] === "function";
}

// image file extensions: alphanumeric only (no path separators, no XML)
const IMAGE_EXTENSION = /^[a-z0-9]{1,5}$/i;

// the text of an XML part, in chunks for the SAX parser, pausing between them (utils/pacer)
async function* xmlChunks(content: string, pacer: Pacer): AsyncGenerator<string, void, undefined> {
    const chunkSize = 16 * 1024;
    for (let i = 0; i < content.length; i += chunkSize) {
        yield content.substring(i, i + chunkSize);
        const pause = pacer.pause();
        if (pause) {
            await pause;
        }
    }
}

// the names of the parts that load() reads by pattern (the regexes are linear: no nested quantifier)
const WORKSHEET_PART = /xl\/worksheets\/sheet(\d+)[.]xml/;
const WORKSHEET_RELS_PART = /xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/;
const THEME_PART = /xl\/theme\/([a-zA-Z0-9]+)[.]xml/;
const MEDIA_PART = /xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/;
const DRAWING_PART = /xl\/drawings\/([a-zA-Z0-9]+)[.]xml/;
// xl/comments1.xml (Excel), xl/comments/comment1.xml (openpyxl)
const COMMENTS_PART = /xl\/(comments\d+|comments\/comment\d+)[.]xml$/;
const TABLE_PART = /xl\/tables\/(table\d+)[.]xml/;
const DRAWING_RELS_PART = /xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/;
// vmlDrawing1.vml (Excel), commentsDrawing1.vml (openpyxl)
const VML_DRAWING_PART = /xl\/drawings\/(vmlDrawing\d+|commentsDrawing\d+)[.]vml$/;

// The parts that load() reads (keep in sync with its switch): the others (pivot tables, charts, custom XML, printer
// settings...) are not inflated.
const READ_PARTS = new Set([
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "docProps/app.xml",
    "docProps/core.xml",
]);
const READ_PART_PATTERNS = [
    WORKSHEET_PART,
    WORKSHEET_RELS_PART,
    THEME_PART,
    MEDIA_PART,
    DRAWING_PART,
    COMMENTS_PART,
    TABLE_PART,
    DRAWING_RELS_PART,
    VML_DRAWING_PART,
];

function isReadPart(name: string): boolean {
    return READ_PARTS.has(name) || READ_PART_PATTERNS.some((pattern) => pattern.test(name));
}

// the browser bundle defines process.browser (scripts/browser-shims)
const isBrowser = "browser" in process && Boolean(process.browser);

/**
 * Reads and writes the xlsx file of a workbook (`workbook.xlsx`): the zip parts and their XML, through the xforms.
 * Works in browsers too, except {@link readFile} / {@link writeFile}.
 */
class XLSX<TWorkbook extends XlsxWorkbook = XlsxWorkbook> {
    /** The relationship types (rel-type.ts). */
    static readonly RelType = RelType;

    /** The workbook read into and written from (a Workbook or a ModelContainer): what the read methods resolve with. */
    readonly workbook: TWorkbook;

    constructor(workbook: TWorkbook) {
        this.workbook = workbook;
    }

    // ===============================================================================
    // Workbook
    // =========================================================================
    // Read

    /**
     * Reads a file into the workbook (Node.js and Bun only).
     *
     * @throws (rejects) when the file does not exist, or as {@link load}
     */
    async readFile(filename: string, options?: XlsxLoadOptions): Promise<TWorkbook> {
        if (!(await exists(filename))) {
            throw new Error(`File not found: ${filename}`);
        }
        const stream = fs.createReadStream(filename);
        try {
            const workbook = await this.read(stream, options);
            stream.close();
            return workbook;
        } catch (error) {
            stream.close();
            throw error;
        }
    }

    /** Parses a relationships part. */
    parseRels(stream: SaxInput): Promise<RelationshipModel[] | null | undefined> {
        const xform = new RelationshipsXform();
        return xform.parseStream(stream);
    }

    /**
     * Parses the workbook part.
     * @throws Error when the part has no `<workbook>` element
     */
    async parseWorkbook(stream: SaxInput): Promise<WorkbookPartModel> {
        const xform = new WorkbookXform();
        const workbook = await xform.parseStream(stream);
        if (!workbook) {
            // it used to be a TypeError on the missing model
            throw new Error("Invalid xl/workbook.xml: no <workbook> element");
        }
        return workbook;
    }

    /** Parses the shared strings part. */
    parseSharedStrings(stream: SaxInput): ReturnType<SharedStringsXform["parseStream"]> {
        const xform = new SharedStringsXform();
        return xform.parseStream(stream);
    }

    /** Reconciles the parts read (see {@link load}), at once. */
    reconcile(model: XlsxReadModel, options?: XlsxLoadOptions): void {
        runSteps(this._reconcileSteps(model, options));
    }

    // reconcile() worksheet row by worksheet row: the caller may pause between rows
    *_reconcileSteps(model: XlsxReadModel, options?: XlsxLoadOptions): Generator<unknown, void, undefined> {
        const workbookXform = new WorkbookXform();
        const worksheetXform = new WorksheetXform(options);
        const drawingXform = new DrawingXform();
        const tableXform = new TableXform();

        workbookXform.reconcile(model);

        // reconcile drawings with their rels
        const drawingOptions: DrawingReconcileOptions = {
            media: model.media,
            mediaIndex: model.mediaIndex,
        };
        for (const [name, drawing] of Object.entries(model.drawings)) {
            // part names and relationship ids come from the file: ignore inherited members
            const drawingRel = Object.hasOwn(model.drawingRels, name) ? model.drawingRels[name] : undefined;
            if (drawingRel) {
                const rels = (drawingOptions.rels = drawingRel.reduce<Record<string, RelationshipModel>>((o, rel) => {
                    o[rel.Id] = rel;
                    return o;
                }, createDictionary()));
                for (const anchor of drawing.anchors || []) {
                    const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
                    // String(): a picture without r:id looks up "undefined", as it always did
                    const rel = hyperlinks && rels[String(hyperlinks.rId)];
                    if (hyperlinks && rel) {
                        hyperlinks.hyperlink = rel.Target;
                        delete hyperlinks.rId;
                    }
                }
                drawingXform.reconcile(drawing, drawingOptions);
            }
        }

        // reconcile tables with the default styles
        // a file without styles part has no dxfs: the columns with a dxfId get no style (instead of a TypeError)
        const tableOptions = {
            styles: model.styles ?? new StylesXform(),
        };
        for (const table of Object.values(model.tables)) {
            tableXform.reconcile(table, tableOptions);
        }

        const sheetOptions: WorkbookReconcileOptions = {
            styles: model.styles,
            sharedStrings: model.sharedStrings,
            media: model.media,
            mediaIndex: model.mediaIndex,
            date1904: model.properties && model.properties.date1904,
            drawings: model.drawings,
            comments: model.comments,
            tables: model.tables,
            vmlDrawings: model.vmlDrawings,
            // the parts already claimed by a worksheet (see WorksheetXform.reconcile)
            claimedParts: new Set(),
        };
        for (const worksheet of model.worksheets) {
            // sheetNo: the number of the part (a string)
            worksheet.relationships = model.worksheetRels[String(worksheet.sheetNo)];
            yield* worksheetXform.reconcileSteps(worksheet, sheetOptions);
        }

        // delete unnecessary parts
        const readState: ReadState = model;
        delete readState.worksheetHash;
        delete readState.worksheetRels;
        delete readState.globalRels;
        delete readState.sharedStrings;
        delete readState.workbookRels;
        delete readState.sheetDefs;
        delete readState.styles;
        delete readState.mediaIndex;
        delete readState.drawings;
        delete readState.drawingRels;
        delete readState.vmlDrawings;
        delete readState.loadBudget;
    }

    async _processWorksheetEntry(
        stream: SaxInput,
        model: XlsxReadModel,
        sheetNo: string,
        options: XlsxLoadOptions | undefined,
        path: string
    ): Promise<void> {
        const xform = new WorksheetXform({ ...options, loadBudget: model.loadBudget });
        await xform.parseStream(stream);
        const worksheet = xform.model;
        if (!worksheet) {
            // it used to be a TypeError on the missing model
            throw new Error(`Invalid worksheet part ${path}: no <worksheet> element`);
        }
        worksheet.sheetNo = sheetNo;
        model.worksheetHash[path] = worksheet;
        model.worksheets.push(worksheet);
    }

    async _processCommentEntry(stream: SaxInput, model: XlsxReadModel, entryName: string): Promise<void> {
        const xform = new CommentsXform();
        const comments = await xform.parseStream(stream);
        // a part without <comments> parses to nothing: no entry (a lookup gives undefined, as the stored undefined did)
        if (comments) model.comments[entryName] = comments;
    }

    async _processTableEntry(stream: SaxInput, model: XlsxReadModel, entryName: string): Promise<void> {
        const xform = new TableXform();
        const table = await xform.parseStream(stream);
        // a part without <table> parses to nothing: no entry (a lookup gives undefined, as the stored undefined did)
        if (table) model.tables[entryName] = table;
    }

    async _processWorksheetRelsEntry(stream: SaxInput, model: XlsxReadModel, sheetNo: string): Promise<void> {
        const relationships = await parseRelationships(stream);
        // an empty part: no entry (a lookup gives undefined, as the stored undefined did)
        if (relationships) model.worksheetRels[sheetNo] = relationships;
    }

    /** Adds a media part (its bytes, a fresh Buffer) to the media of the model. */
    _processMediaEntry(bytes: Buffer, model: XlsxReadModel, filename: string): void {
        const lastDot = filename.lastIndexOf(".");
        // if we can't determine extension, ignore it
        if (lastDot >= 1) {
            const extension = filename.substring(lastDot + 1);
            const name = filename.substring(0, lastDot);
            model.mediaIndex[filename] = model.media.length;
            model.mediaIndex[name] = model.media.length;
            const medium: MediaModel = {
                type: "image",
                name,
                extension,
                // null for an empty part, as the stream of its bytes gave
                buffer: bytes.length ? bytes : null,
            };
            model.media.push(medium);
        }
    }

    async _processDrawingEntry(entry: SaxInput, model: XlsxReadModel, name: string): Promise<void> {
        const xform = new DrawingXform();
        const drawing = await xform.parseStream(entry);
        // null when the part is not a spreadsheet drawing (e.g. a chart's <c:userShapes>)
        if (drawing) {
            model.drawings[name] = drawing;
        }
    }

    async _processDrawingRelsEntry(entry: SaxInput, model: XlsxReadModel, name: string): Promise<void> {
        const relationships = await parseRelationships(entry);
        // an empty part: no entry (a lookup gives undefined, as the stored undefined did)
        if (relationships) model.drawingRels[name] = relationships;
    }

    async _processVmlDrawingEntry(entry: SaxInput, model: XlsxReadModel, entryName: string): Promise<void> {
        const xform = new VmlNotesXform();
        const notes = await xform.parseStream(entry);
        // an empty part parses to nothing: no entry (a lookup gives undefined, as the stored undefined did)
        if (notes) model.vmlDrawings[entryName] = notes;
    }

    /** Keeps the XML of a theme part (its bytes), as is. */
    _processThemeEntry(bytes: Buffer, model: XlsxReadModel, name: string): void {
        model.themes[name] = bytes.toString();
    }

    /**
     * @deprecated since version 4.0. You should use `#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
     * @throws Error always
     */
    createInputStream(): never {
        throw new Error(
            "`XLSX#createInputStream` is deprecated. You should use `XLSX#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md"
        );
    }

    /**
     * Reads a stream of the bytes of a file into the workbook, as {@link load}.
     *
     * @throws (rejects) on an error of the stream, or as {@link load}
     */
    async read(stream: AsyncIterable<Uint8Array> | NodeJS.ReadableStream | Stream, options?: XlsxLoadOptions): Promise<TWorkbook> {
        // an old-style stream (without async iterator) is read through a PassThrough
        const input = isAsyncIterable(stream) ? stream : stream.pipe(new PassThrough());
        const chunks: Uint8Array[] = [];
        for await (const chunk of input) {
            if (!(chunk instanceof Uint8Array)) {
                // e.g. a stream with an encoding, which gives strings
                throw new TypeError("Can't read the zip file: the stream must give bytes (Buffers or Uint8Arrays)");
            }
            chunks.push(chunk);
        }
        return this.load(Buffer.concat(chunks), options);
    }

    /**
     * Loads a file into the workbook: its worksheets, styles, shared strings, images, notes, tables, defined names and
     * properties (pivot tables, charts... are not read).
     *
     * @param data - the bytes of the file, a Blob, or a base64 string (`base64` option)
     * @returns the workbook (its model replaced), once every part is read
     * @throws (rejects) on an invalid zip or XML, a decompression limit or a limit of the worksheets exceeded
     * (rows, columns, merged cells...: see the invariants of CLAUDE.md)
     */
    async load(data: XlsxLoadData, options?: XlsxLoadOptions): Promise<TWorkbook> {
        let buffer: XlsxLoadData;
        if (options && options.base64) {
            // eslint-disable-next-line @typescript-eslint/no-base-to-string -- a base64 string (or its Buffer), as it always was
            buffer = Buffer.from(data.toString(), "base64");
        } else {
            buffer = data;
        }

        // keyed by the names of the parts (zip entry names): no prototype (CLAUDE.md invariant)
        const model: XlsxReadModel = {
            worksheets: [],
            worksheetHash: createDictionary(),
            worksheetRels: createDictionary(),
            themes: createDictionary(),
            media: [],
            mediaIndex: createDictionary(),
            drawings: createDictionary(),
            drawingRels: createDictionary(),
            comments: createDictionary(),
            tables: createDictionary(),
            vmlDrawings: createDictionary(),
            // replaced by those of xl/workbook.xml
            views: [],
        };
        // the cap on merged cells applies to the whole workbook, and is kept by the loaded workbook
        const maxMergedCells =
            options && options.maxMergedCells !== undefined ? checkMaxMergedCells(options.maxMergedCells) : this.workbook.maxMergedCells;
        model.loadBudget = WorksheetXform.createLoadBudget(maxMergedCells);

        const pacer = new Pacer();

        let zipBytes: ZipInput;
        if (typeof Blob !== "undefined" && buffer instanceof Blob) {
            // like JSZip did, accept a Blob (e.g. a File from an <input type="file">)
            zipBytes = new Uint8Array(await buffer.arrayBuffer());
        } else if (typeof buffer !== "string" && (ArrayBuffer.isView(buffer) || buffer instanceof ArrayBuffer || Array.isArray(buffer))) {
            zipBytes = buffer;
        } else {
            // a string without the base64 option, or anything else from a JavaScript caller (as readZip rejects it)
            throw new TypeError(ZIP_INPUT_ERROR);
        }
        // the limits (maxEntrySize...) are checked while inflating
        for (const entry of readZip(zipBytes, options)) {
            const pause = pacer.pause();
            if (pause) {
                await pause;
            }
            if (!entry.dir) {
                let entryName = entry.name;
                if (entryName[0] === "/") {
                    entryName = entryName.substring(1);
                }
                if (!isReadPart(entryName)) {
                    continue;
                }
                const entryData = await entry.data();
                const bytes = Buffer.from(entryData.buffer, entryData.byteOffset, entryData.byteLength);
                // media and themes are read as bytes (the stream of their bytes is parsed as XML only when their name also
                // matches an XML part pattern tested before theirs); the XML parts are parsed as text, chunk by chunk
                let stream: SaxInput;
                if (
                    entryName.includes("xl/media/") ||
                    // themes are not parsed as stream
                    THEME_PART.test(entryName)
                ) {
                    const byteStream = new PassThrough();
                    byteStream.write(bytes);
                    byteStream.end();
                    stream = byteStream;
                } else {
                    // https://www.npmjs.com/package/process: running in browser, use TextDecoder if possible
                    const content = isBrowser ? bufferToString(entryData) : bytes.toString("utf8");
                    stream = xmlChunks(content, pacer);
                }
                switch (entryName) {
                    case "_rels/.rels":
                        model.globalRels = (await this.parseRels(stream)) ?? undefined;
                        break;

                    case "xl/workbook.xml": {
                        const workbook = await this.parseWorkbook(stream);
                        model.sheets = workbook.sheets;
                        model.definedNames = workbook.definedNames;
                        model.views = workbook.views;
                        model.properties = workbook.properties;
                        model.calcProperties = workbook.calcProperties;
                        break;
                    }

                    case "xl/_rels/workbook.xml.rels":
                        model.workbookRels = (await this.parseRels(stream)) ?? undefined;
                        break;

                    case "xl/sharedStrings.xml":
                        model.sharedStrings = new SharedStringsXform();
                        await model.sharedStrings.parseStream(stream);
                        break;

                    case "xl/styles.xml":
                        model.styles = new StylesXform();
                        await model.styles.parseStream(stream);
                        break;

                    case "docProps/app.xml": {
                        const appXform = new AppXform();
                        const appProperties = await appXform.parseStream(stream);
                        // undefined when the part has no <Properties> root (e.g. a prefixed <ep:Properties>)
                        if (appProperties) {
                            model.company = appProperties.company;
                            model.manager = appProperties.manager;
                        }
                        break;
                    }

                    case "docProps/core.xml": {
                        const coreXform = new CoreXform();
                        const coreProperties = await coreXform.parseStream(stream);
                        Object.assign(model, coreProperties);
                        break;
                    }

                    default:
                        await this._processPatternEntry(entryName, stream, bytes, model, options);
                        break;
                }
            }
        }

        await pacer.run(this._reconcileSteps(model, options));

        // apply model
        this.workbook.maxMergedCells = maxMergedCells;
        this.workbook.model = model;
        return this.workbook;
    }

    // the parts read by the name patterns (READ_PART_PATTERNS)
    private async _processPatternEntry(
        entryName: string,
        stream: SaxInput,
        bytes: Buffer,
        model: XlsxReadModel,
        options: XlsxLoadOptions | undefined
    ): Promise<void> {
        let name = firstGroup(WORKSHEET_PART, entryName);
        if (name !== undefined) {
            await this._processWorksheetEntry(stream, model, name, options, entryName);
            return;
        }
        name = firstGroup(WORKSHEET_RELS_PART, entryName);
        if (name !== undefined) {
            await this._processWorksheetRelsEntry(stream, model, name);
            return;
        }
        name = firstGroup(THEME_PART, entryName);
        if (name !== undefined) {
            this._processThemeEntry(bytes, model, name);
            return;
        }
        name = firstGroup(MEDIA_PART, entryName);
        if (name !== undefined) {
            this._processMediaEntry(bytes, model, name);
            return;
        }
        name = firstGroup(DRAWING_PART, entryName);
        if (name !== undefined) {
            await this._processDrawingEntry(stream, model, name);
            return;
        }
        if (COMMENTS_PART.test(entryName)) {
            await this._processCommentEntry(stream, model, entryName);
            return;
        }
        if (TABLE_PART.test(entryName)) {
            await this._processTableEntry(stream, model, entryName);
            return;
        }
        name = firstGroup(DRAWING_RELS_PART, entryName);
        if (name !== undefined) {
            await this._processDrawingRelsEntry(stream, model, name);
            return;
        }
        if (VML_DRAWING_PART.test(entryName)) {
            await this._processVmlDrawingEntry(stream, model, entryName);
        }
    }

    // =========================================================================
    // Write

    /**
     * Adds the media parts (an image file is read now).
     *
     * @throws (rejects) for a medium without data, or a file that cannot be read
     */
    async addMedia(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        await Promise.all(
            model.media.map(async (medium) => {
                if (medium.type === "image") {
                    const filename = `xl/media/${medium.name}.${medium.extension}`;
                    if (medium.filename) {
                        const data = await fsReadFileAsync(medium.filename);
                        return zip.append(data, { name: filename });
                    }
                    if (medium.buffer) {
                        return zip.append(medium.buffer, { name: filename });
                    }
                    if (medium.base64) {
                        const dataimg64 = medium.base64;
                        const content = dataimg64.substring(dataimg64.indexOf(",") + 1);
                        return zip.append(content, { name: filename, base64: true });
                    }
                }
                throw new Error("Unsupported media");
            })
        );
    }

    /** Adds the drawing parts of the worksheets and their relationships. */
    addDrawings(zip: ZipWriter, model: XlsxWriteModel): void {
        const drawingXform = new DrawingXform();
        const relsXform = new RelationshipsXform();

        for (const worksheet of model.worksheets) {
            const { drawing } = worksheet;
            if (drawing) {
                drawingXform.prepare(drawing);
                let xml = drawingXform.toXml(drawing);
                zip.append(xml, { name: `xl/drawings/${drawing.name}.xml` });

                // worksheet-xform prepare gives the drawing its relationships
                if (!drawing.rels) {
                    throw new Error(`addDrawings: the drawing ${drawing.name} has no relationships (prepareModel was not run)`);
                }
                xml = relsXform.toXml(drawing.rels);
                zip.append(xml, { name: `xl/drawings/_rels/${drawing.name}.xml.rels` });
            }
        }
    }

    /** Adds the table parts of the worksheets. */
    addTables(zip: ZipWriter, model: XlsxWriteModel): void {
        const tableXform = new TableXform();

        for (const worksheet of model.worksheets) {
            for (const table of worksheet.tables) {
                tableXform.prepare(table, {});
                const tableXml = tableXform.toXml(table);
                // prepareModel names the table parts
                zip.append(tableXml, { name: `xl/tables/${table.target}` });
            }
        }
    }

    /** Adds the parts of the pivot table of the workbook (one at most). */
    addPivotTables(zip: ZipWriter, model: XlsxWriteModel): void {
        const [pivotTable] = model.pivotTables;
        if (!pivotTable) return;

        const pivotCacheRecordsXform = new PivotCacheRecordsXform();
        const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
        const pivotTableXform = new PivotTableXform();
        const relsXform = new RelationshipsXform();

        // pivot cache records
        // --------------------------------------------------
        // copy of the source data.
        //
        // Note: cells in the columns of the source data which are part
        // of the "rows" or "columns" of the pivot table configuration are
        // replaced by references to their __cache field__ identifiers.
        // See "pivot cache definition" below.

        let xml = pivotCacheRecordsXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotCache/pivotCacheRecords1.xml" });

        // pivot cache definition
        // --------------------------------------------------
        // cache source (source data):
        //    ref="A1:E7" on sheet="Sheet1"
        // cache fields:
        //    - 0: "A" (a1, a2, a3)
        //    - 1: "B" (b1, b2)
        //    - ...

        xml = pivotCacheDefinitionXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotCache/pivotCacheDefinition1.xml" });

        xml = relsXform.toXml([
            {
                Id: "rId1",
                Type: XLSX.RelType.PivotCacheRecords,
                Target: "pivotCacheRecords1.xml",
            },
        ]);
        zip.append(xml, { name: "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels" });

        // pivot tables (on destination worksheet)
        // --------------------------------------------------
        // location: ref="A3:E15"
        // pivotFields
        // rowFields and rowItems
        // colFields and colItems
        // dataFields
        // pivotTableStyleInfo

        xml = pivotTableXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotTables/pivotTable1.xml" });

        xml = relsXform.toXml([
            {
                Id: "rId1",
                Type: XLSX.RelType.PivotCacheDefinition,
                Target: "../pivotCache/pivotCacheDefinition1.xml",
            },
        ]);
        zip.append(xml, { name: "xl/pivotTables/_rels/pivotTable1.xml.rels" });
    }

    // The add* methods below are async (a failure rejects the promise they return), most without awaiting anything
    /* eslint-disable @typescript-eslint/require-await */

    /** Adds `[Content_Types].xml`. */
    async addContentTypes(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        const xform = new ContentTypesXform();
        const xml = xform.toXml(model);
        zip.append(xml, { name: "[Content_Types].xml" });
    }

    /** Adds the extended properties (`docProps/app.xml`). */
    async addApp(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        const xform = new AppXform();
        const xml = xform.toXml(model);
        zip.append(xml, { name: "docProps/app.xml" });
    }

    /** Adds the core properties (`docProps/core.xml`). */
    async addCore(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        const coreXform = new CoreXform();
        zip.append(coreXform.toXml(model), { name: "docProps/core.xml" });
    }

    /** Adds the theme parts: theme1 (the default one, unless the workbook has its own) and those of the workbook. */
    async addThemes(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        // workbook rels and content types always reference theme1.xml, so it must be written
        // even when the workbook was read from a file without a theme
        const themes: Record<string, string> = { theme1: theme1Xml, ...model.themes };
        for (const [name, xml] of Object.entries(themes)) {
            const path = `xl/theme/${name}.xml`;
            zip.append(xml, { name: path });
        }
    }

    /** Adds the package relationships (`_rels/.rels`). */
    async addOfficeRels(zip: ZipWriter): Promise<void> {
        const xform = new RelationshipsXform();
        const xml = xform.toXml([
            { Id: "rId1", Type: XLSX.RelType.OfficeDocument, Target: "xl/workbook.xml" },
            { Id: "rId2", Type: XLSX.RelType.CoreProperties, Target: "docProps/core.xml" },
            { Id: "rId3", Type: XLSX.RelType.ExtenderProperties, Target: "docProps/app.xml" },
        ]);
        zip.append(xml, { name: "_rels/.rels" });
    }

    /** Adds the workbook relationships; sets the `rId` of each worksheet and of the pivot table. */
    async addWorkbookRels(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        let count = 1;
        const relationships: RelationshipModel[] = [
            { Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: "styles.xml" },
            { Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: "theme/theme1.xml" },
        ];
        if (model.sharedStrings.count) {
            relationships.push({
                Id: `rId${count++}`,
                Type: XLSX.RelType.SharedStrings,
                Target: "sharedStrings.xml",
            });
        }
        const [pivotTable] = model.pivotTables || [];
        if (pivotTable) {
            const rId = (pivotTable.rId = `rId${count++}`);
            relationships.push({
                Id: rId,
                Type: XLSX.RelType.PivotCacheDefinition,
                Target: "pivotCache/pivotCacheDefinition1.xml",
            });
        }
        for (const worksheet of model.worksheets) {
            const rId = (worksheet.rId = `rId${count++}`);
            relationships.push({
                Id: rId,
                Type: XLSX.RelType.Worksheet,
                Target: `worksheets/sheet${worksheet.id}.xml`,
            });
        }
        const xform = new RelationshipsXform();
        const xml = xform.toXml(relationships);
        zip.append(xml, { name: "xl/_rels/workbook.xml.rels" });
    }

    /** Adds the shared strings part, when the cells used it. */
    async addSharedStrings(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        if (model.sharedStrings && model.sharedStrings.count) {
            zip.append(model.sharedStrings.xml, { name: "xl/sharedStrings.xml" });
        }
    }

    /** Adds the styles part. */
    async addStyles(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        const { xml } = model.styles;
        if (xml) {
            zip.append(xml, { name: "xl/styles.xml" });
        }
    }

    /** Adds the workbook part. */
    async addWorkbook(zip: ZipWriter, model: XlsxWriteModel): Promise<void> {
        const xform = new WorkbookXform();
        zip.append(xform.toXml(model), { name: "xl/workbook.xml" });
    }

    /* eslint-enable @typescript-eslint/require-await */

    /** Adds the worksheet parts, and their relationships, comments and VML drawings. */
    async addWorksheets(zip: ZipWriter, model: XlsxWriteModel, pacer = new Pacer()): Promise<void> {
        // preparation phase
        const worksheetXform = new WorksheetXform();
        const relationshipsXform = new RelationshipsXform();
        const commentsXform = new CommentsXform();
        const vmlNotesXform = new VmlNotesXform();

        // the blocks of VML shape ids claimed by the notes of each worksheet (VmlNotesXform)
        let nextShapeBlock = 1;

        // write sheets
        for (const worksheet of model.worksheets) {
            let xmlStream = new XmlStream();
            await pacer.run(worksheetXform.renderSteps(xmlStream, worksheet));
            // appended whole, so that its sizes are in the local header (the zip layer compresses it at once)
            zip.append(xmlStream.xml, { name: `xl/worksheets/sheet${worksheet.id}.xml` });

            if (worksheet.rels && worksheet.rels.length) {
                xmlStream = new XmlStream();
                relationshipsXform.render(xmlStream, worksheet.rels);
                zip.append(xmlStream.xml, { name: `xl/worksheets/_rels/sheet${worksheet.id}.xml.rels` });
            }

            // worksheet-xform prepare collects the notes
            const { comments } = worksheet;
            if (!comments) {
                throw new Error(`addWorksheets: the worksheet ${worksheet.name} has no notes list (prepareModel was not run)`);
            }
            if (comments.length > 0) {
                xmlStream = new XmlStream();
                commentsXform.render(xmlStream, { comments });
                zip.append(xmlStream.xml, { name: `xl/comments${worksheet.id}.xml` });

                xmlStream = new XmlStream();
                worksheet.shapeBlock = nextShapeBlock;
                nextShapeBlock += VmlNotesXform.blockCount(comments.length);
                vmlNotesXform.render(xmlStream, { comments, shapeBlock: worksheet.shapeBlock });
                zip.append(xmlStream.xml, { name: `xl/drawings/vmlDrawing${worksheet.id}.vml` });
            }
        }
    }

    async _finalize(zip: ZipWriter): Promise<this> {
        // rejects on any failure of the zip writer or of the output stream
        await zip.finalize();
        return this;
    }

    /** Prepares the model of the workbook for writing (see {@link write}), at once. */
    prepareModel(model: WorkbookModel, options: XlsxWriteOptions): void {
        runSteps(this._prepareModelSteps(model, options));
    }

    // prepareModel() worksheet row by worksheet row: the caller may pause between rows. Returns the document model, with
    // the write state added (mutation)
    *_prepareModelSteps(documentModel: WorkbookModel, options: XlsxWriteOptions): Generator<unknown, XlsxWriteModel, undefined> {
        const useStyles = options.useStyles !== undefined ? options.useStyles : true;
        const drawings: DrawingModel[] = [];
        const commentRefs: XlsxWriteModel["commentRefs"] = [];
        const tables: TableModel[] = [];
        const model: XlsxWriteModel = Object.assign(documentModel, {
            // ensure following properties have sane values
            creator: documentModel.creator || "ExcelJS",
            lastModifiedBy: documentModel.lastModifiedBy || "ExcelJS",
            created: documentModel.created || new Date(),
            modified: documentModel.modified || new Date(),

            useSharedStrings: options.useSharedStrings !== undefined ? options.useSharedStrings : true,
            useStyles,

            // Manage the shared strings
            sharedStrings: new SharedStringsXform(),

            // add a style manager to handle cell formats, fonts, etc.
            styles: useStyles ? new StylesXform(true) : new StylesXform.Mock(),

            // set again by WorkbookXform.prepare
            sheets: documentModel.worksheets,
            drawings,
            commentRefs,
            tables,
        });

        // prepare all of the things before the render
        const workbookXform = new WorkbookXform();
        const worksheetXform = new WorksheetXform();

        workbookXform.prepare(model);

        const worksheetOptions: WorkbookPrepareOptions = {
            sharedStrings: model.sharedStrings,
            styles: model.styles,
            date1904: model.properties.date1904,
            drawingsCount: 0,
            media: model.media,
            drawings,
            commentRefs,
        };
        let tableCount = 0;
        for (const worksheet of model.worksheets) {
            // assign unique filenames to tables
            for (const table of worksheet.tables) {
                tableCount++;
                table.target = `table${tableCount}.xml`;
                table.id = tableCount;
                tables.push(table);
            }

            yield* worksheetXform.prepareSteps(worksheet, worksheetOptions);
        }

        // TODO: workbook drawing list
        return model;
    }

    /**
     * Writes the workbook as an xlsx file to a stream (or any object with `write()` and `end()`).
     *
     * @throws (rejects) for an image extension that is not 1 to 5 letters or digits (before anything is written), an
     * unsupported medium, or a failure of the zip writer or of the output stream
     */
    async write(stream: ZipOutput, options?: XlsxWriteOptions): Promise<this> {
        options = options || {};
        const { model: documentModel } = this.workbook;
        // before anything is written: the extension ends up in a zip entry name and in XML
        for (const medium of documentModel.media) {
            if (medium.type === "image" && !IMAGE_EXTENSION.test(medium.extension)) {
                throw new Error(`Invalid image extension: ${JSON.stringify(medium.extension)}`);
            }
        }
        const zip = new ZipWriter(options.zip);
        zip.pipe(stream);

        const pacer = new Pacer();
        const model = await pacer.run(this._prepareModelSteps(documentModel, options));

        // render
        await this.addContentTypes(zip, model);
        await this.addOfficeRels(zip);
        await this.addWorkbookRels(zip, model);
        await this.addWorksheets(zip, model, pacer);
        await this.addSharedStrings(zip, model); // always after worksheets
        this.addDrawings(zip, model);
        this.addTables(zip, model);
        this.addPivotTables(zip, model);
        await Promise.all([this.addThemes(zip, model), this.addStyles(zip, model)]);
        await this.addMedia(zip, model);
        await Promise.all([this.addApp(zip, model), this.addCore(zip, model)]);
        await this.addWorkbook(zip, model);
        // a stream reports a failed write on a later tick: let it do so before the archive is completed. An output that
        // cannot report errors (the collector of writeBuffer) needs no turn, so that writing a small workbook does not wait for
        // a timer (fake timers of test frameworks never run it)
        if (typeof stream.on === "function") {
            await nextTurn();
        }
        return this._finalize(zip);
    }

    /**
     * Writes the workbook to a file (Node.js and Bun only), as {@link write}.
     *
     * @throws (rejects) as {@link write}, or on an error of the file
     */
    writeFile(filename: string, options?: XlsxWriteOptions): Promise<void> {
        const stream = fs.createWriteStream(filename);

        return new Promise((resolve, reject) => {
            stream.on("finish", () => {
                resolve();
            });
            stream.on("error", (error) => {
                reject(error);
            });

            this.write(stream, options)
                .then(() => {
                    stream.end();
                })
                .catch(reject);
        });
    }

    /**
     * Writes the workbook as an xlsx file in memory, as {@link write}.
     *
     * @returns the bytes of the file
     */
    async writeBuffer(options?: XlsxWriteOptions): Promise<Buffer<ArrayBuffer>> {
        const chunks: Buffer[] = [];
        const collector: ZipOutput = {
            write(chunk: Buffer) {
                chunks.push(chunk);
            },
            end() {},
        };
        await this.write(collector, options);
        return Buffer.concat(chunks);
    }
}

export default XLSX;
