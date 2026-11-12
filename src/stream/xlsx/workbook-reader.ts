import * as fs from "fs";
import { EventEmitter } from "events";
import { PassThrough, Readable, pipeline } from "stream";
import type { ReadableStream as WebReadableStream } from "stream/web";
import * as os from "os";
import * as path from "path";
import { parseZipStream, type ZipReadOptions, type ZipStreamEntry } from "../../utils/zip-reader";
import parseSax, { type SaxInput } from "../../utils/parse-sax";

import StylesXform from "../../xlsx/xform/style/styles-xform";
import WorkbookXform from "../../xlsx/xform/book/workbook-xform";
import RelationshipsXform from "../../xlsx/xform/core/relationships-xform";
import SharedStringXform from "../../xlsx/xform/strings/shared-string-xform";
import type { WorkbookProperties, stream as publicStream } from "../../../index";
import type { RelationshipModel, SharedStringValue, SheetDefModel, WorkbookPartModel } from "../../types";

import WorksheetReader from "./worksheet-reader";
import HyperlinkReader from "./hyperlink-reader";

/**
 * What the WorkbookReader reads: a file name, a Node.js stream (or any object with a `pipe` method), a web
 * ReadableStream or an async iterable of byte chunks.
 */
export type WorkbookReaderInput = publicStream.xlsx.WorkbookReaderInput;

/**
 * The options of the WorkbookReader, defaults applied: what to do with each kind of part, and the decompression limits
 * (16 GiB per entry, 32 GiB in all, by default).
 */
export interface WorkbookReaderOptions extends ZipReadOptions {
    /** `emit`: the worksheets are read (WorksheetReader); `ignore`: skipped. */
    worksheets: "emit" | "ignore";
    /** `cache`: kept for the cells; `emit`: emitted one by one (the cells then hold `{ sharedString: index }`); `ignore`. */
    sharedStrings: "cache" | "emit" | "ignore";
    /** The hyperlinks of the worksheets and their relationships: `cache`, `emit` or `ignore`. */
    hyperlinks: "cache" | "emit" | "ignore";
    /** `cache`: the cells get their styles; `ignore`. */
    styles: "cache" | "ignore";
    /** `emit`: an `entry` event for each part read. */
    entries: "emit" | "ignore";
}

/** The payload of an `entry` event (`entries: "emit"`): the kind of part, and the number of a worksheet part. */
export interface WorkbookReaderEntry {
    type: "workbook" | "shared-strings" | "styles" | "worksheet" | "hyperlinks";
    id?: string;
}

/** A shared string emitted by the reader (`sharedStrings: "emit"`): its index in the table and its value. */
export interface EmittedSharedString {
    index: number;
    text: SharedStringValue;
}

/** An event of {@link WorkbookReader.parse}. */
export type WorkbookReaderEvent =
    | { readonly eventType: "shared-strings"; readonly value: EmittedSharedString }
    | { readonly eventType: "worksheet"; readonly value: WorksheetReader }
    | { readonly eventType: "hyperlinks"; readonly value: HyperlinkReader };

/** The events of a {@link WorkbookReader} (with {@link WorkbookReader.read}). */
interface WorkbookReaderEvents {
    worksheet: [worksheet: WorksheetReader];
    "shared-strings": [sharedString: EmittedSharedString];
    hyperlinks: [hyperlinks: HyperlinkReader];
    entry: [entry: WorkbookReaderEntry];
    end: [];
    finished: [];
    error: [error: unknown];
}

/** The styles manager of the reader. */
export type ReaderStyles = StylesXform;

// workbook.xml.rels targets are relative to xl/ or absolute from the package root (openpyxl)
function workbookPartPath(target: unknown): string | undefined {
    if (typeof target !== "string") {
        return undefined;
    }
    return target.startsWith("/") ? target.slice(1) : `xl/${target}`;
}

// size of the chunks read back from the spool file
const SPOOL_CHUNK = 64 * 1024;

// a worksheet in the spool file: bytes [start, end)
interface SpoolRange {
    readonly start: number;
    readonly end: number;
    released: boolean;
}

// Worksheets that come before the parts they depend on are spooled to a single temp file, and read back range by range
// when the consumer reads them: one file descriptor whatever the number of worksheets, and no I/O for the worksheets
// that are skipped. The file stays open, so that a consumer can collect the readers and read them after the iteration
// (when the temp dir is gone: unlink semantics of POSIX and of Windows' POSIX delete, which libuv uses). It is closed
// when the iteration stops early, or once every worksheet has been read or its reader garbage collected.
class SheetSpool {
    readonly handle: fs.promises.FileHandle;
    size: number;
    // ranges that may still be read
    pending: number;
    // no more ranges will be added
    complete: boolean;
    closing: Promise<void> | null;

    constructor(handle: fs.promises.FileHandle) {
        this.handle = handle;
        this.size = 0;
        this.pending = 0;
        this.complete = false;
        this.closing = null;
    }

    static async create(dir: string): Promise<SheetSpool> {
        return new SheetSpool(await fs.promises.open(path.join(dir, "worksheets.xml"), "w+"));
    }

    // writes the data of a zip entry; returns its range
    async append(entry: AsyncIterable<Uint8Array>): Promise<SpoolRange> {
        const start = this.size;
        for await (const chunk of entry) {
            for (let offset = 0; offset < chunk.length;) {
                const { bytesWritten } = await this.handle.write(chunk, offset, chunk.length - offset, this.size);
                offset += bytesWritten;
                this.size += bytesWritten;
            }
        }
        this.pending++;
        return { start, end: this.size, released: false };
    }

    // the data of a range, read on demand
    async *read(range: SpoolRange): AsyncGenerator<Buffer, void, undefined> {
        try {
            for (let position = range.start; position < range.end;) {
                const buffer = Buffer.allocUnsafe(Math.min(SPOOL_CHUNK, range.end - position));
                const { bytesRead } = await this.handle.read(buffer, 0, buffer.length, position);
                if (!bytesRead) {
                    throw new Error("Unexpected end of a spooled worksheet");
                }
                position += bytesRead;
                yield buffer.subarray(0, bytesRead);
            }
        } finally {
            // a failure to close the file rejects the read of the last worksheet
            await this.release(range);
        }
    }

    // the range will not be read (again); returns the closing of the file, if any
    release(range: SpoolRange): Promise<void> | null {
        if (!range.released) {
            range.released = true;
            this.pending--;
        }
        return this._closeWhenDone();
    }

    finish(): Promise<void> | null {
        this.complete = true;
        return this._closeWhenDone();
    }

    close(): Promise<void> {
        if (!this.closing) {
            this.closing = this.handle.close();
        }
        return this.closing;
    }

    private _closeWhenDone(): Promise<void> | null {
        return this.complete && !this.pending ? this.close() : this.closing;
    }
}

// A spooled worksheet that the consumer dropped without reading it (skipped) can never be read: its range is released
// once the reader is garbage collected, so that the spool file does not stay open for the life of the process. Closing
// the file is best-effort cleanup there: nothing awaits it, so a failure is ignored instead of becoming an unhandled
// rejection (the file was already removed with its temp dir).
const spooledSheets = new FinalizationRegistry<{ spool: SheetSpool; range: SpoolRange }>(({ spool, range }) => {
    const closing = spool.release(range);
    if (closing) {
        closing.catch(() => {});
    }
});

const DEFAULT_OPTIONS: Readonly<Pick<WorkbookReaderOptions, "worksheets" | "sharedStrings" | "hyperlinks" | "styles" | "entries">> = {
    worksheets: "emit",
    sharedStrings: "cache",
    hyperlinks: "ignore",
    styles: "ignore",
    entries: "ignore",
};

// the names of the worksheet parts and of their relationships (the number of the part is captured)
const WORKSHEET_PART = /xl\/worksheets\/sheet(\d+)[.]xml/;
const WORKSHEET_RELS_PART = /xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/;

// the byte chunks of the file (Buffers, Uint8Arrays...)
type ByteChunks = AsyncIterable<unknown, unknown, undefined>;

// a loosely typed input, to tell its kind
type InputCandidate = Partial<ByteChunks & { getReader: unknown; pipe: unknown }> | null | undefined;

/**
 * The streaming reader of an xlsx file: reads the parts as they come out of the zip, in one pass, holding only the
 * shared strings (when cached), the styles (when cached) and a worksheet at a time (the rows are not kept).
 *
 * Iterate it (`for await (const worksheet of reader)`), then each worksheet for its rows; or call {@link read} and
 * listen to the events (`worksheet`, `shared-strings`, `hyperlinks`, `entry`, `end`, `finished`, `error`). A worksheet
 * that comes before the parts it depends on (shared strings, workbook) is spooled to a temp file.
 */
class WorkbookReader extends EventEmitter<WorkbookReaderEvents> {
    /** For reference: the valid values of the options. */
    static readonly Options = {
        worksheets: ["emit", "ignore"],
        sharedStrings: ["cache", "emit", "ignore"],
        hyperlinks: ["cache", "emit", "ignore"],
        styles: ["cache", "ignore"],
        entries: ["emit", "ignore"],
    };

    /** The input given to the constructor (read when {@link parse} gets none). */
    input: WorkbookReaderInput | undefined;
    /** The options (defaults applied: see {@link WorkbookReaderOptions}). */
    options: WorkbookReaderOptions;
    /** The styles: the default ones, replaced by those of the file with `styles: "cache"`. */
    styles: ReaderStyles;
    /** The stream being read. */
    stream?: ByteChunks;
    /** `sharedStrings: "cache"`: the shared strings table. */
    sharedStrings?: SharedStringValue[];
    /** The relationships of the workbook part. */
    workbookRels?: RelationshipModel[] | null;
    /** The workbook part. */
    model?: WorkbookPartModel | null;
    /** The workbook properties (`date1904`), as the workbookPr xform holds them. */
    properties?: { readonly model?: Partial<WorkbookProperties> | null };

    private _sheetIndex?: Map<string | undefined, SheetDefModel | undefined> | null;
    private _stylesParsed?: boolean;

    constructor(input: WorkbookReaderInput | undefined, options: Partial<WorkbookReaderOptions> = {}) {
        super();

        this.input = input;

        this.options = { ...DEFAULT_OPTIONS, ...options };

        this.styles = new StylesXform();
        this.styles.init();
    }

    // a file name, or any readable: a Node stream (or a readable-stream / older one), a web ReadableStream or an async
    // iterable of byte chunks
    private _getStream(input: WorkbookReaderInput | undefined): ByteChunks {
        if (typeof input === "string") {
            return fs.createReadStream(input);
        }
        const candidate = input as InputCandidate;
        if (candidate && typeof candidate[Symbol.asyncIterator] === "function") {
            return candidate as ByteChunks;
        }
        if (candidate && typeof candidate.getReader === "function") {
            return Readable.fromWeb(candidate as WebReadableStream);
        }
        if (candidate && typeof candidate.pipe === "function") {
            // pipeline (unlike pipe) forwards the errors of the input
            return pipeline(candidate as NodeJS.ReadableStream, new PassThrough(), () => {});
        }
        // eslint-disable-next-line @typescript-eslint/no-base-to-string -- the input as it always was in this message
        throw new Error(`Could not recognise input: ${String(input)}`);
    }

    /**
     * Reads the whole file, emitting the events of the parts (and reading each worksheet and hyperlinks part emitted),
     * then `end` and `finished`.
     *
     * @param input - replaces the constructor's
     * @param options - replace the constructor's, on top of the defaults
     * @returns resolves once the file is read; errors are emitted (`error`), not thrown
     */
    async read(input?: WorkbookReaderInput, options?: Partial<WorkbookReaderOptions>): Promise<void> {
        try {
            for await (const event of this.parse(input, options)) {
                switch (event.eventType) {
                    case "shared-strings":
                        this.emit(event.eventType, event.value);
                        break;
                    case "worksheet":
                        this.emit(event.eventType, event.value);
                        await event.value.read();
                        break;
                    case "hyperlinks":
                        this.emit(event.eventType, event.value);
                        // the zip moves on to the next entry once this one is read
                        await event.value.read();
                        break;
                }
            }
            this.emit("end");
            this.emit("finished");
        } catch (error) {
            this.emit("error", error);
        }
    }

    /** The worksheets of the file (`worksheets: "emit"`), each to be read before the next one comes. */
    async *[Symbol.asyncIterator](): AsyncGenerator<WorksheetReader, void, undefined> {
        for await (const event of this.parse()) {
            if (event.eventType === "worksheet") {
                yield event.value;
            }
        }
    }

    /**
     * Reads the parts of the file as they come.
     *
     * @param input - replaces the constructor's
     * @param options - replace the constructor's, on top of the defaults
     * @returns the worksheets, shared strings and hyperlinks, as the options select them
     * @throws Error on an unrecognised input, an invalid zip, a decompression limit exceeded, or invalid XML
     */
    async *parse(
        input?: WorkbookReaderInput,
        options?: Partial<WorkbookReaderOptions>
    ): AsyncGenerator<WorkbookReaderEvent, void, undefined> {
        // options given here replace the constructor's, on top of the defaults
        if (options) this.options = { ...DEFAULT_OPTIONS, ...options };
        const stream = (this.stream = this._getStream(input || this.input));

        // worksheets, deferred for parsing after shared strings reading
        const waitingWorkSheets: { sheetNo: string; range: SpoolRange }[] = [];
        let completed = false;
        let tmpDir: string | null = null;
        let spool: SheetSpool | null = null;
        try {
            // stopping the iteration early releases the input (a stream is destroyed); input errors
            // reject it; the data of the entries left unread is skipped
            for await (const entry of parseZipStream(stream, this.options)) {
                switch (entry.path) {
                    case "_rels/.rels":
                        break;
                    case "xl/_rels/workbook.xml.rels":
                        await this._parseRels(entry);
                        break;
                    case "xl/workbook.xml":
                        await this._parseWorkbook(entry);
                        break;
                    case "xl/sharedStrings.xml":
                        yield* this._parseSharedStrings(entry);
                        break;
                    case "xl/styles.xml":
                        await this._parseStyles(entry);
                        break;
                    default: {
                        let match = WORKSHEET_PART.exec(entry.path);
                        if (match) {
                            const sheetNo = match[1]!;
                            if (this._worksheetDependenciesParsed() || this.options.worksheets !== "emit") {
                                // ignored worksheets are not read: nothing to spool
                                yield* this._parseWorksheet(entry, sheetNo);
                            } else {
                                // defer worksheet until the parts it depends on are known: spool to a temp file
                                if (!spool) {
                                    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-"));
                                    spool = await SheetSpool.create(tmpDir);
                                }
                                // rejects on entry errors too (truncated or corrupt zip data)
                                waitingWorkSheets.push({ sheetNo, range: await spool.append(entry) });
                            }
                        } else if ((match = WORKSHEET_RELS_PART.exec(entry.path))) {
                            yield* this._parseHyperlinks(entry, match[1]!);
                        }
                        break;
                    }
                }
            }

            for (const { sheetNo, range } of waitingWorkSheets) {
                // a range was spooled: the spool exists
                const sheetSpool = spool!;
                // read when the consumer reads the worksheet, during or after the iteration
                yield* this._parseWorksheet(sheetSpool.read(range), sheetNo, (reader) =>
                    spooledSheets.register(reader, { spool: sheetSpool, range })
                );
            }
            completed = true;
        } finally {
            try {
                if (spool) {
                    // stopped early (break / throw / error): nobody will read the worksheets, close the file now; else
                    // once they are read. Closed before the temp dir is removed (Windows cannot remove a file being closed)
                    await (completed ? spool.finish() : spool.close());
                }
            } finally {
                // also runs when the consumer stops iterating early (break / throw)
                if (tmpDir) {
                    await fs.promises.rm(tmpDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
                }
            }
        }
    }

    // shared strings, sheet names (workbook rels + workbook) and, when cached, styles
    private _worksheetDependenciesParsed(): boolean {
        const stylesReady = this.options.styles !== "cache" || this._stylesParsed;
        return Boolean(this.sharedStrings && this.workbookRels && this.model && stylesReady);
    }

    private _emitEntry(payload: WorkbookReaderEntry): void {
        if (this.options.entries === "emit") {
            this.emit("entry", payload);
        }
    }

    private async _parseRels(entry: ZipStreamEntry): Promise<void> {
        const xform = new RelationshipsXform();
        this.workbookRels = await xform.parseStream(entry);
        this._sheetIndex = null;
    }

    private async _parseWorkbook(entry: ZipStreamEntry): Promise<void> {
        this._emitEntry({ type: "workbook" });

        const workbook = new WorkbookXform();
        await workbook.parseStream(entry);

        // the <workbookPr> xform itself, whose `model` is the parsed properties (the shape the reader always exposed)
        this.properties = workbook.map.workbookPr;
        this.model = workbook.model;
        this._sheetIndex = null;
    }

    // the <sheet> of each worksheet part, from the workbook rels and workbook.xml (built once: a lookup per worksheet
    // part in the lists would be quadratic)
    private _sheetOfPart(sheetPath: string): SheetDefModel | undefined {
        if (!this._sheetIndex) {
            const sheetsByRId = new Map<string, SheetDefModel>();
            for (const sheet of (this.model && this.model.sheets) || []) {
                // a <sheet> without r:id matches no relationship
                if (sheet.rId !== undefined && !sheetsByRId.has(sheet.rId)) {
                    sheetsByRId.set(sheet.rId, sheet);
                }
            }
            // the first relationship of each part, as before
            const sheetIndex = (this._sheetIndex = new Map());
            for (const rel of this.workbookRels || []) {
                const partPath = workbookPartPath(rel.Target);
                if (!sheetIndex.has(partPath)) {
                    sheetIndex.set(partPath, sheetsByRId.get(rel.Id));
                }
            }
        }
        return this._sheetIndex.get(sheetPath);
    }

    private async *_parseSharedStrings(entry: ZipStreamEntry): AsyncGenerator<WorkbookReaderEvent, void, undefined> {
        this._emitEntry({ type: "shared-strings" });
        let cache: SharedStringValue[] | null = null;
        if (this.options.sharedStrings === "cache") {
            cache = this.sharedStrings = [];
        } else if (this.options.sharedStrings !== "emit") {
            return;
        }

        // each <si> is read by the xform of the document reader: same rich text fonts, phonetic runs and _xHHHH_ escapes
        const xform = new SharedStringXform();
        let inString = false;
        let index = 0;
        for await (const events of parseSax(entry)) {
            for (const event of events) {
                if (event.eventType === "opentag") {
                    if (inString || event.value.name === "si") {
                        inString = true;
                        xform.parseOpen(event.value);
                    }
                } else if (event.eventType === "text") {
                    if (inString) {
                        xform.parseText(event.value);
                    }
                } else if (inString && !xform.parseClose(event.value.name)) {
                    inString = false;
                    // the <si> was opened: the xform has its model
                    const text = xform.model!;
                    if (cache) {
                        cache.push(text);
                    } else {
                        yield { eventType: "shared-strings", value: { index: index++, text } };
                    }
                }
            }
        }
    }

    private async _parseStyles(entry: ZipStreamEntry): Promise<void> {
        this._emitEntry({ type: "styles" });
        if (this.options.styles === "cache") {
            this.styles = new StylesXform();
            await this.styles.parseStream(entry);
            this._stylesParsed = true;
        }
    }

    private *_parseWorksheet(
        iterator: SaxInput,
        sheetNo: string,
        onCreated?: (reader: WorksheetReader) => void
    ): Generator<WorkbookReaderEvent, void, undefined> {
        this._emitEntry({ type: "worksheet", id: sheetNo });
        const worksheetReader = new WorksheetReader({
            workbook: this,
            id: sheetNo,
            iterator,
            options: this.options,
        });
        if (onCreated) {
            onCreated(worksheetReader);
        }

        const matchingSheet = this._sheetOfPart(`xl/worksheets/sheet${sheetNo}.xml`);
        if (matchingSheet) {
            worksheetReader.id = matchingSheet.id;
            worksheetReader.name = matchingSheet.name;
            worksheetReader.state = matchingSheet.state as WorksheetReader["state"];
        }
        if (this.options.worksheets === "emit") {
            yield { eventType: "worksheet", value: worksheetReader };
        }
    }

    private *_parseHyperlinks(iterator: SaxInput, sheetNo: string): Generator<WorkbookReaderEvent, void, undefined> {
        this._emitEntry({ type: "hyperlinks", id: sheetNo });
        const hyperlinksReader = new HyperlinkReader({
            workbook: this,
            id: sheetNo,
            iterator,
            options: this.options,
        });
        if (this.options.hyperlinks === "emit") {
            yield { eventType: "hyperlinks", value: hyperlinksReader };
        }
    }
}

export default WorkbookReader;
