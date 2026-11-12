/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import * as fs from "fs";
import { EventEmitter } from "events";
import { PassThrough, Readable, pipeline } from "stream";
import * as os from "os";
import * as path from "path";
import { parseZipStream } from "../../utils/zip-reader";
import parseSax from "../../utils/parse-sax";

import StyleManager from "../../xlsx/xform/style/styles-xform";
import WorkbookXform from "../../xlsx/xform/book/workbook-xform";
import RelationshipsXform from "../../xlsx/xform/core/relationships-xform";
import SharedStringXform from "../../xlsx/xform/strings/shared-string-xform";

import WorksheetReader from "./worksheet-reader";
import HyperlinkReader from "./hyperlink-reader";

// workbook.xml.rels targets are relative to xl/ or absolute from the package root (openpyxl)
function workbookPartPath(target) {
    if (typeof target !== "string") {
        return undefined;
    }
    return target.startsWith("/") ? target.slice(1) : `xl/${target}`;
}

// size of the chunks read back from the spool file
const SPOOL_CHUNK = 64 * 1024;

// Worksheets that come before the parts they depend on are spooled to a single temp file, and read back range by range
// when the consumer reads them: one file descriptor whatever the number of worksheets, and no I/O for the worksheets
// that are skipped. The file stays open, so that a consumer can collect the readers and read them after the iteration
// (when the temp dir is gone: unlink semantics of POSIX and of Windows' POSIX delete, which libuv uses). It is closed
// when the iteration stops early, or once every worksheet has been read or its reader garbage collected.
class SheetSpool {
    constructor(handle) {
        this.handle = handle;
        this.size = 0;
        // ranges that may still be read
        this.pending = 0;
        // no more ranges will be added
        this.complete = false;
        this.closing = null;
    }

    static async create(dir) {
        return new SheetSpool(await fs.promises.open(path.join(dir, "worksheets.xml"), "w+"));
    }

    // writes the data of a zip entry; returns its range
    async append(entry) {
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
    async *read(range) {
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
    release(range) {
        if (!range.released) {
            range.released = true;
            this.pending--;
        }
        return this._closeWhenDone();
    }

    finish() {
        this.complete = true;
        return this._closeWhenDone();
    }

    close() {
        if (!this.closing) {
            this.closing = this.handle.close();
        }
        return this.closing;
    }

    _closeWhenDone() {
        return this.complete && !this.pending ? this.close() : this.closing;
    }
}

// A spooled worksheet that the consumer dropped without reading it (skipped) can never be read: its range is released
// once the reader is garbage collected, so that the spool file does not stay open for the life of the process. Closing
// the file is best-effort cleanup there: nothing awaits it, so a failure is ignored instead of becoming an unhandled
// rejection (the file was already removed with its temp dir).
const spooledSheets = new FinalizationRegistry(({ spool, range }) => {
    const closing = spool.release(range);
    if (closing) {
        closing.catch(() => {});
    }
});

const DEFAULT_OPTIONS = {
    worksheets: "emit",
    sharedStrings: "cache",
    hyperlinks: "ignore",
    styles: "ignore",
    entries: "ignore",
};

class WorkbookReader extends EventEmitter {
    constructor(input, options = {}) {
        super();

        this.input = input;

        this.options = { ...DEFAULT_OPTIONS, ...options };

        this.styles = new StyleManager();
        this.styles.init();
    }

    // a file name, or any readable: a Node stream (or a readable-stream / older one), a web
    // ReadableStream or an async iterable of byte chunks
    _getStream(input) {
        if (typeof input === "string") {
            return fs.createReadStream(input);
        }
        if (input && typeof input[Symbol.asyncIterator] === "function") {
            return input;
        }
        if (input && typeof input.getReader === "function") {
            return Readable.fromWeb(input);
        }
        if (input && typeof input.pipe === "function") {
            // pipeline (unlike pipe) forwards the errors of the input
            return pipeline(input, new PassThrough(), () => {});
        }
        throw new Error(`Could not recognise input: ${input}`);
    }

    async read(input, options) {
        try {
            for await (const { eventType, value } of this.parse(input, options)) {
                switch (eventType) {
                    case "shared-strings":
                        this.emit(eventType, value);
                        break;
                    case "worksheet":
                        this.emit(eventType, value);
                        await value.read();
                        break;
                    case "hyperlinks":
                        this.emit(eventType, value);
                        // the zip moves on to the next entry once this one is read
                        await value.read();
                        break;
                }
            }
            this.emit("end");
            this.emit("finished");
        } catch (error) {
            this.emit("error", error);
        }
    }

    async *[Symbol.asyncIterator]() {
        for await (const { eventType, value } of this.parse()) {
            if (eventType === "worksheet") {
                yield value;
            }
        }
    }

    async *parse(input, options) {
        // options given here replace the constructor's, on top of the defaults
        if (options) this.options = { ...DEFAULT_OPTIONS, ...options };
        const stream = (this.stream = this._getStream(input || this.input));

        // worksheets, deferred for parsing after shared strings reading
        const waitingWorkSheets = [];
        let completed = false;
        let tmpDir = null;
        let spool = null;
        try {
            // stopping the iteration early releases the input (a stream is destroyed); input errors
            // reject it; the data of the entries left unread is skipped
            for await (const entry of parseZipStream(stream, this.options)) {
                let match;
                let sheetNo;
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
                    default:
                        if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
                            match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
                            sheetNo = match[1];
                            if (this._worksheetDependenciesParsed() || this.options.worksheets !== "emit") {
                                // ignored worksheets are not read: nothing to spool
                                yield* this._parseWorksheet(entry, sheetNo);
                            } else {
                                // defer worksheet until the parts it depends on are known: spool to a temp file
                                if (!tmpDir) {
                                    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-"));
                                    spool = await SheetSpool.create(tmpDir);
                                }
                                // rejects on entry errors too (truncated or corrupt zip data)
                                waitingWorkSheets.push({ sheetNo, range: await spool.append(entry) });
                            }
                        } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
                            match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
                            sheetNo = match[1];
                            yield* this._parseHyperlinks(entry, sheetNo);
                        }
                        break;
                }
            }

            for (const { sheetNo, range } of waitingWorkSheets) {
                // read when the consumer reads the worksheet, during or after the iteration
                yield* this._parseWorksheet(spool.read(range), sheetNo, (reader) => spooledSheets.register(reader, { spool, range }));
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
    _worksheetDependenciesParsed() {
        const stylesReady = this.options.styles !== "cache" || this._stylesParsed;
        return Boolean(this.sharedStrings && this.workbookRels && this.model && stylesReady);
    }

    _emitEntry(payload) {
        if (this.options.entries === "emit") {
            this.emit("entry", payload);
        }
    }

    async _parseRels(entry) {
        const xform = new RelationshipsXform();
        this.workbookRels = await xform.parseStream(entry);
        this._sheetIndex = null;
    }

    async _parseWorkbook(entry) {
        this._emitEntry({ type: "workbook" });

        const workbook = new WorkbookXform();
        await workbook.parseStream(entry);

        this.properties = workbook.map.workbookPr;
        this.model = workbook.model;
        this._sheetIndex = null;
    }

    // the <sheet> of each worksheet part, from the workbook rels and workbook.xml (built once: a lookup per worksheet
    // part in the lists would be quadratic)
    _sheetOfPart(sheetPath) {
        if (!this._sheetIndex) {
            const sheetsByRId = new Map();
            ((this.model && this.model.sheets) || []).forEach((sheet) => {
                if (!sheetsByRId.has(sheet.rId)) {
                    sheetsByRId.set(sheet.rId, sheet);
                }
            });
            // the first relationship of each part, as before
            this._sheetIndex = new Map();
            (this.workbookRels || []).forEach((rel) => {
                const partPath = workbookPartPath(rel.Target);
                if (!this._sheetIndex.has(partPath)) {
                    this._sheetIndex.set(partPath, sheetsByRId.get(rel.Id));
                }
            });
        }
        return this._sheetIndex.get(sheetPath);
    }

    async *_parseSharedStrings(entry) {
        this._emitEntry({ type: "shared-strings" });
        switch (this.options.sharedStrings) {
            case "cache":
                this.sharedStrings = [];
                break;
            case "emit":
                break;
            default:
                return;
        }

        // each <si> is read by the xform of the document reader: same rich text fonts, phonetic runs and _xHHHH_ escapes
        const xform = new SharedStringXform();
        let inString = false;
        let index = 0;
        for await (const events of parseSax(entry)) {
            for (const { eventType, value } of events) {
                if (eventType === "opentag") {
                    if (inString || value.name === "si") {
                        inString = true;
                        xform.parseOpen(value);
                    }
                } else if (eventType === "text") {
                    if (inString) {
                        xform.parseText(value);
                    }
                } else if (eventType === "closetag" && inString && !xform.parseClose(value.name)) {
                    inString = false;
                    if (this.options.sharedStrings === "cache") {
                        this.sharedStrings.push(xform.model);
                    } else {
                        yield { eventType: "shared-strings", value: { index: index++, text: xform.model } };
                    }
                }
            }
        }
    }

    async _parseStyles(entry) {
        this._emitEntry({ type: "styles" });
        if (this.options.styles === "cache") {
            this.styles = new StyleManager();
            await this.styles.parseStream(entry);
            this._stylesParsed = true;
        }
    }

    *_parseWorksheet(iterator, sheetNo, onCreated) {
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
            worksheetReader.state = matchingSheet.state;
        }
        if (this.options.worksheets === "emit") {
            yield { eventType: "worksheet", value: worksheetReader };
        }
    }

    *_parseHyperlinks(iterator, sheetNo) {
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

// for reference - these are the valid values for options
WorkbookReader.Options = {
    worksheets: ["emit", "ignore"],
    sharedStrings: ["cache", "emit", "ignore"],
    hyperlinks: ["cache", "emit", "ignore"],
    styles: ["cache", "ignore"],
    entries: ["emit", "ignore"],
};

export default WorkbookReader;
