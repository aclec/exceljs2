const fs = require("fs");
const { EventEmitter } = require("events");
const { PassThrough, Readable, pipeline } = require("stream");
const os = require("os");
const path = require("path");
const { parseZipStream } = require("../../utils/zip-reader");
const parseSax = require("../../utils/parse-sax");

const StyleManager = require("../../xlsx/xform/style/styles-xform");
const WorkbookXform = require("../../xlsx/xform/book/workbook-xform");
const RelationshipsXform = require("../../xlsx/xform/core/relationships-xform");

const WorksheetReader = require("./worksheet-reader");
const HyperlinkReader = require("./hyperlink-reader");

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
// once the reader is garbage collected, so that the spool file does not stay open for the life of the process.
const spooledSheets = new FinalizationRegistry(({ spool, range }) => spool.release(range));

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

        let text = null;
        let richText = [];
        let index = 0;
        let font = null;
        for await (const events of parseSax(entry)) {
            for (const { eventType, value } of events) {
                if (eventType === "opentag") {
                    const node = value;
                    switch (node.name) {
                        case "b":
                            font = font || {};
                            font.bold = true;
                            break;
                        case "charset":
                            font = font || {};
                            font.charset = parseInt(node.attributes.charset, 10);
                            break;
                        case "color":
                            font = font || {};
                            font.color = {};
                            if (node.attributes.rgb) {
                                font.color.argb = node.attributes.argb;
                            }
                            if (node.attributes.val) {
                                font.color.argb = node.attributes.val;
                            }
                            if (node.attributes.theme) {
                                font.color.theme = node.attributes.theme;
                            }
                            break;
                        case "family":
                            font = font || {};
                            font.family = parseInt(node.attributes.val, 10);
                            break;
                        case "i":
                            font = font || {};
                            font.italic = true;
                            break;
                        case "outline":
                            font = font || {};
                            font.outline = true;
                            break;
                        case "rFont":
                            font = font || {};
                            font.name = node.value;
                            break;
                        case "si":
                            font = null;
                            richText = [];
                            text = null;
                            break;
                        case "sz":
                            font = font || {};
                            font.size = parseInt(node.attributes.val, 10);
                            break;
                        case "strike":
                            break;
                        case "t":
                            text = null;
                            break;
                        case "u":
                            font = font || {};
                            font.underline = true;
                            break;
                        case "vertAlign":
                            font = font || {};
                            font.vertAlign = node.attributes.val;
                            break;
                    }
                } else if (eventType === "text") {
                    text = text ? text + value : value;
                } else if (eventType === "closetag") {
                    const node = value;
                    switch (node.name) {
                        case "r":
                            richText.push({
                                font,
                                text,
                            });

                            font = null;
                            text = null;
                            break;
                        case "si":
                            if (this.options.sharedStrings === "cache") {
                                this.sharedStrings.push(richText.length ? { richText } : text);
                            } else if (this.options.sharedStrings === "emit") {
                                yield {
                                    eventType: "shared-strings",
                                    value: { index: index++, text: richText.length ? { richText } : text },
                                };
                            }

                            richText = [];
                            font = null;
                            text = null;
                            break;
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

module.exports = WorkbookReader;
