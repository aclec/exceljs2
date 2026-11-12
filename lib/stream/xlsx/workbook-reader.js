const fs = require("fs");
const { EventEmitter } = require("events");
const { PassThrough, Readable, pipeline } = require("stream");
const { pipeline: pipelinePromise } = require("stream/promises");
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

function whenClosed(stream) {
    return stream.closed ? Promise.resolve() : new Promise((resolve) => stream.once("close", resolve));
}

// open, or closed without opening (destroyed early, see sheetFiles, or failed to open)
function whenOpen(stream) {
    if (!stream.pending || stream.closed) {
        return Promise.resolve();
    }
    return new Promise((resolve) => {
        stream.once("ready", resolve);
        stream.once("close", resolve);
    });
}

// The temp file of a spooled worksheet is opened when its WorksheetReader is created, so that a consumer can collect
// the readers and read them after the iteration (when the temp dir is gone). A reader that the consumer dropped
// without reading it (skipped worksheet, worksheets: 'ignore') can never be read: its file is closed once the reader is
// garbage collected, instead of staying open for the life of the process.
const sheetFiles = new FinalizationRegistry((stream) => stream.destroy());

// Removes the temp dir of the spooled worksheets. Windows cannot remove a file while it is being closed,
// so the streams that are done (read to the end, or destroyed) are awaited until closed. The files of
// worksheets not read yet (collected by the consumer to be read later, or skipped) are removed once
// their streams have opened them: their data stays readable through the open file (unlink semantics
// of POSIX and of Windows' POSIX delete, which libuv uses), as before.
async function removeTempDir(dir, sheetStreams) {
    const unread = sheetStreams.filter((stream) => !stream.closed && !stream.destroyed && !stream.readableEnded);
    await Promise.all(sheetStreams.filter((stream) => !unread.includes(stream)).map(whenClosed));
    await Promise.all(unread.map(whenOpen));
    await fs.promises.rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
}

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
        const sheetStreams = [];
        let completed = false;
        let tmpDir = null;
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
                            if (this._worksheetDependenciesParsed()) {
                                yield* this._parseWorksheet(entry, sheetNo);
                            } else {
                                // defer worksheet until the parts it depends on are known: spool to a temp file
                                if (!tmpDir) {
                                    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-"));
                                }
                                const tmpPath = path.join(tmpDir, `${waitingWorkSheets.length}.xml`);
                                waitingWorkSheets.push({ sheetNo, path: tmpPath });
                                // rejects on entry errors too (truncated or corrupt zip data)
                                await pipelinePromise(entry, fs.createWriteStream(tmpPath));
                            }
                        } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
                            match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
                            sheetNo = match[1];
                            yield* this._parseHyperlinks(entry, sheetNo);
                        }
                        break;
                }
            }

            for (const { sheetNo, path: tmpPath } of waitingWorkSheets) {
                // opened now: a consumer may collect the worksheets and read them after the loop
                const sheetStream = fs.createReadStream(tmpPath);
                sheetStreams.push(sheetStream);
                yield* this._parseWorksheet(sheetStream, sheetNo, (reader) => sheetFiles.register(reader, sheetStream));
            }
            completed = true;
        } finally {
            if (!completed) {
                // stopped early (break / throw / error): nobody will read them, release the file handles
                sheetStreams.forEach((sheetStream) => sheetStream.destroy());
            }
            // also runs when the consumer stops iterating early (break / throw)
            if (tmpDir) {
                await removeTempDir(tmpDir, sheetStreams);
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
