// Zip container writing. Compression and CRC-32 come from deflate.js (node:zlib on Node.js and Bun,
// fflate in the browser); the records are written here (fflate's Zip has no ZIP64 support).
//
// Entries are written in the order they are created. Every entry is compressed as soon as its data
// arrives; while an earlier entry is still open, the compressed data waits in memory.
// - An entry whose data is complete before it is written (append) gets its sizes in the local header.
// - A streamed entry (openStream) that is still open sets general purpose flag bit 3 and is followed by
//   a data descriptor (with 64-bit sizes when they do not fit in 32 bits, like archiver and Go do).
// ZIP64 records are added when a size, an offset or the number of entries does not fit in the classic
// fields, or for every entry and the end of central directory with the forceZip64 option.
//
// Backpressure: drained() settles once the output stream (pipe) has drained and at most HIGH_WATER_MARK bytes are
// being compressed. Data of an entry waiting for an earlier, still open entry is not counted: it only leaves memory
// when that entry ends.
const { EventEmitter } = require("events");
const { createCompressor, selectBackend } = require("./deflate");

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_DATA_DESCRIPTOR = 0x08074b50;

const MAX_U16 = 0xffff;
const MAX_U32 = 0xffffffff;
const TWO_32 = 0x100000000;

const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;

const VERSION = 20;
const VERSION_ZIP64 = 45;

// values >= size, or `entries` entries or more, need ZIP64 records (only tests lower them: _zip64Limits)
const ZIP64_LIMITS = { size: MAX_U32, entries: MAX_U16 };

// uncompressed bytes being compressed (asynchronously) above which drained() waits
const HIGH_WATER_MARK = 4 * 1024 * 1024;

const textEncoder = new TextEncoder();
const EMPTY = new Uint8Array(0);

// -----------------------------------------------------------------------------
// options

// The zip options are {level, forceZip64, comment}. The shapes of the zip libraries used before v5
// are still accepted: archiver's {zlib: {level}, store} and JSZip's {compression, compressionOptions}.
function normalizeZipOptions(options) {
    const zip = options || {};
    let level = 6;
    if (zip.level !== undefined) {
        ({ level } = zip);
    } else if (zip.zlib && zip.zlib.level !== undefined) {
        ({ level } = zip.zlib);
    } else if (zip.compressionOptions && zip.compressionOptions.level !== undefined) {
        ({ level } = zip.compressionOptions);
    }
    if (zip.store === true || zip.compression === "STORE") {
        level = 0;
    }
    if (!Number.isInteger(level) || level < 0 || level > 9) {
        throw new TypeError(`zip level must be an integer from 0 (no compression) to 9, got ${level}`);
    }
    const comment = zip.comment === undefined ? EMPTY : textEncoder.encode(String(zip.comment));
    if (comment.length > MAX_U16) {
        throw new RangeError("zip comment is longer than 65535 bytes");
    }
    return {
        level,
        forceZip64: Boolean(zip.forceZip64),
        comment,
        zip64Limits: Object.assign({}, ZIP64_LIMITS, zip._zip64Limits),
        // internal, for tests: "zlib" or "fflate"
        backend: selectBackend(zip._backend),
    };
}

// -----------------------------------------------------------------------------
// binary helpers

function writeU16(data, pos, value) {
    data[pos] = value & 0xff;
    data[pos + 1] = (value >>> 8) & 0xff;
}

function writeU32(data, pos, value) {
    data[pos] = value & 0xff;
    data[pos + 1] = (value >>> 8) & 0xff;
    data[pos + 2] = (value >>> 16) & 0xff;
    data[pos + 3] = (value >>> 24) & 0xff;
}

function writeU64(data, pos, value) {
    writeU32(data, pos, value % TWO_32);
    writeU32(data, pos + 4, Math.floor(value / TWO_32));
}

function dosDateTime(date) {
    const year = Math.min(Math.max(date.getFullYear(), 1980), 2107);
    return {
        time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
        date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    };
}

// ZIP64 extended information extra field
function zip64Extra(values) {
    const extra = new Uint8Array(4 + 8 * values.length);
    writeU16(extra, 0, 1);
    writeU16(extra, 2, 8 * values.length);
    values.forEach((value, index) => writeU64(extra, 4 + 8 * index, value));
    return extra;
}

function toBytes(data) {
    if (typeof data === "string") {
        return textEncoder.encode(data);
    }
    if (data instanceof Uint8Array) {
        return data;
    }
    if (data instanceof ArrayBuffer) {
        return new Uint8Array(data);
    }
    if (ArrayBuffer.isView(data)) {
        return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    }
    // StringBuf (lib/utils/string-buf.js)
    if (data && typeof data.toBuffer === "function") {
        return data.toBuffer();
    }
    throw new TypeError("Zip entry data must be a string, a Buffer, a Uint8Array or an ArrayBuffer");
}

function toBuffer(data) {
    return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
}

// leading slashes are dropped (like archiver did); anything that could extract outside of the target
// directory is refused (zip slip)
function entryName(name) {
    const normalised = String(name).replace(/^\/+/, "");
    if (!normalised || normalised.includes("\\") || normalised.split("/").some((part) => part === ".." || part === ".")) {
        throw new Error(`Invalid zip entry name: ${JSON.stringify(String(name))}`);
    }
    return normalised;
}

// -----------------------------------------------------------------------------

// One entry. Also the writable that the streaming writers write a part to: write(data)... end().
// Emits 'finish' when ended and 'zipped' once all its data is in the archive output.
class ZipEntry extends EventEmitter {
    constructor(zip, name) {
        super();
        this.zip = zip;
        this.name = entryName(name);
        this.nameBytes = textEncoder.encode(this.name);
        if (this.nameBytes.length > MAX_U16) {
            throw new Error(`Zip entry name is too long: ${this.name}`);
        }
        this.flags = this.nameBytes.length === this.name.length ? 0 : FLAG_UTF8;
        this.stored = zip.level === 0;
        this.compressor = createCompressor(zip.level, zip.backend, (error, chunks, final) => this._onData(error, chunks, final));
        this.compressedSize = 0;
        // compressed data waiting for the earlier entries
        this.pending = [];
        this.ended = false;
        this.finished = false;
        this.headerWritten = false;
    }

    get method() {
        return this.stored ? 0 : 8;
    }

    write(data) {
        if (this.ended) {
            throw new Error(`Write after end of the zip entry ${this.name}`);
        }
        const bytes = toBytes(data);
        if (bytes.length) {
            this._compress(bytes, false);
        }
        return true;
    }

    end(data) {
        if (this.ended) {
            return;
        }
        this.ended = true;
        this._compress(data === undefined ? EMPTY : toBytes(data), true);
        this.emit("finish");
    }

    _compress(bytes, final) {
        try {
            this.compressor.push(bytes, final);
        } catch (error) {
            this.zip._fail(error);
        }
    }

    _onData(error, chunks, final) {
        if (error) {
            this.zip._fail(error);
            return;
        }
        for (const chunk of chunks) {
            this.compressedSize += chunk.length;
            if (this.headerWritten) {
                this.zip._output(chunk);
            } else {
                this.pending.push(chunk);
            }
        }
        if (final) {
            this.crc = this.compressor.crc;
            this.size = this.compressor.size;
            this.finished = true;
            this.zip._pump();
        }
        this.zip._checkDrained();
    }
}

class ZipWriter extends EventEmitter {
    constructor(options) {
        super();
        const { level, forceZip64, comment, zip64Limits, backend } = normalizeZipOptions(options);
        this.level = level;
        this.backend = backend;
        this.forceZip64 = forceZip64;
        this.comment = comment;
        this.limits = zip64Limits;
        this.dosDateTime = dosDateTime(new Date());
        this.entries = [];
        // index of the entry being written to the output
        this.current = 0;
        // bytes written to the output
        this.offset = 0;
        this.finalizing = false;
        this.done = false;
        this.error = null;
        // the output stream asked to wait for 'drain'
        this.outputFull = false;
        // the promise drained() returns while the output is backlogged, and its resolve / reject functions
        this.drainPromise = null;
        this.drainWaiter = null;
        // settles when the archive is complete, or as soon as it fails
        this.finished = new Promise((resolve, reject) => {
            this._resolve = resolve;
            this._reject = reject;
        });
        // a failure is reported by finalize() (and the methods called after it), or through `finished`
        this.finished.catch(() => {});
    }

    // the output: 'data' events (Buffers), then 'end'
    pipe(destination) {
        // backpressure needs the 'drain' event
        const drains = typeof destination.on === "function";
        this.on("data", (chunk) => {
            if (destination.write(chunk) === false && drains) {
                this.outputFull = true;
            }
        });
        if (drains) {
            // an error of the destination while the archive is written fails it; once it is complete, the
            // destination's errors are its owner's again (the listener would hide them)
            const onError = (error) => this._fail(error);
            const onDrain = () => {
                this.outputFull = false;
                this._checkDrained();
            };
            // destroyed without an error (e.g. an HTTP response whose client went away): 'drain' will never come
            const onClose = () => this._fail(new Error("The output stream was closed before the zip archive was complete"));
            destination.on("error", onError);
            destination.on("drain", onDrain);
            destination.on("close", onClose);
            this.once("end", () => {
                destination.removeListener("error", onError);
                destination.removeListener("drain", onDrain);
                destination.removeListener("close", onClose);
            });
        }
        this.on("end", () => destination.end());
        return destination;
    }

    // adds an entry with all of its data: {name, base64}
    append(data, options) {
        this._check();
        // converted first: an invalid value must not leave an entry that never ends
        const bytes = options.base64 ? Buffer.from(data, "base64") : toBytes(data);
        const entry = new ZipEntry(this, options.name);
        this.entries.push(entry);
        entry.end(bytes);
        this._pump();
    }

    // adds an entry whose data is written later through the returned ZipEntry
    openStream(name) {
        this._check();
        const entry = new ZipEntry(this, name);
        this.entries.push(entry);
        this._pump();
        return entry;
    }

    // writes the central directory once every entry is complete; resolves when the output has ended
    finalize() {
        if (!this.finalizing && !this.error) {
            this.finalizing = true;
            this._pump();
        }
        return this.finished;
    }

    // a promise that resolves once the output can take more data (see Backpressure above), or rejects with the error
    // of a failed archive. It is marked handled: callers may ignore it.
    drained() {
        if (this.error) {
            return this.finished;
        }
        if (!this._backlogged()) {
            return Promise.resolve();
        }
        // one promise for all the callers until the backlog clears: writers that do not wait allocate nothing
        if (!this.drainPromise) {
            this.drainPromise = new Promise((resolve, reject) => {
                this.drainWaiter = { resolve, reject };
            });
            this.drainPromise.catch(() => {});
        }
        return this.drainPromise;
    }

    _backlogged() {
        if (this.done) {
            return false;
        }
        if (this.outputFull) {
            return true;
        }
        let inFlight = 0;
        for (let i = this.current; i < this.entries.length; i++) {
            inFlight += this.entries[i].compressor.inFlight;
        }
        return inFlight > HIGH_WATER_MARK;
    }

    _checkDrained() {
        if (this.drainPromise && (this.error || !this._backlogged())) {
            const { resolve, reject } = this.drainWaiter;
            this.drainPromise = null;
            this.drainWaiter = null;
            if (this.error) {
                reject(this.error);
            } else {
                resolve();
            }
        }
    }

    _check() {
        if (this.error) {
            throw this.error;
        }
        if (this.finalizing) {
            throw new Error("Cannot add zip entries after finalize()");
        }
    }

    _fail(error) {
        if (!this.error && !this.done) {
            this.error = error;
            this._reject(error);
            this._checkDrained();
        }
    }

    _output(data) {
        // nothing more for a failed destination
        if (this.error) {
            return;
        }
        this.offset += data.length;
        try {
            this.emit("data", toBuffer(data));
        } catch (error) {
            this._fail(error);
        }
    }

    // writes what can be written; re-entrant calls (e.g. from a 'zipped' listener) are folded into the
    // running one
    _pump() {
        if (this.pumping) {
            this.pumpAgain = true;
            return;
        }
        this.pumping = true;
        try {
            do {
                this.pumpAgain = false;
                this._pumpOnce();
            } while (this.pumpAgain);
        } catch (error) {
            this._fail(error);
        } finally {
            this.pumping = false;
        }
    }

    _pumpOnce() {
        while (!this.error && !this.done && this.current < this.entries.length) {
            const entry = this.entries[this.current];
            if (!entry.headerWritten) {
                this._writeLocalHeader(entry);
            }
            const { pending } = entry;
            entry.pending = [];
            pending.forEach((chunk) => this._output(chunk));
            if (!entry.finished) {
                return;
            }
            if (entry.dataDescriptor) {
                this._writeDataDescriptor(entry);
            }
            this.current++;
            entry.emit("zipped");
        }
        if (this.finalizing && !this.error && !this.done && this.current === this.entries.length) {
            this._writeCentralDirectory();
            if (!this.error) {
                this.done = true;
                this.emit("end");
                this._resolve();
                this._checkDrained();
            }
        }
    }

    _needsZip64(...values) {
        return values.some((value) => value >= this.limits.size);
    }

    _writeLocalHeader(entry) {
        entry.offset = this.offset;
        entry.headerWritten = true;
        // the sizes are known unless the entry is still being streamed
        entry.dataDescriptor = !entry.finished;
        entry.zip64 = this.forceZip64 || (entry.finished && this._needsZip64(entry.size, entry.compressedSize));
        // with a data descriptor, the ZIP64 field only announces 64-bit sizes in the descriptor
        const extra = entry.zip64 ? zip64Extra(entry.finished ? [entry.size, entry.compressedSize] : [0, 0]) : EMPTY;
        const header = new Uint8Array(30 + entry.nameBytes.length + extra.length);
        writeU32(header, 0, SIG_LOCAL);
        writeU16(header, 4, entry.zip64 ? VERSION_ZIP64 : VERSION);
        writeU16(header, 6, entry.flags | (entry.dataDescriptor ? FLAG_DATA_DESCRIPTOR : 0));
        writeU16(header, 8, entry.method);
        writeU16(header, 10, this.dosDateTime.time);
        writeU16(header, 12, this.dosDateTime.date);
        if (!entry.dataDescriptor) {
            writeU32(header, 14, entry.crc);
        }
        if (entry.zip64) {
            writeU32(header, 18, MAX_U32);
            writeU32(header, 22, MAX_U32);
        } else if (!entry.dataDescriptor) {
            writeU32(header, 18, entry.compressedSize);
            writeU32(header, 22, entry.size);
        }
        writeU16(header, 26, entry.nameBytes.length);
        writeU16(header, 28, extra.length);
        header.set(entry.nameBytes, 30);
        header.set(extra, 30 + entry.nameBytes.length);
        this._output(header);
    }

    _writeDataDescriptor(entry) {
        // 64-bit sizes when the local header announced them, or when they are needed
        const wide = entry.zip64 || this._needsZip64(entry.size, entry.compressedSize);
        const descriptor = new Uint8Array(wide ? 24 : 16);
        writeU32(descriptor, 0, SIG_DATA_DESCRIPTOR);
        writeU32(descriptor, 4, entry.crc);
        if (wide) {
            writeU64(descriptor, 8, entry.compressedSize);
            writeU64(descriptor, 16, entry.size);
        } else {
            writeU32(descriptor, 8, entry.compressedSize);
            writeU32(descriptor, 12, entry.size);
        }
        this._output(descriptor);
    }

    _centralHeader(entry) {
        // the ZIP64 field holds the values that do not fit, in this order
        const values = [];
        const fit = (value) => {
            if (this.forceZip64 || this._needsZip64(value)) {
                values.push(value);
                return MAX_U32;
            }
            return value;
        };
        const size = fit(entry.size);
        const compressedSize = fit(entry.compressedSize);
        const offset = fit(entry.offset);
        const extra = values.length ? zip64Extra(values) : EMPTY;
        const version = values.length || entry.zip64 ? VERSION_ZIP64 : VERSION;
        const header = new Uint8Array(46 + entry.nameBytes.length + extra.length);
        writeU32(header, 0, SIG_CENTRAL);
        writeU16(header, 4, version);
        writeU16(header, 6, version);
        writeU16(header, 8, entry.flags | (entry.dataDescriptor ? FLAG_DATA_DESCRIPTOR : 0));
        writeU16(header, 10, entry.method);
        writeU16(header, 12, this.dosDateTime.time);
        writeU16(header, 14, this.dosDateTime.date);
        writeU32(header, 16, entry.crc);
        writeU32(header, 20, compressedSize);
        writeU32(header, 24, size);
        writeU16(header, 28, entry.nameBytes.length);
        writeU16(header, 30, extra.length);
        writeU32(header, 42, offset);
        header.set(entry.nameBytes, 46);
        header.set(extra, 46 + entry.nameBytes.length);
        return header;
    }

    _writeCentralDirectory() {
        const start = this.offset;
        this.entries.forEach((entry) => this._output(this._centralHeader(entry)));
        const size = this.offset - start;
        const count = this.entries.length;
        const manyEntries = count >= this.limits.entries;
        const zip64 = this.forceZip64 || manyEntries || this._needsZip64(size, start);

        if (zip64) {
            const recordOffset = this.offset;
            const record = new Uint8Array(56);
            writeU32(record, 0, SIG_ZIP64_EOCD);
            // size of the rest of the record
            writeU64(record, 4, 44);
            writeU16(record, 12, VERSION_ZIP64);
            writeU16(record, 14, VERSION_ZIP64);
            writeU64(record, 24, count);
            writeU64(record, 32, count);
            writeU64(record, 40, size);
            writeU64(record, 48, start);
            this._output(record);

            const locator = new Uint8Array(20);
            writeU32(locator, 0, SIG_ZIP64_LOCATOR);
            writeU64(locator, 8, recordOffset);
            // total number of disks
            writeU32(locator, 16, 1);
            this._output(locator);
        }

        const end = new Uint8Array(22 + this.comment.length);
        writeU32(end, 0, SIG_EOCD);
        writeU16(end, 8, manyEntries ? MAX_U16 : count);
        writeU16(end, 10, manyEntries ? MAX_U16 : count);
        writeU32(end, 12, this._needsZip64(size) ? MAX_U32 : size);
        writeU32(end, 16, this._needsZip64(start) ? MAX_U32 : start);
        writeU16(end, 20, this.comment.length);
        end.set(this.comment, 22);
        this._output(end);
    }
}

module.exports = {
    ZipWriter,
    normalizeZipOptions,
};
