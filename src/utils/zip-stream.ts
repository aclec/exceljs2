// Zip container writing. Compression and CRC-32 come from deflate.ts (node:zlib on Node.js and Bun,
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
import { EventEmitter } from "events";
import type { ZipOptions } from "../../index";
import {
    createCompressor,
    isCompressionLevel,
    selectBackend,
    type CompressionLevel,
    type Compressor,
    type CompressorCallback,
    type DeflateBackend,
} from "./deflate";

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

/** Thresholds of the ZIP64 records: values >= `size`, or `entries` entries or more, need them. */
export interface Zip64Limits {
    readonly size: number;
    readonly entries: number;
}

// only tests lower them (_zip64Limits)
const ZIP64_LIMITS: Zip64Limits = { size: MAX_U32, entries: MAX_U16 };

// uncompressed bytes being compressed (asynchronously) above which drained() waits
const HIGH_WATER_MARK = 4 * 1024 * 1024;

const textEncoder = new TextEncoder();
const EMPTY = new Uint8Array(0);

// -----------------------------------------------------------------------------
// options

/**
 * Options of {@link ZipWriter}: the public {@link ZipOptions} (including the deprecated shapes of the zip libraries used
 * before v5, archiver's `{zlib: {level}, store}` and JSZip's `{compression, compressionOptions}`), plus internal ones for
 * tests.
 */
export interface ZipWriterOptions extends Readonly<Partial<ZipOptions>> {
    /** internal, for tests: lower ZIP64 thresholds */
    readonly _zip64Limits?: Partial<Zip64Limits>;
    /** internal, for tests: the DEFLATE backend */
    readonly _backend?: DeflateBackend;
}

/** {@link ZipWriterOptions} validated, with their defaults */
export interface NormalizedZipOptions {
    readonly level: CompressionLevel;
    readonly forceZip64: boolean;
    /** the archive comment, UTF-8 encoded (at most 65535 bytes) */
    readonly comment: Uint8Array;
    readonly zip64Limits: Zip64Limits;
    readonly backend: DeflateBackend;
}

/**
 * Validates the zip options and applies their defaults (level 6, no comment, ZIP64 records only where needed).
 * `level` wins over the deprecated `zlib.level` and `compressionOptions.level`; `store: true` or
 * `compression: "STORE"` force level 0.
 *
 * @throws TypeError when the level is not an integer from 0 to 9, or the backend is unknown; RangeError when the comment
 * is longer than 65535 bytes (UTF-8)
 */
function normalizeZipOptions(options?: ZipWriterOptions | null): NormalizedZipOptions {
    const zip: ZipWriterOptions = options || {};
    let level: number = 6;
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
    if (!isCompressionLevel(level)) {
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
        zip64Limits: { ...ZIP64_LIMITS, ...zip._zip64Limits },
        // internal, for tests: "zlib" or "fflate"
        backend: selectBackend(zip._backend),
    };
}

// -----------------------------------------------------------------------------
// binary helpers

function writeU16(data: Uint8Array, pos: number, value: number): void {
    data[pos] = value & 0xff;
    data[pos + 1] = (value >>> 8) & 0xff;
}

function writeU32(data: Uint8Array, pos: number, value: number): void {
    data[pos] = value & 0xff;
    data[pos + 1] = (value >>> 8) & 0xff;
    data[pos + 2] = (value >>> 16) & 0xff;
    data[pos + 3] = (value >>> 24) & 0xff;
}

function writeU64(data: Uint8Array, pos: number, value: number): void {
    writeU32(data, pos, value % TWO_32);
    writeU32(data, pos + 4, Math.floor(value / TWO_32));
}

/** MS-DOS date and time of the entries (local time, 2-second resolution, years 1980 to 2107) */
interface DosDateTime {
    readonly time: number;
    readonly date: number;
}

function dosDateTime(date: Date): DosDateTime {
    const year = Math.min(Math.max(date.getFullYear(), 1980), 2107);
    return {
        time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
        date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    };
}

// ZIP64 extended information extra field
function zip64Extra(values: readonly number[]): Uint8Array {
    const extra = new Uint8Array(4 + 8 * values.length);
    writeU16(extra, 0, 1);
    writeU16(extra, 2, 8 * values.length);
    let pos = 4;
    for (const value of values) {
        writeU64(extra, pos, value);
        pos += 8;
    }
    return extra;
}

/**
 * Data of a zip entry: a string (UTF-8 encoded), bytes (a Buffer, any typed array or DataView, an ArrayBuffer), or an
 * object with a `toBuffer()` method (StringBuf, src/utils/string-buf.ts).
 */
export type ZipEntryData = string | ArrayBufferView | ArrayBuffer | { toBuffer(): Uint8Array };

/** @throws TypeError on any other value (JavaScript callers) */
function toBytes(data: ZipEntryData): Uint8Array {
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
    // StringBuf (src/utils/string-buf.ts)
    if (data && typeof data.toBuffer === "function") {
        return data.toBuffer();
    }
    throw new TypeError("Zip entry data must be a string, a Buffer, a Uint8Array or an ArrayBuffer");
}

// a Buffer view of the bytes (not a copy)
function toBuffer(data: Uint8Array): Buffer {
    return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
}

// leading slashes are dropped (like archiver did); anything that could extract outside of the target
// directory is refused (zip slip)
function entryName(name: string): string {
    const normalised = String(name).replace(/^\/+/, "");
    if (!normalised || normalised.includes("\\") || normalised.split("/").some((part) => part === ".." || part === ".")) {
        throw new Error(`Invalid zip entry name: ${JSON.stringify(String(name))}`);
    }
    return normalised;
}

// -----------------------------------------------------------------------------

/** events of a {@link ZipEntry}: 'finish' once ended, 'zipped' once all its data is in the archive output */
type ZipEntryEvents = {
    finish: [];
    zipped: [];
};

/**
 * One entry of a {@link ZipWriter}. It is also the writable that the streaming writers write a part to:
 * `write(data)`... `end()`. Emits 'finish' when ended and 'zipped' once all its data is in the archive output.
 */
class ZipEntry extends EventEmitter<ZipEntryEvents> {
    /** the name in the archive (leading slashes dropped) */
    readonly name: string;
    /** The name encoded in UTF-8 (as written in the headers). */
    readonly nameBytes: Uint8Array;
    /** general purpose flags, but bit 3 (data descriptor) */
    readonly flags: number;
    /** Stored without compression (level 0). */
    readonly stored: boolean;
    /** The compressor of the entry data (see deflate.ts). */
    readonly compressor: Compressor;
    /** compressed bytes produced so far */
    compressedSize = 0;
    /** compressed data waiting for the earlier entries */
    pending: Uint8Array[] = [];
    /** `end()` was called: no more data can be written. */
    ended = false;
    /** all the data is compressed: `crc` and `size` are known */
    finished = false;
    /** The local file header was written to the output. */
    headerWritten = false;
    /** CRC-32 of the uncompressed data, once finished */
    crc = 0;
    /** uncompressed size, once finished */
    size = 0;
    /** offset of the local header in the archive, once written */
    offset = 0;
    /** the sizes follow the data, in a data descriptor (the entry was not finished when its header was written) */
    dataDescriptor = false;
    /** the local header has a ZIP64 extra field */
    zip64 = false;

    /**
     * @throws Error when the name is invalid (empty, a `.` or `..` segment, a backslash) or longer than 65535 bytes
     */
    constructor(
        private readonly zip: ZipWriter,
        name: string
    ) {
        super();
        this.name = entryName(name);
        this.nameBytes = textEncoder.encode(this.name);
        if (this.nameBytes.length > MAX_U16) {
            throw new Error(`Zip entry name is too long: ${this.name}`);
        }
        this.flags = this.nameBytes.length === this.name.length ? 0 : FLAG_UTF8;
        this.stored = zip.level === 0;
        this.compressor = createCompressor(zip.level, zip.backend, (...args) => this._onData(...args));
    }

    /** compression method: 0 (stored) or 8 (deflated) */
    get method(): 0 | 8 {
        return this.stored ? 0 : 8;
    }

    /**
     * Adds data to the entry. The data is compressed at once (or copied): the caller may reuse its buffer.
     *
     * @returns true (backpressure goes through {@link ZipWriter.drained})
     * @throws Error after end(); TypeError on an invalid data type
     */
    write(data: ZipEntryData): boolean {
        if (this.ended) {
            throw new Error(`Write after end of the zip entry ${this.name}`);
        }
        const bytes = toBytes(data);
        if (bytes.length) {
            this._compress(bytes, false);
        }
        return true;
    }

    /**
     * Ends the entry, with optional last data. Further calls are ignored.
     *
     * @throws TypeError on an invalid data type
     */
    end(data?: ZipEntryData): void {
        if (this.ended) {
            return;
        }
        this.ended = true;
        this._compress(data === undefined ? EMPTY : toBytes(data), true);
        this.emit("finish");
    }

    private _compress(bytes: Uint8Array, final: boolean): void {
        try {
            this.compressor.push(bytes, final);
        } catch (error) {
            this.zip._fail(error);
        }
    }

    private _onData(...[error, chunks, final]: Parameters<CompressorCallback>): void {
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

/** options of {@link ZipWriter.append} */
export interface ZipAppendOptions {
    /** the entry name (leading slashes are dropped) */
    readonly name: string;
    /** the data is a base64 string, to decode */
    readonly base64?: boolean;
}

/**
 * Where the archive is written: a Node.js Writable, a StreamBuf, or any object with `write()` and `end()`. When it has
 * `on()` (a stream, which then also has `removeListener()`), its 'drain' event drives the backpressure and its
 * 'error' and 'close' events fail an incomplete archive.
 */
export interface ZipOutput {
    /** @returns false to ask to wait for 'drain' (streams) */
    write(chunk: Buffer): unknown;
    end(): unknown;
    on?(event: "error" | "drain" | "close", listener: (error: Error) => void): unknown;
    removeListener?(event: "error" | "drain" | "close", listener: (error: Error) => void): unknown;
}

type ZipOutputStream = ZipOutput & Required<Pick<ZipOutput, "on" | "removeListener">>;

// streams have both on() and removeListener()
function isStream(output: ZipOutput): output is ZipOutputStream {
    return typeof output.on === "function";
}

/** events of a {@link ZipWriter}: the archive output, then its end */
type ZipWriterEvents = {
    data: [chunk: Buffer];
    end: [];
};

/**
 * Streaming zip writer. Entries are added with {@link append} (all of the data at once) or {@link openStream} (data
 * written later), and written to the output (see {@link pipe}) in the order they are added; {@link finalize} writes the
 * central directory once every entry is complete.
 */
class ZipWriter extends EventEmitter<ZipWriterEvents> {
    /** The compression level (0: stored, 1–9: deflate). */
    readonly level: CompressionLevel;
    /** The deflate implementation: node:zlib when present, else fflate (browser bundle). */
    readonly backend: DeflateBackend;
    /** Write ZIP64 records even when the sizes do not need them. */
    readonly forceZip64: boolean;
    /** The archive comment, encoded in UTF-8. */
    readonly comment: Uint8Array;
    /** The sizes from which ZIP64 records are needed (lowered by tests). */
    readonly limits: Zip64Limits;
    /** The modification date and time written for every entry (MS-DOS format). */
    readonly dosDateTime: DosDateTime;
    /** The entries, in the order they were added (and are written). */
    readonly entries: ZipEntry[] = [];
    /** index of the entry being written to the output */
    current = 0;
    /** bytes written to the output */
    offset = 0;
    /** `finalize()` was called: no entry can be added. */
    finalizing = false;
    /** the archive is complete */
    done = false;
    /** the first failure; the archive then stays incomplete */
    private failure: { readonly error: unknown } | null = null;
    /** the output stream asked to wait for 'drain' */
    private outputFull = false;
    /** the promise drained() returns while the output is backlogged, and its resolve / reject functions */
    private drainPromise: Promise<void> | null = null;
    private drainWaiter: { readonly resolve: () => void; readonly reject: (error: unknown) => void } | null = null;
    /** settles when the archive is complete, or as soon as it fails */
    readonly finished: Promise<void>;
    // set by the executor of `finished`, which runs synchronously in the constructor
    private _resolve!: () => void;
    private _reject!: (error: unknown) => void;
    // _pump() is running, and has to run once more
    private pumping = false;
    private pumpAgain = false;

    /** @throws see {@link normalizeZipOptions} */
    constructor(options?: ZipWriterOptions | null) {
        super();
        const { level, forceZip64, comment, zip64Limits, backend } = normalizeZipOptions(options);
        this.level = level;
        this.backend = backend;
        this.forceZip64 = forceZip64;
        this.comment = comment;
        this.limits = zip64Limits;
        this.dosDateTime = dosDateTime(new Date());
        this.finished = new Promise((resolve, reject) => {
            this._resolve = resolve;
            this._reject = reject;
        });
        // a failure is reported by finalize() (and the methods called after it), or through `finished`
        this.finished.catch(() => {});
    }

    /** the first failure of the archive, or null */
    get error(): unknown {
        return this.failure ? this.failure.error : null;
    }

    /**
     * Writes the archive output to `destination` ('data' events, then `end()`).
     *
     * @returns `destination`
     */
    pipe<T extends ZipOutput>(destination: T): T {
        // backpressure needs the 'drain' event
        const drains = isStream(destination);
        this.on("data", (chunk) => {
            if (destination.write(chunk) === false && drains) {
                this.outputFull = true;
            }
        });
        if (isStream(destination)) {
            // an error of the destination while the archive is written fails it; once it is complete, the
            // destination's errors are its owner's again (the listener would hide them)
            const onError = (error: Error): void => this._fail(error);
            const onDrain = (): void => {
                this.outputFull = false;
                this._checkDrained();
            };
            // destroyed without an error (e.g. an HTTP response whose client went away): 'drain' will never come
            const onClose = (): void => this._fail(new Error("The output stream was closed before the zip archive was complete"));
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

    /**
     * Adds an entry with all of its data, compressed at once.
     *
     * @param data - the entry data; with `options.base64`, a base64 string to decode
     * @throws the error of a failed archive; Error after finalize() or on an invalid name; TypeError on an invalid data
     * type
     */
    append(data: ZipEntryData, options: ZipAppendOptions): void {
        this._check();
        // converted first: an invalid value must not leave an entry that never ends
        const bytes = options.base64 && typeof data === "string" ? Buffer.from(data, "base64") : toBytes(data);
        const entry = new ZipEntry(this, options.name);
        this.entries.push(entry);
        entry.end(bytes);
        this._pump();
    }

    /**
     * Adds an entry whose data is written later, through the returned entry (`write()`... `end()`).
     *
     * @throws the error of a failed archive; Error after finalize() or on an invalid name
     */
    openStream(name: string): ZipEntry {
        this._check();
        const entry = new ZipEntry(this, name);
        this.entries.push(entry);
        this._pump();
        return entry;
    }

    /**
     * Writes the central directory once every entry is complete.
     *
     * @returns {@link finished}: resolves when the output has ended, rejects with the error of a failed archive
     */
    finalize(): Promise<void> {
        if (!this.finalizing && !this.failure) {
            this.finalizing = true;
            this._pump();
        }
        return this.finished;
    }

    /**
     * A promise that resolves once the output can take more data (see Backpressure above), or rejects with the error
     * of a failed archive. It is marked handled: callers may ignore it.
     */
    drained(): Promise<void> {
        if (this.failure) {
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

    private _backlogged(): boolean {
        if (this.done) {
            return false;
        }
        if (this.outputFull) {
            return true;
        }
        let inFlight = 0;
        for (let i = this.current; i < this.entries.length; i++) {
            inFlight += this.entries[i]?.compressor.inFlight ?? 0;
        }
        return inFlight > HIGH_WATER_MARK;
    }

    /** settles the promise of drained() once the backlog clears or the archive fails (@internal, for ZipEntry) */
    _checkDrained(): void {
        const waiter = this.drainWaiter;
        if (waiter && (this.failure || !this._backlogged())) {
            this.drainPromise = null;
            this.drainWaiter = null;
            if (this.failure) {
                waiter.reject(this.failure.error);
            } else {
                waiter.resolve();
            }
        }
    }

    private _check(): void {
        if (this.failure) {
            throw this.failure.error;
        }
        if (this.finalizing) {
            throw new Error("Cannot add zip entries after finalize()");
        }
    }

    /** fails the archive, unless it is complete or already failed (@internal, for ZipEntry) */
    _fail(error: unknown): void {
        if (!this.failure && !this.done) {
            this.failure = { error };
            this._reject(error);
            this._checkDrained();
        }
    }

    /** writes bytes to the output (@internal, for ZipEntry) */
    _output(data: Uint8Array): void {
        // nothing more for a failed destination
        if (this.failure) {
            return;
        }
        this.offset += data.length;
        try {
            this.emit("data", toBuffer(data));
        } catch (error) {
            this._fail(error);
        }
    }

    /**
     * Writes what can be written; re-entrant calls (e.g. from a 'zipped' listener) are folded into the running one
     * (@internal, for ZipEntry).
     */
    _pump(): void {
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

    private _pumpOnce(): void {
        let entry: ZipEntry | undefined;
        while (!this.failure && !this.done && (entry = this.entries[this.current])) {
            if (!entry.headerWritten) {
                this._writeLocalHeader(entry);
            }
            const { pending } = entry;
            entry.pending = [];
            for (const chunk of pending) {
                this._output(chunk);
            }
            if (!entry.finished) {
                return;
            }
            if (entry.dataDescriptor) {
                this._writeDataDescriptor(entry);
            }
            this.current++;
            entry.emit("zipped");
        }
        if (this.finalizing && !this.failure && !this.done && this.current === this.entries.length) {
            this._writeCentralDirectory();
            if (!this.failure) {
                this.done = true;
                this.emit("end");
                this._resolve();
                this._checkDrained();
            }
        }
    }

    private _needsZip64(value: number): boolean {
        return value >= this.limits.size;
    }

    private _writeLocalHeader(entry: ZipEntry): void {
        entry.offset = this.offset;
        entry.headerWritten = true;
        // the sizes are known unless the entry is still being streamed
        entry.dataDescriptor = !entry.finished;
        entry.zip64 = this.forceZip64 || (entry.finished && (this._needsZip64(entry.size) || this._needsZip64(entry.compressedSize)));
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

    private _writeDataDescriptor(entry: ZipEntry): void {
        // 64-bit sizes when the local header announced them, or when they are needed
        const wide = entry.zip64 || this._needsZip64(entry.size) || this._needsZip64(entry.compressedSize);
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

    private _centralHeader(entry: ZipEntry): Uint8Array {
        // the ZIP64 field holds the values that do not fit, in this order
        const values: number[] = [];
        const fit = (value: number): number => {
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

    private _writeCentralDirectory(): void {
        const start = this.offset;
        for (const entry of this.entries) {
            this._output(this._centralHeader(entry));
        }
        const size = this.offset - start;
        const count = this.entries.length;
        const manyEntries = count >= this.limits.entries;
        const zip64 = this.forceZip64 || manyEntries || this._needsZip64(size) || this._needsZip64(start);

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

        // with the ZIP64 records, the fields hold their "see the ZIP64 record" value (APPNOTE 4.4.1.4, as archiver
        // wrote): readers such as JSZip only look for the ZIP64 records when they do
        const end = new Uint8Array(22 + this.comment.length);
        writeU32(end, 0, SIG_EOCD);
        writeU16(end, 8, zip64 ? MAX_U16 : count);
        writeU16(end, 10, zip64 ? MAX_U16 : count);
        writeU32(end, 12, zip64 ? MAX_U32 : size);
        writeU32(end, 16, zip64 ? MAX_U32 : start);
        writeU16(end, 20, this.comment.length);
        end.set(this.comment, 22);
        this._output(end);
    }
}

export { ZipWriter, normalizeZipOptions };
export type { ZipEntry };
