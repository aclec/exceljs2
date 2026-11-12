// Zip container reading. Decompression comes from deflate.ts (node:zlib on Node.js and Bun, fflate in
// the browser).
//
// - readZip(buffer, options): in-memory archive (xlsx.load). The central directory is the source of
//   truth (like JSZip): entries keep the central directory order, ZIP64 records are understood and
//   data that was prepended to the archive is tolerated. Entries are inflated one at a time, on demand.
// - parseZipStream(chunks, options): forward-only archive (stream.xlsx.WorkbookReader) read from an
//   async iterable of byte chunks, local header by local header, with no random access.
//
// Both enforce the decompression limits (see ZipLimits) while inflating, so that a zip bomb is
// rejected before its output is materialised.
import type { ZipReadLimits } from "../../index";
import { createInflater, inflateWhole, selectBackend, sliceSize, type DeflateBackend } from "./deflate";

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_DATA_DESCRIPTOR = 0x08074b50;
const SIG_SPANNING = 0x30304b50;

const MAX_U16 = 0xffff;
const MAX_U32 = 0xffffffff;
const TWO_32 = 0x100000000;

const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;

const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

const KiB = 1024;
const MiB = 1024 * KiB;
const GiB = 1024 * MiB;

/**
 * Options of {@link readZip} and {@link parseZipStream}: the decompression limits, each one defaulting to the reader's
 * ({@link DOCUMENT_LIMITS} or {@link STREAM_LIMITS}). Other properties are ignored: the reader options can be passed as
 * they are.
 */
export interface ZipReadOptions extends Readonly<Partial<ZipReadLimits>> {
    /** internal, for tests: the DEFLATE backend */
    readonly _backend?: DeflateBackend;
}

/**
 * Default limits of the document reader (xlsx.load / read / readFile), which materialises every part: a single part
 * cannot be parsed past V8's maximum string length (~512 MiB) anyway, and the whole model has to fit in memory.
 */
const DOCUMENT_LIMITS: Readonly<ZipReadLimits> = {
    maxEntrySize: 1 * GiB,
    maxUncompressedSize: 2 * GiB,
    maxCompressionRatio: 1000,
};
/**
 * Default limits of the streaming reader, which parses worksheets without materialising them (multi-GiB sheets are its
 * purpose); worksheets that arrive before the parts they depend on are spooled to temp files, which these sizes also
 * bound.
 */
const STREAM_LIMITS: Readonly<ZipReadLimits> = {
    maxEntrySize: 16 * GiB,
    maxUncompressedSize: 32 * GiB,
    maxCompressionRatio: 1000,
};
// the compression ratio of a part, and of all the parts together, is only checked past this size: tiny parts can have any
// ratio
const RATIO_MIN_SIZE = 10 * MiB;

const textDecoder = new TextDecoder("utf-8");

function zipError(message: string): Error {
    return new Error(`Invalid or unsupported zip file: ${message}`);
}

// Little-endian unsigned integers. A byte past the end of `data` reads as 0 (as the bitwise operators of the original
// JavaScript did with `undefined`): callers check the bounds where it matters.
function u16(data: Uint8Array, pos: number): number {
    return (data[pos] ?? 0) | ((data[pos + 1] ?? 0) << 8);
}

function u32(data: Uint8Array, pos: number): number {
    return ((data[pos] ?? 0) | ((data[pos + 1] ?? 0) << 8) | ((data[pos + 2] ?? 0) << 16) | ((data[pos + 3] ?? 0) << 24)) >>> 0;
}

// sizes and offsets must stay exact: values past Number.MAX_SAFE_INTEGER (2^53 - 1) are refused
function u64(data: Uint8Array, pos: number): number {
    const value = u32(data, pos + 4) * TWO_32 + u32(data, pos);
    if (value > Number.MAX_SAFE_INTEGER) {
        throw zipError("64-bit size or offset out of range");
    }
    return value;
}

function decodeName(bytes: Uint8Array): string {
    // like JSZip: names are decoded as UTF-8 whether or not the UTF-8 flag (bit 11) is set
    return textDecoder.decode(bytes);
}

/** bounds of some data in a buffer: [start, end) */
interface Span {
    readonly start: number;
    readonly end: number;
}

/** a size or an offset, whose 0xFFFFFFFF placeholder may be replaced by its ZIP64 value */
interface Zip64Field {
    value: number;
}

// the value of the ZIP64 extended information extra field (header ID 1), or null
function findZip64Extra(data: Uint8Array, start: number, length: number): Span | null {
    const end = start + length;
    let pos = start;
    while (pos + 4 <= end) {
        const id = u16(data, pos);
        const size = u16(data, pos + 2);
        if (pos + 4 + size > end) {
            break;
        }
        if (id === 1) {
            return { start: pos + 4, end: pos + 4 + size };
        }
        pos += 4 + size;
    }
    return null;
}

// replace the 0xFFFFFFFF placeholders by their ZIP64 values, which follow a fixed order
function applyZip64(data: Uint8Array, extraStart: number, extraLength: number, fields: readonly Zip64Field[], name: string): void {
    const wanted = fields.filter(({ value }) => value === MAX_U32);
    if (!wanted.length) {
        return;
    }
    const extra = findZip64Extra(data, extraStart, extraLength);
    if (!extra || extra.end - extra.start < wanted.length * 8) {
        throw zipError(`missing ZIP64 extra field for "${name}"`);
    }
    let pos = extra.start;
    for (const field of wanted) {
        field.value = u64(data, pos);
        pos += 8;
    }
}

// =============================================================================
// Decompression limits

// the option, else the default; options also come from JavaScript callers, whatever their declared type
function limitValue(options: ZipReadOptions | undefined, name: keyof ZipReadLimits, defaults: Readonly<ZipReadLimits>): number {
    const value: unknown = options && options[name] !== undefined ? options[name] : defaults[name];
    if (!(typeof value === "number" && value > 0)) {
        throw new TypeError(`${name} must be a positive number, got ${String(value)}`);
    }
    return value;
}

/**
 * The decompression limits of one archive, and the sizes counted against them. A check that fails throws an Error that
 * names the limit exceeded.
 */
class ZipLimits implements Readonly<ZipReadLimits> {
    readonly maxEntrySize: number;
    readonly maxUncompressedSize: number;
    readonly maxCompressionRatio: number;
    /** bytes inflated so far */
    private total = 0;
    /** sum of the declared uncompressed sizes */
    private declaredTotal = 0;
    /** compressed bytes of the entries inflated (`total` comes from them) */
    private compressedTotal = 0;

    /** @throws TypeError when a limit is not a positive number */
    constructor(options: ZipReadOptions | undefined, defaults: Readonly<ZipReadLimits>) {
        this.maxEntrySize = limitValue(options, "maxEntrySize", defaults);
        this.maxUncompressedSize = limitValue(options, "maxUncompressedSize", defaults);
        this.maxCompressionRatio = limitValue(options, "maxCompressionRatio", defaults);
    }

    // declared sizes (central directory or local header): fail before any work
    checkDeclared(name: string, compressedSize: number, uncompressedSize: number): void {
        this.declaredTotal += uncompressedSize;
        this.check(name, uncompressedSize, compressedSize, Math.max(this.total, this.declaredTotal));
    }

    // declared sizes of one entry
    checkEntry(name: string, compressedSize: number, uncompressedSize: number): void {
        this.check(name, uncompressedSize, compressedSize, this.total);
    }

    // compressed data about to be inflated
    addCompressed(bytes: number): void {
        this.compressedTotal += bytes;
    }

    // actual sizes, while inflating
    add(name: string, bytes: number, entrySize: number, compressedSize: number): void {
        this.total += bytes;
        this.check(name, entrySize, compressedSize, this.total);
        // small parts too: the content of the archive is bounded by its size (many parts that share their data, or many
        // parts that each compress at the maximum ratio)
        if (this.total > RATIO_MIN_SIZE && this.total > this.compressedTotal * this.maxCompressionRatio) {
            throw new Error(`Zip content exceeds maxCompressionRatio (${this.maxCompressionRatio})`);
        }
    }

    private check(name: string, entrySize: number, compressedSize: number, total: number): void {
        if (entrySize > this.maxEntrySize) {
            throw new Error(`Zip entry "${name}" is larger than maxEntrySize (${this.maxEntrySize} bytes)`);
        }
        if (total > this.maxUncompressedSize) {
            throw new Error(`Zip content is larger than maxUncompressedSize (${this.maxUncompressedSize} bytes)`);
        }
        if (entrySize > RATIO_MIN_SIZE && entrySize > compressedSize * this.maxCompressionRatio) {
            throw new Error(`Zip entry "${name}" exceeds maxCompressionRatio (${this.maxCompressionRatio})`);
        }
    }
}

function checkMethod(name: string, method: number, flags: number): void {
    if (flags & FLAG_ENCRYPTED) {
        throw new Error(`Encrypted zip entries are not supported: "${name}"`);
    }
    if (method !== METHOD_STORE && method !== METHOD_DEFLATE) {
        throw zipError(`unknown compression method ${method} for "${name}"`);
    }
}

// =============================================================================
// In-memory archive

function findEndOfCentralDirectory(data: Uint8Array): number {
    // the record is 22 bytes followed by a comment of up to 65535 bytes
    const min = Math.max(0, data.length - 22 - MAX_U16);
    for (let pos = data.length - 22; pos >= min; pos--) {
        if (u32(data, pos) === SIG_EOCD) {
            return pos;
        }
    }
    throw zipError("end of central directory not found (is this an xlsx file?)");
}

/** where the central directory is in the buffer */
interface CentralDirectoryLocation {
    /** number of entries recorded in the end records */
    readonly count: number;
    readonly start: number;
    /** where the central directory ends (it is followed by the end records) */
    readonly end: number;
    /** number of bytes prepended to the archive */
    readonly shift: number;
}

function readCentralDirectoryLocation(data: Uint8Array): CentralDirectoryLocation {
    const eocd = findEndOfCentralDirectory(data);
    let count = u16(data, eocd + 10);
    let size = u32(data, eocd + 12);
    let offset = u32(data, eocd + 16);
    // where the central directory ends in this buffer (it is followed by the end records)
    let end = eocd;

    const locator = eocd - 20;
    if (locator >= 0 && u32(data, locator) === SIG_ZIP64_LOCATOR) {
        const recordOffset = u64(data, locator + 8);
        // like the central directory, the record may be shifted by data prepended to the archive
        const record = u32(data, recordOffset) === SIG_ZIP64_EOCD ? recordOffset : locator - 56;
        if (record < 0 || u32(data, record) !== SIG_ZIP64_EOCD) {
            throw zipError("ZIP64 end of central directory record not found");
        }
        count = u64(data, record + 32);
        size = u64(data, record + 40);
        offset = u64(data, record + 48);
        end = record;
    } else if (count === MAX_U16 || size === MAX_U32 || offset === MAX_U32) {
        throw zipError("ZIP64 end of central directory locator not found");
    }

    // bytes prepended to the archive (e.g. a self-extractor stub) shift every offset
    const shift = end - (offset + size);
    if (shift < 0) {
        throw zipError("central directory out of bounds");
    }
    return { count, start: offset + shift, end, shift };
}

/** an entry of the central directory */
interface CentralEntry {
    readonly name: string;
    readonly dir: boolean;
    readonly flags: number;
    readonly method: number;
    readonly compressedSize: number;
    readonly uncompressedSize: number;
    /** offset of the local header in the buffer (prepended data included) */
    readonly localOffset: number;
}

// Each entry must have its own data: entries that share their data (the central directory points several times at
// the same local header, or at a local header within the data of another entry) multiply the inflated output without
// adding compressed data (overlapping zip bomb). No zip writer produces them.
function checkOverlaps(data: Uint8Array, entries: readonly CentralEntry[]): void {
    const ranges = entries.map(({ name, localOffset, compressedSize }) => {
        // the lengths of the local header, or a lower bound when it is invalid (then reading the entry fails)
        const valid = localOffset + 30 <= data.length && u32(data, localOffset) === SIG_LOCAL;
        const header = valid ? 30 + u16(data, localOffset + 26) + u16(data, localOffset + 28) : 30;
        return { name, start: localOffset, end: localOffset + header + compressedSize };
    });
    ranges.sort((a, b) => a.start - b.start);
    let previous = ranges[0];
    for (let i = 1; i < ranges.length; i++) {
        const range = ranges[i];
        if (previous && range && range.start < previous.end) {
            throw zipError(`entries "${previous.name}" and "${range.name}" overlap`);
        }
        previous = range;
    }
}

function readCentralDirectory(data: Uint8Array, limits: ZipLimits): CentralEntry[] {
    const { count, start, end, shift } = readCentralDirectoryLocation(data);
    // every entry of the central directory, and the entry kept for each name
    const all: CentralEntry[] = [];
    const entries = new Map<string, CentralEntry>();
    let pos = start;
    let found = 0;
    while (pos + 46 <= end && u32(data, pos) === SIG_CENTRAL) {
        const flags = u16(data, pos + 8);
        const method = u16(data, pos + 10);
        const nameLength = u16(data, pos + 28);
        const extraLength = u16(data, pos + 30);
        const commentLength = u16(data, pos + 32);
        const nameStart = pos + 46;
        const extraStart = nameStart + nameLength;
        if (extraStart + extraLength + commentLength > end) {
            throw zipError("central directory entry out of bounds");
        }
        const name = decodeName(data.subarray(nameStart, extraStart));
        const uncompressedSize: Zip64Field = { value: u32(data, pos + 24) };
        const compressedSize: Zip64Field = { value: u32(data, pos + 20) };
        const localOffset: Zip64Field = { value: u32(data, pos + 42) };
        applyZip64(data, extraStart, extraLength, [uncompressedSize, compressedSize, localOffset], name);

        limits.checkDeclared(name, compressedSize.value, uncompressedSize.value);
        const entry: CentralEntry = {
            name,
            dir: name.endsWith("/"),
            flags,
            method,
            compressedSize: compressedSize.value,
            uncompressedSize: uncompressedSize.value,
            localOffset: localOffset.value + shift,
        };
        all.push(entry);
        // a later entry with the same name replaces the earlier one but keeps its position (JSZip)
        entries.set(name, entry);
        pos = extraStart + extraLength + commentLength;
        found++;
    }
    if (found !== count && count !== MAX_U16) {
        throw zipError(`expected ${count} central directory entries, found ${found}`);
    }
    checkOverlaps(data, all);
    return [...entries.values()];
}

// the compressed data of an entry: a view of the archive buffer
function entryData(data: Uint8Array, entry: CentralEntry): Uint8Array {
    const { localOffset, compressedSize, name } = entry;
    if (localOffset + 30 > data.length || u32(data, localOffset) !== SIG_LOCAL) {
        throw zipError(`local header not found for "${name}"`);
    }
    const start = localOffset + 30 + u16(data, localOffset + 26) + u16(data, localOffset + 28);
    if (start + compressedSize > data.length) {
        throw zipError(`data of "${name}" out of bounds`);
    }
    return data.subarray(start, start + compressedSize);
}

/** An in-memory archive: a Buffer, any typed array or DataView, an ArrayBuffer, or an array of byte values. */
export type ZipInput = ArrayBufferView | ArrayBuffer | number[];

/** An entry of an in-memory archive ({@link readZip}). */
export interface ZipFileEntry {
    /** path in the archive, as recorded (decoded as UTF-8) */
    readonly name: string;
    /** true when the name ends with "/" */
    readonly dir: boolean;
    /**
     * Inflates the entry, checking the decompression limits (of the entry and of the whole archive).
     *
     * @returns a new Uint8Array, never a view of the archive buffer
     * @throws (rejects) when a limit is exceeded, the entry is encrypted or uses an unknown compression method, or its
     * data is invalid or does not have the declared size
     */
    data(): Promise<Uint8Array>;
}

/**
 * Reads an in-memory zip archive. The central directory is read at once (the declared sizes are checked against the
 * limits); the entries are inflated on demand, one at a time.
 *
 * @param buffer - the archive; data prepended to it (e.g. a self-extractor stub) is tolerated
 * @param options - decompression limits, defaulting to {@link DOCUMENT_LIMITS}
 * @returns the entries in central directory order; a later entry with the same name replaces an earlier one but keeps
 * its position (like JSZip)
 * @throws TypeError on an invalid input or limit; Error on an invalid or unsupported archive, declared sizes over the
 * limits, or entries that share their data
 */
function readZip(buffer: ZipInput, options?: ZipReadOptions): ZipFileEntry[] {
    // checked for JavaScript callers
    if (!(ArrayBuffer.isView(buffer) || buffer instanceof ArrayBuffer || Array.isArray(buffer))) {
        throw new TypeError("Can't read the zip file: expected a Buffer, a Uint8Array, an ArrayBuffer or a Blob");
    }
    const data = Array.isArray(buffer) ? new Uint8Array(buffer) : toBytes(buffer);
    const limits = new ZipLimits(options, DOCUMENT_LIMITS);
    const backend = selectBackend(options && options._backend);
    return readCentralDirectory(data, limits).map((entry): ZipFileEntry => ({
        name: entry.name,
        dir: entry.dir,
        // synchronous: what inflateEntry throws rejects the promise
        data: () => new Promise((resolve) => resolve(inflateEntry(data, entry, limits, backend))),
    }));
}

// inflates an entry of an in-memory archive into a new array
function inflateEntry(data: Uint8Array, entry: CentralEntry, limits: ZipLimits, backend: DeflateBackend): Uint8Array {
    const { name, method, flags, uncompressedSize } = entry;
    checkMethod(name, method, flags);
    const compressed = entryData(data, entry);
    limits.addCompressed(compressed.length);
    // the declared size was checked against the limits: inflating stops past it
    let output: Uint8Array | null;
    if (method === METHOD_STORE) {
        // a copy: not a view of the caller's buffer
        output = compressed.length > uncompressedSize ? null : new Uint8Array(compressed);
    } else {
        output = inflateWhole(compressed, uncompressedSize, backend);
    }
    if (!output || output.length > uncompressedSize) {
        throw zipError(`"${name}" is larger than its declared size`);
    }
    if (output.length !== uncompressedSize) {
        throw zipError(`"${name}" is smaller than its declared size`);
    }
    limits.add(name, output.length, output.length, compressed.length);
    return output;
}

// =============================================================================
// Forward-only archive

// a view of the bytes of an input chunk (not a copy)
function toBytes(value: unknown): Uint8Array {
    if (value instanceof Uint8Array) {
        return value;
    }
    if (ArrayBuffer.isView(value)) {
        return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
    }
    if (value instanceof ArrayBuffer) {
        return new Uint8Array(value);
    }
    throw new TypeError(`Zip input chunks must be bytes, got ${typeof value}`);
}

/** the input of the forward-only reader, buffered: `buffer[pos...]` is read and not consumed yet */
class ByteSource {
    private readonly iterator: AsyncIterator<unknown, unknown, undefined>;
    buffer: Uint8Array = new Uint8Array(0);
    pos = 0;
    /** true at the end of the input (or once closed) */
    ended = false;

    constructor(chunks: AsyncIterable<unknown, unknown, undefined>) {
        this.iterator = chunks[Symbol.asyncIterator]();
    }

    /** bytes read and not consumed */
    get available(): number {
        return this.buffer.length - this.pos;
    }

    // read one more chunk; false at the end of the input
    async more(): Promise<boolean> {
        if (this.ended) {
            return false;
        }
        const { value, done } = await this.iterator.next();
        if (done) {
            this.ended = true;
            return false;
        }
        const chunk = toBytes(value);
        if (this.available === 0) {
            this.buffer = chunk;
        } else {
            const merged = new Uint8Array(this.available + chunk.length);
            merged.set(this.buffer.subarray(this.pos));
            merged.set(chunk, this.available);
            this.buffer = merged;
        }
        this.pos = 0;
        return true;
    }

    // at least `length` bytes available; false if the input ends first
    async ensure(length: number): Promise<boolean> {
        while (this.available < length) {
            if (!(await this.more())) {
                return false;
            }
        }
        return true;
    }

    /** @throws Error when the input ends before `length` bytes are available */
    async require(length: number): Promise<void> {
        if (!(await this.ensure(length))) {
            throw new Error("Unexpected end of zip data");
        }
    }

    // consumes `length` available bytes: a view of the input
    take(length: number): Uint8Array {
        const result = this.buffer.subarray(this.pos, this.pos + length);
        this.pos += length;
        return result;
    }

    u32(offset: number): number {
        return u32(this.buffer, this.pos + offset);
    }

    // stop reading: releases the input (e.g. destroys a stream)
    async close(): Promise<void> {
        if (!this.ended && typeof this.iterator.return === "function") {
            this.ended = true;
            await this.iterator.return();
        }
    }
}

// the records that may follow the data of an entry
const NEXT_RECORD: readonly number[] = [SIG_LOCAL, SIG_CENTRAL, SIG_ZIP64_EOCD, SIG_EOCD];

/** a local header */
interface LocalHeader {
    readonly name: string;
    readonly flags: number;
    readonly method: number;
    readonly compressedSize: number;
    /** with a data descriptor, the size it records, set once the data is read */
    uncompressedSize: number;
    /** the header has a ZIP64 extra field (the data descriptor then has 64-bit sizes) */
    readonly zip64: boolean;
    /** general purpose flag bit 3: the sizes are in a data descriptor that follows the data */
    readonly dataDescriptor: boolean;
}

/** An entry of a forward-only archive ({@link parseZipStream}). */
export interface ZipStreamEntry extends AsyncIterable<Uint8Array> {
    /** path in the archive, as recorded (decoded as UTF-8) */
    readonly path: string;
    readonly type: "Directory" | "File";
    /**
     * The inflated data, chunk by chunk, in one pass (iterating again continues the same pass). The chunks are new
     * arrays, not views of the input. The iteration throws when a limit is exceeded, the entry is encrypted or uses an
     * unknown compression method, or its data is invalid or does not have the recorded size.
     */
    [Symbol.asyncIterator](): AsyncIterator<Uint8Array>;
}

class StreamEntry implements ZipStreamEntry {
    readonly path: string;
    readonly type: "Directory" | "File";
    /** the compressed data, in slices */
    private readonly raw: AsyncGenerator<Uint8Array, void, undefined>;
    private chunks: AsyncGenerator<Uint8Array, void, undefined> | null = null;

    constructor(
        private readonly parser: ZipStreamParser,
        private readonly header: LocalHeader
    ) {
        this.path = header.name;
        this.type = header.name.endsWith("/") ? "Directory" : "File";
        this.raw = parser._rawData(header);
    }

    // the inflated data, chunk by chunk (one pass)
    [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
        this.chunks ??= this.inflate();
        return this.chunks;
    }

    private async *inflate(): AsyncGenerator<Uint8Array, void, undefined> {
        const { name, method, flags, dataDescriptor, compressedSize, uncompressedSize } = this.header;
        const { limits, backend } = this.parser;
        checkMethod(name, method, flags);
        if (!dataDescriptor) {
            // only the entries that are read count: skipped entries are never inflated
            limits.checkEntry(name, compressedSize, uncompressedSize);
        }
        // sizes of the data read so far: the limits are checked on every inflated chunk
        let size = 0;
        let compressed = 0;
        const consume = (slice: Uint8Array): void => {
            compressed += slice.length;
            limits.addCompressed(slice.length);
        };
        const count = (chunk: Uint8Array): void => {
            size += chunk.length;
            limits.add(name, chunk.length, size, compressed);
        };
        // not `for await`: when the consumer stops early, it would close `raw`, which drain() then needs
        // to skip the rest of the entry
        const next = (): Promise<IteratorResult<Uint8Array, void>> => this.raw.next();
        // the size recorded in the local header or, once the data is read, in the data descriptor
        const checkSize = (): void => {
            if (size !== this.header.uncompressedSize) {
                throw zipError(`"${name}" does not have the size recorded in the archive`);
            }
        };
        if (method === METHOD_STORE) {
            for (let slice = await next(); !slice.done; slice = await next()) {
                consume(slice.value);
                count(slice.value);
                // a copy: not a view of the input
                yield new Uint8Array(slice.value);
            }
            checkSize();
            return;
        }
        const inflater = createInflater(count, backend);
        try {
            // hold one slice back to know which one is the last
            let previous: Uint8Array | null = null;
            for (let slice = await next(); !slice.done; slice = await next()) {
                if (previous) {
                    consume(previous);
                    yield* await inflater.push(previous, false);
                }
                previous = slice.value;
            }
            previous ??= new Uint8Array(0);
            consume(previous);
            yield* await inflater.push(previous, true);
            checkSize();
        } finally {
            inflater.close();
        }
    }

    /** skips whatever has not been read; compressed data of an unread entry is skipped without inflating it */
    async drain(): Promise<void> {
        let slice = await this.raw.next();
        while (!slice.done) {
            slice = await this.raw.next();
        }
    }
}

class ZipStreamParser {
    private readonly source: ByteSource;
    readonly limits: ZipLimits;
    readonly backend: DeflateBackend;
    /** maximum size of the compressed slices handed to the decompressor */
    private readonly slice: number;

    constructor(chunks: AsyncIterable<unknown, unknown, undefined>, options: ZipReadOptions | undefined) {
        this.source = new ByteSource(chunks);
        this.limits = new ZipLimits(options, STREAM_LIMITS);
        this.backend = selectBackend(options && options._backend);
        this.slice = sliceSize(this.backend);
    }

    async *entries(): AsyncGenerator<ZipStreamEntry, void, undefined> {
        const { source } = this;
        try {
            let first = true;
            for (;;) {
                if (!(await source.ensure(4))) {
                    if (source.available === 0 && first) {
                        return;
                    }
                    throw new Error("Unexpected end of zip data");
                }
                const signature = source.u32(0);
                if (first && (signature === SIG_SPANNING || signature === SIG_DATA_DESCRIPTOR)) {
                    // split archive marker
                    source.take(4);
                    first = false;
                    continue;
                }
                first = false;
                if (signature !== SIG_LOCAL) {
                    if (NEXT_RECORD.includes(signature)) {
                        // central directory: no more entries
                        return;
                    }
                    throw zipError(`unexpected signature 0x${signature.toString(16)}`);
                }
                const entry = new StreamEntry(this, await this.readLocalHeader());
                yield entry;
                await entry.drain();
            }
        } finally {
            await source.close();
        }
    }

    private async readLocalHeader(): Promise<LocalHeader> {
        const { source } = this;
        await source.require(30);
        const flags = u16(source.buffer, source.pos + 6);
        const method = u16(source.buffer, source.pos + 8);
        const compressedSize: Zip64Field = { value: source.u32(18) };
        const uncompressedSize: Zip64Field = { value: source.u32(22) };
        const nameLength = u16(source.buffer, source.pos + 26);
        const extraLength = u16(source.buffer, source.pos + 28);
        await source.require(30 + nameLength + extraLength);
        const header = source.take(30 + nameLength + extraLength);
        const name = decodeName(header.subarray(30, 30 + nameLength));
        const zip64Extra = findZip64Extra(header, 30 + nameLength, extraLength);
        if (zip64Extra && zip64Extra.end - zip64Extra.start >= 16) {
            // in a local header the ZIP64 field holds both sizes, uncompressed first, whichever is flagged
            if (uncompressedSize.value === MAX_U32 || compressedSize.value === MAX_U32) {
                uncompressedSize.value = u64(header, zip64Extra.start);
                compressedSize.value = u64(header, zip64Extra.start + 8);
            }
        } else {
            applyZip64(header, 30 + nameLength, extraLength, [uncompressedSize, compressedSize], name);
        }
        const zip64 = Boolean(zip64Extra);
        return {
            name,
            flags,
            method,
            compressedSize: compressedSize.value,
            uncompressedSize: uncompressedSize.value,
            zip64,
            dataDescriptor: Boolean(flags & FLAG_DATA_DESCRIPTOR),
        };
    }

    /**
     * The compressed data of an entry, in slices of at most `slice` bytes (views of the input).
     *
     * @internal used by StreamEntry
     */
    async *_rawData(header: LocalHeader): AsyncGenerator<Uint8Array, void, undefined> {
        if (header.dataDescriptor) {
            yield* this.rawDataUntilDescriptor(header);
            return;
        }
        const { source } = this;
        let remaining = header.compressedSize;
        while (remaining > 0) {
            if (source.available === 0) {
                await source.require(1);
            }
            const length = Math.min(remaining, source.available, this.slice);
            remaining -= length;
            yield source.take(length);
        }
    }

    // Sizes unknown (general purpose flag bit 3): the data ends with a data descriptor. Its signature
    // may also occur inside the compressed data, so a candidate is accepted only when the compressed
    // size it records matches the number of bytes seen.
    private async *rawDataUntilDescriptor(header: LocalHeader): AsyncGenerator<Uint8Array, void, undefined> {
        const { source } = this;
        // compressed bytes already handed out
        let size = 0;
        for (;;) {
            const { buffer, pos, available } = source;
            // bytes known to be data (before any possible descriptor)
            let safe = available;
            let found: DescriptorLength = 0;
            for (let at = buffer.indexOf(0x50, pos) - pos; at >= 0; at = buffer.indexOf(0x50, pos + at + 1) - pos) {
                if (at + 4 > available) {
                    safe = at;
                    break;
                }
                if (u32(buffer, pos + at) === SIG_DATA_DESCRIPTOR) {
                    // up to 24 bytes for a ZIP64 descriptor, + 4 to check the signature that follows
                    if (at + 28 > available && !source.ended) {
                        safe = at;
                        break;
                    }
                    found = matchDescriptor(header, buffer, pos + at, available - at, size + at);
                    if (found) {
                        safe = at;
                        break;
                    }
                }
            }
            if (safe > 0) {
                size += safe;
                yield* sliced(source.take(safe), this.slice);
            }
            if (found) {
                const descriptor = source.take(found);
                // the uncompressed size it records
                header.uncompressedSize = found === 24 ? u64(descriptor, 16) : u32(descriptor, 12);
                return;
            }
            if (source.ended) {
                throw new Error("Unexpected end of zip data");
            }
            // at the end of the input, the loop checks the tail once more
            await source.more();
        }
    }
}

/** length of a data descriptor: 16 (32-bit sizes) or 24 (64-bit sizes) bytes; 0 for none */
type DescriptorLength = 0 | 16 | 24;

// the length of the data descriptor starting at `pos` if it is valid, else 0
function matchDescriptor(
    header: LocalHeader,
    buffer: Uint8Array,
    pos: number,
    available: number,
    compressedSize: number
): DescriptorLength {
    if (available < 16 || u32(buffer, pos + 8) !== compressedSize % TWO_32) {
        return 0;
    }
    const isNext = (offset: number): boolean => available < offset + 4 || NEXT_RECORD.includes(u32(buffer, pos + offset));
    const high = u32(buffer, pos + 12);
    if (header.zip64 || compressedSize >= TWO_32) {
        // 64-bit sizes
        return available >= 24 && high === Math.floor(compressedSize / TWO_32) ? 24 : 0;
    }
    if (isNext(16)) {
        return 16;
    }
    // some writers use 64-bit sizes without a ZIP64 field in the local header
    if (high === 0 && available >= 24 && isNext(24)) {
        return 24;
    }
    return 0;
}

function* sliced(data: Uint8Array, size: number): Generator<Uint8Array, void, undefined> {
    for (let pos = 0; pos < data.length; pos += size) {
        yield data.subarray(pos, pos + size);
    }
}

/**
 * Reads a zip archive from an async iterable of byte chunks (e.g. a Node.js Readable), local header by local header.
 * The data of an entry must be read before the next entry is requested: whatever is left is skipped (never inflated).
 * Stopping the iteration early releases the input (its iterator's `return()`, which destroys a stream).
 *
 * @param chunks - Uint8Array / Buffer / ArrayBuffer / typed array chunks
 * @param options - decompression limits, defaulting to {@link STREAM_LIMITS}
 * @returns the entries in archive order
 * @throws TypeError on an invalid limit (at once) or input chunk; Error (from the iteration) on an invalid or truncated
 * archive
 */
function parseZipStream(
    chunks: AsyncIterable<unknown, unknown, undefined>,
    options?: ZipReadOptions
): AsyncGenerator<ZipStreamEntry, void, undefined> {
    return new ZipStreamParser(chunks, options).entries();
}

export { readZip, parseZipStream, DOCUMENT_LIMITS, STREAM_LIMITS };
