// Zip container reading. Decompression comes from deflate.js (node:zlib on Node.js and Bun, fflate in
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
const { createInflater, inflateWhole, selectBackend, sliceSize } = require("./deflate");

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

// The document reader (xlsx.load / read / readFile) materialises every part: a single part cannot be
// parsed past V8's maximum string length (~512 MiB) anyway, and the whole model has to fit in memory.
const DOCUMENT_LIMITS = {
    maxEntrySize: 1 * GiB,
    maxUncompressedSize: 2 * GiB,
    maxCompressionRatio: 1000,
};
// The streaming reader parses worksheets without materialising them (multi-GiB sheets are its purpose);
// worksheets that arrive before the parts they depend on are spooled to temp files, which these sizes
// also bound.
const STREAM_LIMITS = {
    maxEntrySize: 16 * GiB,
    maxUncompressedSize: 32 * GiB,
    maxCompressionRatio: 1000,
};
// the compression ratio of a part, and of all the parts together, is only checked past this size: tiny parts can have any
// ratio
const RATIO_MIN_SIZE = 10 * MiB;

const textDecoder = new TextDecoder("utf-8");

function zipError(message) {
    return new Error(`Invalid or unsupported zip file: ${message}`);
}

function u16(data, pos) {
    return data[pos] | (data[pos + 1] << 8);
}

function u32(data, pos) {
    return (data[pos] | (data[pos + 1] << 8) | (data[pos + 2] << 16) | (data[pos + 3] << 24)) >>> 0;
}

function u64(data, pos) {
    const value = u32(data, pos + 4) * TWO_32 + u32(data, pos);
    if (value > Number.MAX_SAFE_INTEGER) {
        throw zipError("64-bit size or offset out of range");
    }
    return value;
}

function decodeName(bytes) {
    // like JSZip: names are decoded as UTF-8 whether or not the UTF-8 flag (bit 11) is set
    return textDecoder.decode(bytes);
}

// the value of the ZIP64 extended information extra field (header ID 1), or null
function findZip64Extra(data, start, length) {
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
function applyZip64(data, extraStart, extraLength, fields, name) {
    const wanted = fields.filter(({ value }) => value === MAX_U32);
    if (!wanted.length) {
        return;
    }
    const extra = findZip64Extra(data, extraStart, extraLength);
    if (!extra || extra.end - extra.start < wanted.length * 8) {
        throw zipError(`missing ZIP64 extra field for "${name}"`);
    }
    wanted.forEach((field, index) => {
        field.value = u64(data, extra.start + index * 8);
    });
}

// =============================================================================
// Decompression limits

function limitValue(options, name, defaults) {
    const value = options && options[name] !== undefined ? options[name] : defaults[name];
    if (!(typeof value === "number" && value > 0)) {
        throw new TypeError(`${name} must be a positive number, got ${value}`);
    }
    return value;
}

class ZipLimits {
    constructor(options, defaults) {
        this.maxEntrySize = limitValue(options, "maxEntrySize", defaults);
        this.maxUncompressedSize = limitValue(options, "maxUncompressedSize", defaults);
        this.maxCompressionRatio = limitValue(options, "maxCompressionRatio", defaults);
        this.total = 0;
        this.declaredTotal = 0;
        // compressed bytes of the entries inflated (total comes from them)
        this.compressedTotal = 0;
    }

    // declared sizes (central directory or local header): fail before any work
    checkDeclared(name, compressedSize, uncompressedSize) {
        this.declaredTotal += uncompressedSize;
        this._check(name, uncompressedSize, compressedSize, Math.max(this.total, this.declaredTotal));
    }

    // declared sizes of one entry
    checkEntry(name, compressedSize, uncompressedSize) {
        this._check(name, uncompressedSize, compressedSize, this.total);
    }

    // compressed data about to be inflated
    addCompressed(bytes) {
        this.compressedTotal += bytes;
    }

    // actual sizes, while inflating
    add(name, bytes, entrySize, compressedSize) {
        this.total += bytes;
        this._check(name, entrySize, compressedSize, this.total);
        // small parts too: the content of the archive is bounded by its size (many parts that share their data, or many
        // parts that each compress at the maximum ratio)
        if (this.total > RATIO_MIN_SIZE && this.total > this.compressedTotal * this.maxCompressionRatio) {
            throw new Error(`Zip content exceeds maxCompressionRatio (${this.maxCompressionRatio})`);
        }
    }

    _check(name, entrySize, compressedSize, total) {
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

function checkMethod(name, method, flags) {
    if (flags & FLAG_ENCRYPTED) {
        throw new Error(`Encrypted zip entries are not supported: "${name}"`);
    }
    if (method !== METHOD_STORE && method !== METHOD_DEFLATE) {
        throw zipError(`unknown compression method ${method} for "${name}"`);
    }
}

// =============================================================================
// In-memory archive

function findEndOfCentralDirectory(data) {
    // the record is 22 bytes followed by a comment of up to 65535 bytes
    const min = Math.max(0, data.length - 22 - MAX_U16);
    for (let pos = data.length - 22; pos >= min; pos--) {
        if (u32(data, pos) === SIG_EOCD) {
            return pos;
        }
    }
    throw zipError("end of central directory not found (is this an xlsx file?)");
}

function readCentralDirectoryLocation(data) {
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

// Each entry must have its own data: entries that share their data (the central directory points several times at
// the same local header, or at a local header within the data of another entry) multiply the inflated output without
// adding compressed data (overlapping zip bomb). No zip writer produces them.
function checkOverlaps(data, entries) {
    const ranges = entries.map(({ name, localOffset, compressedSize }) => {
        // the lengths of the local header, or a lower bound when it is invalid (then reading the entry fails)
        const valid = localOffset + 30 <= data.length && u32(data, localOffset) === SIG_LOCAL;
        const header = valid ? 30 + u16(data, localOffset + 26) + u16(data, localOffset + 28) : 30;
        return { name, start: localOffset, end: localOffset + header + compressedSize };
    });
    ranges.sort((a, b) => a.start - b.start);
    for (let i = 1; i < ranges.length; i++) {
        if (ranges[i].start < ranges[i - 1].end) {
            throw zipError(`entries "${ranges[i - 1].name}" and "${ranges[i].name}" overlap`);
        }
    }
}

function readCentralDirectory(data, limits) {
    const { count, start, end, shift } = readCentralDirectoryLocation(data);
    // every entry of the central directory, and the entry kept for each name
    const all = [];
    const entries = new Map();
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
        const uncompressedSize = { value: u32(data, pos + 24) };
        const compressedSize = { value: u32(data, pos + 20) };
        const localOffset = { value: u32(data, pos + 42) };
        applyZip64(data, extraStart, extraLength, [uncompressedSize, compressedSize, localOffset], name);

        limits.checkDeclared(name, compressedSize.value, uncompressedSize.value);
        const entry = {
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

function entryData(data, entry) {
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

// Reads an in-memory zip archive (a Uint8Array, a Buffer, an ArrayBuffer...). Returns its entries in
// central directory order: {name, dir, async data()}, data() inflating the entry into a Uint8Array.
function readZip(buffer, options) {
    if (!(ArrayBuffer.isView(buffer) || buffer instanceof ArrayBuffer || Array.isArray(buffer))) {
        throw new TypeError("Can't read the zip file: expected a Buffer, a Uint8Array, an ArrayBuffer or a Blob");
    }
    const data = Array.isArray(buffer) ? new Uint8Array(buffer) : toBytes(buffer);
    const limits = new ZipLimits(options, DOCUMENT_LIMITS);
    const backend = selectBackend(options && options._backend);
    return readCentralDirectory(data, limits).map((entry) => ({
        name: entry.name,
        dir: entry.dir,
        async data() {
            const { name, method, flags, uncompressedSize } = entry;
            checkMethod(name, method, flags);
            const compressed = entryData(data, entry);
            limits.addCompressed(compressed.length);
            // the declared size was checked against the limits: inflating stops past it
            let output;
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
        },
    }));
}

// =============================================================================
// Forward-only archive

function toBytes(value) {
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

class ByteSource {
    constructor(chunks) {
        this.iterator = chunks[Symbol.asyncIterator]();
        this.buffer = new Uint8Array(0);
        this.pos = 0;
        this.ended = false;
    }

    get available() {
        return this.buffer.length - this.pos;
    }

    // read one more chunk; false at the end of the input
    async more() {
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

    async ensure(length) {
        while (this.available < length) {
            if (!(await this.more())) {
                return false;
            }
        }
        return true;
    }

    async require(length) {
        if (!(await this.ensure(length))) {
            throw new Error("Unexpected end of zip data");
        }
    }

    take(length) {
        const result = this.buffer.subarray(this.pos, this.pos + length);
        this.pos += length;
        return result;
    }

    u32(offset) {
        return u32(this.buffer, this.pos + offset);
    }

    // stop reading: releases the input (e.g. destroys a stream)
    async close() {
        if (!this.ended && typeof this.iterator.return === "function") {
            this.ended = true;
            await this.iterator.return();
        }
    }
}

const NEXT_RECORD = [SIG_LOCAL, SIG_CENTRAL, SIG_ZIP64_EOCD, SIG_EOCD];

class StreamEntry {
    constructor(parser, header) {
        this.parser = parser;
        this.header = header;
        this.path = header.name;
        this.type = header.name.endsWith("/") ? "Directory" : "File";
        this.raw = parser._rawData(header);
        this.chunks = null;
    }

    // the inflated data, chunk by chunk (one pass)
    [Symbol.asyncIterator]() {
        if (!this.chunks) {
            this.chunks = this._inflate();
        }
        return this.chunks;
    }

    async *_inflate() {
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
        const consume = (slice) => {
            compressed += slice.length;
            limits.addCompressed(slice.length);
        };
        const count = (chunk) => {
            size += chunk.length;
            limits.add(name, chunk.length, size, compressed);
        };
        // not `for await`: when the consumer stops early, it would close `raw`, which drain() then needs
        // to skip the rest of the entry
        const next = () => this.raw.next();
        // the size recorded in the local header or, once the data is read, in the data descriptor
        const checkSize = () => {
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
            let previous = null;
            for (let slice = await next(); !slice.done; slice = await next()) {
                if (previous) {
                    consume(previous);
                    yield* await inflater.push(previous, false);
                }
                previous = slice.value;
            }
            previous = previous || new Uint8Array(0);
            consume(previous);
            yield* await inflater.push(previous, true);
            checkSize();
        } finally {
            inflater.close();
        }
    }

    // skip whatever has not been read
    async drain() {
        // eslint-disable-next-line no-unused-vars
        for await (const slice of this.raw) {
            // compressed data of an unread entry is skipped without inflating it
        }
    }
}

class ZipStreamParser {
    constructor(chunks, options) {
        this.source = new ByteSource(chunks);
        this.limits = new ZipLimits(options, STREAM_LIMITS);
        this.backend = selectBackend(options && options._backend);
        this.slice = sliceSize(this.backend);
    }

    async *entries() {
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
                const entry = new StreamEntry(this, await this._readLocalHeader());
                yield entry;
                await entry.drain();
            }
        } finally {
            await source.close();
        }
    }

    async _readLocalHeader() {
        const { source } = this;
        await source.require(30);
        const flags = u16(source.buffer, source.pos + 6);
        const method = u16(source.buffer, source.pos + 8);
        const compressedSize = { value: source.u32(18) };
        const uncompressedSize = { value: source.u32(22) };
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

    // the compressed data of an entry, in slices
    async *_rawData(header) {
        if (header.dataDescriptor) {
            yield* this._rawDataUntilDescriptor(header);
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
    async *_rawDataUntilDescriptor(header) {
        const { source } = this;
        // compressed bytes already handed out
        let size = 0;
        for (;;) {
            const { buffer, pos, available } = source;
            // bytes known to be data (before any possible descriptor)
            let safe = available;
            let found = null;
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

// the length of the data descriptor starting at `pos` if it is valid, else 0
function matchDescriptor(header, buffer, pos, available, compressedSize) {
    if (available < 16 || u32(buffer, pos + 8) !== compressedSize % TWO_32) {
        return 0;
    }
    const isNext = (offset) => available < offset + 4 || NEXT_RECORD.includes(u32(buffer, pos + offset));
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

function* sliced(data, size) {
    for (let pos = 0; pos < data.length; pos += size) {
        yield data.subarray(pos, pos + size);
    }
}

// Reads a zip archive from an async iterable of byte chunks. Yields entries in archive order
// ({path, type, [Symbol.asyncIterator]}); the data of an entry must be read before the next entry is
// requested, whatever is left is skipped.
function parseZipStream(chunks, options) {
    return new ZipStreamParser(chunks, options).entries();
}

module.exports = {
    readZip,
    parseZipStream,
    DOCUMENT_LIMITS,
    STREAM_LIMITS,
};
