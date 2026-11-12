// DEFLATE backends of the zip code (zip-stream.ts, zip-reader.ts):
// - "zlib": node:zlib, native, on Node.js and Bun;
// - "fflate": pure JavaScript, in the browser bundle, where `zlib` is an empty module (scripts/build.mjs).
// Both produce standard raw DEFLATE data: either one reads what the other writes.
import * as zlib from "zlib";
import { Inflate, ZipDeflate, ZipPassThrough, type FlateError } from "fflate";

/** Name of a DEFLATE implementation: `"zlib"` (node:zlib, Node.js and Bun) or `"fflate"` (pure JavaScript, browser). */
export type DeflateBackend = "zlib" | "fflate";

/** DEFLATE level: 0 stores the data uncompressed, 1 is the fastest, 9 the smallest. */
export type CompressionLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/**
 * Whether `value` is a {@link CompressionLevel}: an integer from 0 to 9.
 *
 * @param value - any value (zip options come from JavaScript callers too)
 */
export function isCompressionLevel(value: unknown): value is CompressionLevel {
    return typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 9;
}

/**
 * True when node:zlib is usable (Node.js, Bun); false in the browser bundle, where `zlib` is an empty module and every
 * function checked here is `undefined` whatever its declared type.
 */
const NATIVE: boolean =
    typeof zlib.deflateRawSync === "function" &&
    typeof zlib.inflateRawSync === "function" &&
    typeof zlib.createInflateRaw === "function" &&
    typeof zlib.crc32 === "function";

/**
 * The backend to use: the native one when available. Tests can ask for one with the internal `_backend` option.
 *
 * @param name - `"zlib"`, `"fflate"`, or a falsy value for the default backend
 * @throws TypeError when `name` is not a backend name; Error when `"zlib"` is asked for and node:zlib is not available
 */
function selectBackend(name?: string): DeflateBackend {
    const backend = name || (NATIVE ? "zlib" : "fflate");
    if (backend !== "zlib" && backend !== "fflate") {
        throw new TypeError(`Unknown deflate backend: ${backend}`);
    }
    if (backend === "zlib" && !NATIVE) {
        throw new Error("node:zlib is not available");
    }
    return backend;
}

/**
 * Compressed bytes handed to a decompressor at once. DEFLATE expands at most ~1032:1, so the output of one slice stays
 * bounded before the limits are checked again.
 */
function sliceSize(backend: DeflateBackend): number {
    return backend === "zlib" ? 64 * 1024 : 16 * 1024;
}

// -----------------------------------------------------------------------------
// Compression of one zip entry

/**
 * Receives the compressed chunks of an entry, in order: either an error (and nothing more is produced), or chunks, the
 * last call having `final` set.
 */
export type CompressorCallback = (
    ...args: [error: Error, chunks: null, final: false] | [error: null, chunks: readonly Uint8Array[], final: boolean]
) => void;

/**
 * Compression of one zip entry: `push(bytes, final)`; the compressed chunks are handed, in order, to the
 * {@link CompressorCallback}, synchronously or later. `crc` (CRC-32) and `size` (uncompressed bytes) describe the data
 * pushed. Level 0 stores the data (copied: callers may reuse their buffers).
 */
export interface Compressor {
    /** CRC-32 of the data pushed (final once the last chunk was handed out) */
    readonly crc: number;
    /** uncompressed bytes pushed */
    readonly size: number;
    /** uncompressed bytes being compressed asynchronously (backpressure) */
    readonly inFlight: number;
    /**
     * Compresses `data`. It is not retained after the call returns.
     *
     * @param final - true for the last data of the entry (may be empty)
     */
    push(data: Uint8Array, final: boolean): void;
}

// Data that arrives at once is compressed synchronously. A streamed entry is compressed in segments,
// asynchronously (in libuv's thread pool, in parallel with the producer, as archiver did): each segment
// is an independent DEFLATE run ending with a sync flush (byte aligned, not final), the last one is
// finished, and their concatenation is one valid DEFLATE stream (as with pigz).
const SEGMENT = 1024 * 1024;

/** a segment of a streamed entry, in compression; `output` is set once it is compressed */
interface Segment {
    output: Uint8Array | null;
    readonly final: boolean;
}

class ZlibCompressor implements Compressor {
    crc = 0;
    size = 0;
    inFlight = 0;
    /** the segment being filled, and the bytes in it */
    private buffer: Buffer | null = null;
    private length = 0;
    /** segments being compressed, in order */
    private segments: Segment[] = [];

    constructor(
        private readonly level: CompressionLevel,
        private readonly ondata: CompressorCallback
    ) {}

    push(data: Uint8Array, final: boolean): void {
        this.crc = zlib.crc32(data, this.crc);
        this.size += data.length;
        if (!this.level) {
            this.ondata(null, data.length ? [new Uint8Array(data)] : [], final);
            return;
        }
        if (final && !this.buffer && !this.segments.length) {
            // all the data at once
            this.ondata(null, [zlib.deflateRawSync(data, { level: this.level })], true);
            return;
        }
        for (let pos = 0; pos < data.length;) {
            if (!this.buffer) {
                this.buffer = Buffer.allocUnsafe(SEGMENT);
                this.length = 0;
            }
            const length = Math.min(SEGMENT - this.length, data.length - pos);
            this.buffer.set(data.subarray(pos, pos + length), this.length);
            this.length += length;
            pos += length;
            if (this.length === SEGMENT) {
                this.compress(this.buffer, false);
                this.buffer = null;
            }
        }
        if (final) {
            this.compress(this.buffer ? this.buffer.subarray(0, this.length) : Buffer.alloc(0), true);
            this.buffer = null;
        }
    }

    private compress(data: Buffer, final: boolean): void {
        const segment: Segment = { output: null, final };
        this.segments.push(segment);
        this.inFlight += data.length;
        const options = final ? { level: this.level } : { level: this.level, finishFlush: zlib.constants.Z_SYNC_FLUSH };
        zlib.deflateRaw(data, options, (error, output) => {
            this.inFlight -= data.length;
            if (error) {
                this.segments = [];
                this.ondata(error, null, false);
                return;
            }
            segment.output = output;
            // the segments are handed out in order: a later one that is done first waits for the earlier ones
            for (let done = this.segments[0]; done?.output; done = this.segments[0]) {
                this.segments.shift();
                this.ondata(null, [done.output], done.final);
            }
        });
    }
}

/** compresses synchronously: nothing is ever in flight */
class FflateCompressor implements Compressor {
    crc = 0;
    size = 0;
    readonly inFlight = 0;
    private readonly engine: ZipDeflate | ZipPassThrough;

    constructor(level: CompressionLevel, ondata: CompressorCallback) {
        this.engine = level ? new ZipDeflate("", { level }) : new ZipPassThrough("");
        this.engine.ondata = (error: FlateError | null, chunk: Uint8Array, final: boolean): void => {
            if (error) {
                ondata(error, null, false);
                return;
            }
            if (final) {
                this.crc = this.engine.crc;
            }
            this.size = this.engine.size;
            // ZipPassThrough hands out the caller's buffer
            ondata(null, [level ? chunk : new Uint8Array(chunk)], final);
        };
    }

    push(data: Uint8Array, final: boolean): void {
        this.engine.push(data, final);
    }
}

/**
 * Creates the compressor of one zip entry.
 *
 * @param level - 0 stores the data
 * @param backend - from {@link selectBackend}
 * @param ondata - receives the compressed chunks, in order
 */
function createCompressor(level: CompressionLevel, backend: DeflateBackend, ondata: CompressorCallback): Compressor {
    return backend === "zlib" ? new ZlibCompressor(level, ondata) : new FflateCompressor(level, ondata);
}

// -----------------------------------------------------------------------------
// Decompression

/** the `code` property of a thrown value (Node.js errors), if any */
function errorCode(error: unknown): unknown {
    return typeof error === "object" && error !== null && "code" in error ? error.code : undefined;
}

/**
 * Inflates a whole entry (raw DEFLATE data). Inflating stops as soon as the output exceeds `maxSize`.
 *
 * @param data - compressed data
 * @param maxSize - maximum number of bytes of the output
 * @returns the inflated data, or null if it is larger than `maxSize`
 * @throws on invalid DEFLATE data
 */
function inflateWhole(data: Uint8Array, maxSize: number, backend: DeflateBackend): Uint8Array | null {
    if (backend === "zlib") {
        try {
            return zlib.inflateRawSync(data, { maxOutputLength: maxSize + 1 });
        } catch (error) {
            if (errorCode(error) === "ERR_BUFFER_TOO_LARGE") {
                return null;
            }
            throw error;
        }
    }
    const chunks: Uint8Array[] = [];
    let size = 0;
    const inflate = new Inflate((chunk) => {
        size += chunk.length;
        chunks.push(chunk);
    });
    const slice = sliceSize(backend);
    for (let pos = 0; pos < data.length || pos === 0; pos += slice) {
        const end = Math.min(pos + slice, data.length);
        inflate.push(data.subarray(pos, end), end === data.length);
        if (size > maxSize) {
            return null;
        }
    }
    const result = new Uint8Array(size);
    let offset = 0;
    for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.length;
    }
    return result;
}

/**
 * Streaming decompression of one entry (raw DEFLATE data): `await push(slice, final)` returns the chunks inflated so
 * far. Call `close()` when done, or to give up.
 */
export interface Inflater {
    /**
     * Inflates one slice of compressed data.
     *
     * @param final - true for the last slice (may be empty)
     * @returns the chunks inflated from this slice (the fflate backend returns them synchronously)
     * @throws on invalid DEFLATE data, or what the `onChunk` callback threw
     */
    push(slice: Uint8Array, final: boolean): Promise<Uint8Array[]> | Uint8Array[];
    /** releases the decompressor */
    close(): void;
}

/** sees every inflated chunk as it is produced; may throw to stop the decompression (limits) */
export type InflatedChunkCallback = (chunk: Uint8Array) => void;

class ZlibInflater implements Inflater {
    private chunks: Uint8Array[] = [];
    /** the first failure (of the stream or of onChunk), thrown by push() */
    private failure: { readonly error: unknown } | null = null;
    private readonly stream: zlib.InflateRaw;

    constructor(onChunk: InflatedChunkCallback) {
        this.stream = zlib.createInflateRaw();
        this.stream.on("data", (chunk: Buffer) => {
            if (this.failure) {
                return;
            }
            try {
                onChunk(chunk);
                this.chunks.push(chunk);
            } catch (error) {
                this.fail(error);
            }
        });
        this.stream.on("error", (error) => this.fail(error));
    }

    private fail(error: unknown): void {
        if (!this.failure) {
            this.failure = { error };
            this.stream.destroy();
        }
    }

    async push(slice: Uint8Array, final: boolean): Promise<Uint8Array[]> {
        if (!this.failure) {
            const { stream } = this;
            await new Promise<void>((resolve) => {
                // 'close': a failed stream does not call back the write in progress
                const done = (): void => {
                    stream.removeListener("close", done);
                    stream.removeListener("end", done);
                    resolve();
                };
                stream.once("close", done);
                if (final) {
                    stream.once("end", done);
                    stream.end(slice);
                } else {
                    // called once the slice is inflated
                    stream.write(slice, done);
                }
            });
        }
        if (this.failure) {
            throw this.failure.error;
        }
        const { chunks } = this;
        this.chunks = [];
        return chunks;
    }

    close(): void {
        if (!this.stream.destroyed) {
            this.stream.destroy();
        }
    }
}

class FflateInflater implements Inflater {
    private chunks: Uint8Array[] = [];
    private readonly inflate: Inflate;

    constructor(onChunk: InflatedChunkCallback) {
        this.inflate = new Inflate((chunk) => {
            if (chunk.length) {
                onChunk(chunk);
                this.chunks.push(chunk);
            }
        });
    }

    push(slice: Uint8Array, final: boolean): Uint8Array[] {
        this.inflate.push(slice, final);
        const { chunks } = this;
        this.chunks = [];
        return chunks;
    }

    close(): void {}
}

/**
 * Creates the streaming decompressor of one entry.
 *
 * @param onChunk - sees every inflated chunk as it is produced; may throw to stop the decompression
 * @param backend - from {@link selectBackend}
 */
function createInflater(onChunk: InflatedChunkCallback, backend: DeflateBackend): Inflater {
    return backend === "zlib" ? new ZlibInflater(onChunk) : new FflateInflater(onChunk);
}

export { NATIVE, selectBackend, sliceSize, createCompressor, inflateWhole, createInflater };
