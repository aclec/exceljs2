/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
// DEFLATE backends of the zip code (zip-stream.js, zip-reader.js):
// - "zlib": node:zlib, native, on Node.js and Bun;
// - "fflate": pure JavaScript, in the browser bundle, where `zlib` is an empty module (scripts/build.mjs).
// Both produce standard raw DEFLATE data: either one reads what the other writes.
import * as zlib from "zlib";
import { Inflate, ZipDeflate, ZipPassThrough } from "fflate";

const NATIVE =
    typeof zlib.deflateRawSync === "function" &&
    typeof zlib.inflateRawSync === "function" &&
    typeof zlib.createInflateRaw === "function" &&
    typeof zlib.crc32 === "function";

// the backend to use: the native one when available (tests can ask for one with the internal `_backend` option)
function selectBackend(name) {
    const backend = name || (NATIVE ? "zlib" : "fflate");
    if (backend !== "zlib" && backend !== "fflate") {
        throw new TypeError(`Unknown deflate backend: ${backend}`);
    }
    if (backend === "zlib" && !NATIVE) {
        throw new Error("node:zlib is not available");
    }
    return backend;
}

// compressed data handed to a decompressor at once: DEFLATE expands at most ~1032:1, so the output of one
// slice stays bounded before the limits are checked again
function sliceSize(backend) {
    return backend === "zlib" ? 64 * 1024 : 16 * 1024;
}

// -----------------------------------------------------------------------------
// Compression of one zip entry: push(bytes, final); the compressed chunks are handed, in order, to
// ondata(error, chunks, final), synchronously or later. `crc` (CRC-32) and `size` (uncompressed) describe
// the data pushed. Level 0 stores the data (copied: callers may reuse their buffers).

// Data that arrives at once is compressed synchronously. A streamed entry is compressed in segments,
// asynchronously (in libuv's thread pool, in parallel with the producer, as archiver did): each segment
// is an independent DEFLATE run ending with a sync flush (byte aligned, not final), the last one is
// finished, and their concatenation is one valid DEFLATE stream (as with pigz).
const SEGMENT = 1024 * 1024;

class ZlibCompressor {
    constructor(level, ondata) {
        this.level = level;
        this.ondata = ondata;
        this.crc = 0;
        this.size = 0;
        this.buffer = null;
        this.length = 0;
        // segments being compressed, in order, and their uncompressed size
        this.segments = [];
        this.inFlight = 0;
    }

    push(data, final) {
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
                this._compress(this.buffer, false);
                this.buffer = null;
            }
        }
        if (final) {
            this._compress(this.buffer ? this.buffer.subarray(0, this.length) : Buffer.alloc(0), true);
            this.buffer = null;
        }
    }

    _compress(data, final) {
        const segment = { output: null, final };
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
            while (this.segments.length && this.segments[0].output) {
                const done = this.segments.shift();
                this.ondata(null, [done.output], done.final);
            }
        });
    }
}

// compresses synchronously: nothing is ever in flight
class FflateCompressor {
    constructor(level, ondata) {
        this.inFlight = 0;
        this.engine = level ? new ZipDeflate("", { level }) : new ZipPassThrough("");
        this.engine.ondata = (error, chunk, final) => {
            if (error) {
                ondata(error, null, false);
                return;
            }
            if (final) {
                this.crc = this.engine.crc;
            }
            this.size = this.engine.size;
            ondata(null, [level ? chunk : new Uint8Array(chunk)], final);
        };
        this.crc = 0;
        this.size = 0;
    }

    push(data, final) {
        this.engine.push(data, final);
    }
}

function createCompressor(level, backend, ondata) {
    return backend === "zlib" ? new ZlibCompressor(level, ondata) : new FflateCompressor(level, ondata);
}

// -----------------------------------------------------------------------------
// Decompression

// inflates a whole entry; null if it is larger than maxSize (it is not inflated further)
function inflateWhole(data, maxSize, backend) {
    if (backend === "zlib") {
        try {
            return zlib.inflateRawSync(data, { maxOutputLength: maxSize + 1 });
        } catch (error) {
            if (error.code === "ERR_BUFFER_TOO_LARGE") {
                return null;
            }
            throw error;
        }
    }
    const chunks = [];
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

// Streaming decompression of one entry: `await push(slice, final)` returns the chunks inflated so far.
// onChunk(chunk) sees every chunk as it is produced and may throw to stop the decompression (limits).
class ZlibInflater {
    constructor(onChunk) {
        this.chunks = [];
        this.error = null;
        this.stream = zlib.createInflateRaw();
        this.stream.on("data", (chunk) => {
            if (this.error) {
                return;
            }
            try {
                onChunk(chunk);
                this.chunks.push(chunk);
            } catch (error) {
                this._fail(error);
            }
        });
        this.stream.on("error", (error) => this._fail(error));
    }

    _fail(error) {
        if (!this.error) {
            this.error = error;
            this.stream.destroy();
        }
    }

    async push(slice, final) {
        if (!this.error) {
            const { stream } = this;
            await new Promise((resolve) => {
                // 'close': a failed stream does not call back the write in progress
                const done = () => {
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
        if (this.error) {
            throw this.error;
        }
        const { chunks } = this;
        this.chunks = [];
        return chunks;
    }

    close() {
        if (!this.stream.destroyed) {
            this.stream.destroy();
        }
    }
}

class FflateInflater {
    constructor(onChunk) {
        this.chunks = [];
        this.inflate = new Inflate((chunk) => {
            if (chunk.length) {
                onChunk(chunk);
                this.chunks.push(chunk);
            }
        });
    }

    push(slice, final) {
        this.inflate.push(slice, final);
        const { chunks } = this;
        this.chunks = [];
        return chunks;
    }

    close() {}
}

function createInflater(onChunk, backend) {
    return backend === "zlib" ? new ZlibInflater(onChunk) : new FflateInflater(onChunk);
}

export { NATIVE, selectBackend, sliceSize, createCompressor, inflateWhole, createInflater };
