import { EventEmitter } from "events";
import { Duplex } from "stream";

import { nop } from "./utils";
import StringBuf from "./string-buf";

/** The callback of {@link StreamBuf.write}. */
type WriteCallback = (error?: Error | null) => void;

// =============================================================================
// data chunks - encapsulating incoming data
interface Chunk {
    /** The size in bytes. */
    readonly length: number;
    /** Copies the bytes [sourceStart, sourceEnd) to target at targetOffset. */
    copy(target: Uint8Array, targetOffset: number, sourceStart: number, sourceEnd: number): number;
    toBuffer(): Buffer;
}

/** A string, or a String object. */
type StringLike = string | { valueOf(): string };

class StringChunk implements Chunk {
    private _data: StringLike | ArrayBuffer;
    private _encoding: BufferEncoding | undefined;
    private _buffer: Buffer | undefined;

    constructor(data: StringLike | ArrayBuffer, encoding: BufferEncoding | undefined) {
        this._data = data;
        this._encoding = encoding;
    }

    get length(): number {
        return this.toBuffer().length;
    }

    // copy to target buffer
    copy(target: Uint8Array, targetOffset: number, sourceStart: number, sourceEnd: number): number {
        return this.toBuffer().copy(target, targetOffset, sourceStart, sourceEnd);
    }

    toBuffer(): Buffer {
        if (!this._buffer) {
            const data = this._data;
            // Buffer.from(arrayBuffer, encoding) took the encoding as a byte offset, read as 0
            this._buffer = data instanceof ArrayBuffer ? Buffer.from(data) : Buffer.from(data, this._encoding);
        }
        return this._buffer;
    }
}

class StringBufChunk implements Chunk {
    private _data: StringBuf;

    constructor(data: StringBuf) {
        this._data = data;
    }

    get length(): number {
        return this._data.length;
    }

    // copy to target buffer
    copy(target: Uint8Array, targetOffset: number, sourceStart: number, sourceEnd: number): number {
        return this._data.buffer.copy(target, targetOffset, sourceStart, sourceEnd);
    }

    toBuffer(): Buffer {
        return this._data.toBuffer();
    }
}

class BufferChunk implements Chunk {
    private _data: Buffer;

    constructor(data: Buffer) {
        this._data = data;
    }

    get length(): number {
        return this._data.length;
    }

    // copy to target buffer
    copy(target: Uint8Array, targetOffset: number, sourceStart: number, sourceEnd: number): number {
        return this._data.copy(target, targetOffset, sourceStart, sourceEnd);
    }

    toBuffer(): Buffer {
        return this._data;
    }
}

// =============================================================================
// ReadWriteBuf - a single buffer supporting simple read-write
class ReadWriteBuf {
    size: number;
    // the buffer
    buffer: Buffer;
    // read index
    iRead: number;
    // write index
    iWrite: number;

    constructor(size: number) {
        this.size = size;
        this.buffer = Buffer.alloc(size);
        this.iRead = 0;
        this.iWrite = 0;
    }

    toBuffer(): Buffer {
        if (this.iRead === 0 && this.iWrite === this.size) {
            return this.buffer;
        }

        const buf = Buffer.alloc(this.iWrite - this.iRead);
        this.buffer.copy(buf, 0, this.iRead, this.iWrite);
        return buf;
    }

    get length(): number {
        return this.iWrite - this.iRead;
    }

    get eod(): boolean {
        return this.iRead === this.iWrite;
    }

    get full(): boolean {
        return this.iWrite === this.size;
    }

    // size bytes from the buffer (all of them when undefined or larger than the content); null for 0
    read(size?: number): Buffer | null {
        let buf;
        // read size bytes from buffer and return buffer
        if (size === 0) {
            // special case - return null if no data requested
            return null;
        }

        if (size === undefined || size >= this.length) {
            // if no size specified or size is at least what we have then return all of the bytes
            buf = this.toBuffer();
            this.iRead = this.iWrite;
            return buf;
        }

        // otherwise return a chunk
        buf = Buffer.alloc(size);
        // copies [iRead, size), not [iRead, iRead + size): kept as it always was (StreamBuf#read(size) is unused)
        this.buffer.copy(buf, 0, this.iRead, size);
        this.iRead += size;
        return buf;
    }

    // writes as many bytes of chunk from offset as fit (at most length); returns the number of bytes written
    write(chunk: Chunk, offset: number, length: number): number {
        const size = Math.min(length, this.size - this.iWrite);
        chunk.copy(this.buffer, this.iWrite, offset, offset + size);
        this.iWrite += size;
        return size;
    }
}

/** Options of a {@link StreamBuf}. */
export interface StreamBufOptions {
    /** The size of each internal buffer, in bytes (1 MiB by default). */
    bufSize?: number;
    /** Batch mode: fill a buffer completely before passing the data on to the pipes. */
    batch?: boolean;
}

// =============================================================================
/**
 * A multi-purpose read-write stream:
 * - as a memory buffer: write as much data as you like, then call toBuffer() (or read()) to consolidate it;
 * - as a stream hub: pipe it to several writables;
 * - as a readable stream: feed data into the writable part and have some other code read from it.
 *
 * It implements the read / write / pipe methods of Duplex itself and does not use Duplex's flow control: write() and
 * end() emit 'data', 'readable' and 'finish' directly, and the listener methods (on, addListener, removeListener, off,
 * removeAllListeners) are EventEmitter's, so that adding a 'data' or 'readable' listener does not resume the stream nor
 * read (drain) its data. The Duplex state exists (the constructor builds it): destroy() works and emits 'close'.
 * Async iteration is not supported. write() takes strings, Buffers, ArrayBuffers and StringBufs.
 */
class StreamBuf extends Duplex {
    /** The size of each internal buffer, in bytes. */
    bufSize: number;
    /** The data written and not read yet. */
    buffers: ReadWriteBuf[];
    // batch mode fills a buffer completely before passing the data on
    // to pipes or 'readable' event listeners
    batch: boolean;
    corked: boolean;
    // where in the current writable buffer we're up to
    inPos: number;
    // where in the current readable buffer we've read up to
    outPos: number;
    // consuming pipe streams go here
    pipes: NodeJS.WritableStream[];
    // controls emit('data')
    paused: boolean;
    encoding: BufferEncoding | null;

    constructor(options?: StreamBufOptions) {
        super();
        options = options || {};
        this.bufSize = options.bufSize || 1024 * 1024;
        this.buffers = [];
        this.batch = options.batch || false;
        this.corked = false;
        this.inPos = 0;
        this.outPos = 0;
        this.pipes = [];
        this.paused = false;
        this.encoding = null;
    }

    /** All the data written and not read yet, in one Buffer; null when nothing was written. */
    toBuffer(): Buffer | null {
        switch (this.buffers.length) {
            case 0:
                return null;
            case 1:
                return this.buffers[0]!.toBuffer();
            default:
                return Buffer.concat(this.buffers.map((rwBuf) => rwBuf.toBuffer()));
        }
    }

    // writable
    // event drain - if write returns false (which it won't), indicates when safe to write again.
    // finish - end() has been called
    // pipe(src) - pipe() has been called on readable
    // unpipe(src) - unpipe() has been called on readable
    // error - duh

    private _getWritableBuffer(): ReadWriteBuf {
        const last = this.buffers[this.buffers.length - 1];
        if (last && !last.full) {
            return last;
        }
        const buf = new ReadWriteBuf(this.bufSize);
        this.buffers.push(buf);
        return buf;
    }

    private async _pipe(chunk: { toBuffer(): Buffer }): Promise<void> {
        const write = function (pipe: NodeJS.WritableStream): Promise<void> {
            return new Promise((resolve) => {
                pipe.write(chunk.toBuffer(), () => {
                    resolve();
                });
            });
        };
        await Promise.all(this.pipes.map(write));
    }

    private _writeToBuffers(chunk: Chunk): void {
        let inPos = 0;
        const inLen = chunk.length;
        while (inPos < inLen) {
            // find writable buffer
            const buffer = this._getWritableBuffer();

            // write some data
            inPos += buffer.write(chunk, inPos, inLen - inPos);
        }
    }

    /**
     * Writes data: to the pipes (once they took it, callback is called), else to the internal buffers ('data' and
     * 'readable' are emitted; callback is not called).
     * @param data - A string (or String), a Buffer, an ArrayBuffer or a StringBuf.
     * @param encoding - The encoding of a string (utf8 by default), or the callback.
     * @returns Always true (no backpressure).
     * @throws Error when data is of another type.
     */
    override write(data: unknown, encoding?: BufferEncoding | WriteCallback, callback?: WriteCallback): boolean {
        if (encoding instanceof Function) {
            callback = encoding;
            encoding = "utf8";
        }
        const done = callback || nop;

        // encapsulate data into a chunk
        let chunk: Chunk;
        if (data instanceof StringBuf) {
            chunk = new StringBufChunk(data);
        } else if (data instanceof Buffer) {
            chunk = new BufferChunk(data);
        } else if (typeof data === "string" || data instanceof String || data instanceof ArrayBuffer) {
            chunk = new StringChunk(data, encoding);
        } else {
            throw new Error("Chunk must be one of type String, Buffer or StringBuf.");
        }

        // now, do something with the chunk
        if (this.pipes.length) {
            if (this.batch) {
                this._writeToBuffers(chunk);
                while (!this.corked && this.buffers.length > 1) {
                    void this._pipe(this.buffers.shift()!);
                }
            } else if (!this.corked) {
                void this._pipe(chunk).then(() => done());
            } else {
                this._writeToBuffers(chunk);
                process.nextTick(done);
            }
        } else {
            if (!this.paused) {
                this.emit("data", chunk.toBuffer());
            }

            this._writeToBuffers(chunk);
            this.emit("readable");
        }

        return true;
    }

    override cork(): void {
        this.corked = true;
    }

    // passes the buffered data on to the pipes, if any
    private _flush(): void {
        // if we have comsumers...
        if (this.pipes.length) {
            // and there's stuff not written
            let buffer;
            while ((buffer = this.buffers.shift())) {
                void this._pipe(buffer);
            }
        }
    }

    override uncork(): void {
        this.corked = false;
        this._flush();
    }

    /**
     * Writes the last chunk if any, then passes the buffered data on to the pipes, ends them and emits 'finish'.
     * With a chunk and no pipe, 'finish' is not emitted (write() does not call back without a pipe).
     * @throws Error when chunk is not of a type write() takes; a callback must be the third argument (`end(cb)` throws).
     */
    override end(chunk?: unknown, encoding?: BufferEncoding | WriteCallback, callback?: WriteCallback): this {
        const writeComplete = (error?: Error | null): void => {
            if (error) {
                callback?.(error);
            } else {
                this._flush();
                this.pipes.forEach((pipe) => {
                    pipe.end();
                });
                this.emit("finish");
            }
        };
        if (chunk) {
            this.write(chunk, encoding, writeComplete);
        } else {
            writeComplete();
        }
        return this;
    }

    // readable
    // event readable - some data is now available
    // event data - switch to flowing mode - feeds chunks to handler
    // event end - no more data
    // event close - optional, indicates upstream close
    // event error - duh
    /** Reads (and removes) up to size bytes, or all the data when size is 0 or undefined. */
    override read(size?: number): Buffer {
        let buffers: Buffer[];
        // read min(buffer, size || infinity)
        if (size) {
            buffers = [];
            let first;
            while (size && (first = this.buffers[0]) && !first.eod) {
                // not null: size is not 0
                const buffer = first.read(size)!;
                size -= buffer.length;
                buffers.push(buffer);
                if (first.eod && first.full) {
                    this.buffers.shift();
                }
            }
            return Buffer.concat(buffers);
        }

        buffers = this.buffers.map((buf) => buf.toBuffer()).filter(Boolean);
        this.buffers = [];
        return Buffer.concat(buffers);
    }

    /** Records the encoding (not used: read() and 'data' give Buffers). */
    override setEncoding(encoding: BufferEncoding): this {
        // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
        this.encoding = encoding;
        return this;
    }

    /** Stops the 'data' events (the data is still buffered). */
    override pause(): this {
        this.paused = true;
        return this;
    }

    override resume(): this {
        this.paused = false;
        return this;
    }

    override isPaused(): boolean {
        return !!this.paused;
    }

    /** Adds a destination; when not paused, passes the buffered data on to it and ends (see end()). */
    override pipe<T extends NodeJS.WritableStream>(destination: T): T {
        // add destination to pipe list & write current buffer
        this.pipes.push(destination);
        if (!this.paused && this.buffers.length) {
            this.end();
        }
        return destination;
    }

    override unpipe(destination?: NodeJS.WritableStream): this {
        // remove destination from pipe list
        this.pipes = this.pipes.filter((pipe) => pipe !== destination);
        return this;
    }

    /** @throws Error always: not implemented. */
    override unshift(): never {
        // some numpty has read some data that's not for them and they want to put it back!
        // Might implement this some day
        throw new Error("Not Implemented");
    }

    /** @throws Error always: not implemented. */
    override wrap(): never {
        // not implemented
        throw new Error("Not Implemented");
    }

    /**
     * @throws Error always: not implemented (Readable's iterator would loop forever: read() never returns null).
     */
    override [Symbol.asyncIterator](): never {
        throw new Error("Not Implemented");
    }
}

// Readable's listener methods drive its flow state: on('data') calls resume() (undoing pause()), on('readable') calls
// read(0) on the next tick (which returns, so drains, all the data here), removeListener can resume. StreamBuf emits its
// events itself, as when it did not construct the Duplex state: it uses the plain EventEmitter methods.
const emitterMethods = ["on", "addListener", "removeListener", "off", "removeAllListeners"] as const;
for (const name of emitterMethods) {
    Object.defineProperty(StreamBuf.prototype, name, Object.getOwnPropertyDescriptor(EventEmitter.prototype, name)!);
}

export default StreamBuf;
