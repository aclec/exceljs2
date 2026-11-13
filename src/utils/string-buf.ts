/** Options of a {@link StringBuf}. */
export interface StringBufOptions {
    /** The initial capacity, in bytes (16384 by default); it doubles when needed. */
    size?: number;
    /** The encoding of the text added (utf8 by default). */
    encoding?: BufferEncoding;
}

/**
 * A growable byte buffer to build the XML of a part with few string operations: text is encoded into it as it is added.
 */
class StringBuf {
    private _buf: Buffer;
    private _encoding: BufferEncoding;
    // where in the buffer we are at
    private _inPos: number;
    // the copy that toBuffer() returned, until the content changes
    private _buffer: Buffer | undefined;

    constructor(options?: StringBufOptions) {
        this._buf = Buffer.alloc((options && options.size) || 16384);
        this._encoding = (options && options.encoding) || "utf8";

        // where in the buffer we are at
        this._inPos = 0;

        // for use by toBuffer()
        this._buffer = undefined;
    }

    /** The number of bytes written. */
    get length(): number {
        return this._inPos;
    }

    /** The size of the underlying buffer, in bytes. */
    get capacity(): number {
        return this._buf.length;
    }

    /** The underlying buffer: its first {@link length} bytes are the content. */
    get buffer(): Buffer {
        return this._buf;
    }

    /** A copy of the content, kept until the content changes. */
    toBuffer(): Buffer {
        // return the current data as a single enclosing buffer
        if (!this._buffer) {
            // not zero-filled: the copy overwrites every byte (2 to 3 times faster than Buffer.alloc)
            this._buffer = Buffer.allocUnsafe(this.length);
            this._buf.copy(this._buffer, 0, 0, this.length);
        }
        return this._buffer;
    }

    /** Moves the write position back to position (0 by default: empties the buffer). */
    reset(position?: number): void {
        position = position || 0;
        this._buffer = undefined;
        this._inPos = position;
    }

    // doubles the capacity until it is at least min
    private _grow(min: number): void {
        let size = this._buf.length * 2;
        while (size < min) {
            size *= 2;
        }
        const buf = Buffer.alloc(size);
        this._buf.copy(buf, 0);
        this._buf = buf;
    }

    /** Appends text, encoded. */
    addText(text: string): void {
        this._buffer = undefined;

        let inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);

        // if we've hit (or nearing capacity), grow the buf
        while (inPos >= this._buf.length - 4) {
            this._grow(this._inPos + text.length);

            // keep trying to write until we've completely written the text
            inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);
        }

        this._inPos = inPos;
    }

    /** Appends the content of another StringBuf. */
    addStringBuf(inBuf: StringBuf): void {
        if (inBuf.length) {
            this._buffer = undefined;

            if (this.length + inBuf.length > this.capacity) {
                this._grow(this.length + inBuf.length);
            }

            inBuf._buf.copy(this._buf, this._inPos, 0, inBuf.length);
            this._inPos += inBuf.length;
        }
    }
}

export default StringBuf;
