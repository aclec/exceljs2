import { EventEmitter } from "events";

/** Options of a {@link LineBuffer}. */
export interface LineBufferOptions {
    /** Kept for reference: the chunks written are strings, already decoded. */
    encoding?: BufferEncoding;
}

/** The events of a {@link LineBuffer}. */
interface LineBufferEvents {
    /** a complete line (without its end of line) */
    line: [line: string];
    /** every line was emitted (after {@link LineBuffer.end}) */
    done: [];
    /** the buffer was uncorked: the source can write again */
    drain: [];
}

/**
 * Splits text written in chunks into lines (`\n` or `\r\n`), emitted as `line` events; while corked, the lines are
 * queued and emitted on {@link uncork}. Not used by the library itself (kept for compatibility of deep imports).
 */
class LineBuffer extends EventEmitter<LineBufferEvents> {
    /** The encoding given in the options (unused). */
    encoding: BufferEncoding | undefined;

    /** The incomplete last line, until more text or {@link end} completes it. */
    buffer: string | null;

    /** While true, the lines are queued instead of emitted. */
    corked: boolean;

    /** The lines read while corked. */
    queue: string[];

    constructor(options: LineBufferOptions) {
        super();

        this.encoding = options.encoding;

        this.buffer = null;

        // part of cork/uncork
        this.corked = false;
        this.queue = [];
    }

    /**
     * Adds text: emits (or queues, when corked) the lines it completes.
     *
     * @returns `false` when corked (the caller should wait for `drain`)
     */
    write(chunk: string): boolean {
        const data = this.buffer ? this.buffer + chunk : chunk;
        const lines = data.split(/\r?\n/g);

        // save the last line
        this.buffer = lines.pop() ?? null;

        for (const line of lines) {
            if (this.corked) {
                this.queue.push(line);
            } else {
                this.emit("line", line);
            }
        }

        return !this.corked;
    }

    /** Queues the lines instead of emitting them. */
    cork(): void {
        this.corked = true;
    }

    /** Emits the queued lines, then `drain`. */
    uncork(): void {
        this.corked = false;
        this._flush();

        // tell the source I'm ready again
        this.emit("drain");
    }

    /** No-op (stream interface). */
    setDefaultEncoding(): void {
        // ?
    }

    /** Emits the last line (if not empty), then `done`. */
    end(): void {
        if (this.buffer) {
            this.emit("line", this.buffer);
            this.buffer = null;
        }
        this.emit("done");
    }

    private _flush(): void {
        if (!this.corked) {
            for (const line of this.queue) {
                this.emit("line", line);
            }
            this.queue = [];
        }
    }
}

export default LineBuffer;
