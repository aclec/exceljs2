// Long reads and writes of a workbook give the event loop a turn about every PAUSE_INTERVAL ms, so that timers and I/O
// (e.g. the other requests of a server) are not blocked for the whole operation. The work is split into steps: generators
// that yield between items (rows, chunks); run synchronously by runSteps(), or with pauses by Pacer#run().
// Only work that lasts PAUSE_INTERVAL ms pauses: small workbooks are read and written without waiting for a timer, also
// under the fake timers of test frameworks (which only run setImmediate when the test advances them).

/** The time (ms) a {@link Pacer} lets work run before it gives the event loop a turn. */
const PAUSE_INTERVAL = 25;

/** Work split into steps: an iterator (usually a generator) that yields between items. Yielded values are ignored. */
export type Steps<R = unknown> = Iterator<unknown, R, undefined>;

/**
 * A promise that resolves after a turn of the event loop (timers and I/O run): setImmediate in Node.js and Bun, a
 * message in browsers (setTimeout(0) is clamped to 4 ms or more there).
 */
function nextTurn(): Promise<void> {
    if (typeof setImmediate === "function") {
        return new Promise((resolve) => setImmediate(resolve));
    }
    if (typeof MessageChannel === "function") {
        return new Promise((resolve) => {
            const channel = new MessageChannel();
            channel.port1.onmessage = () => {
                channel.port1.close();
                resolve();
            };
            channel.port2.postMessage(null);
        });
    }
    return new Promise((resolve) => setTimeout(resolve, 0));
}

/** Runs all the steps of a generator at once. */
function runSteps<R>(steps: Steps<R>): R {
    let step = steps.next();
    while (!step.done) step = steps.next();
    return step.value;
}

/** Runs steps of work, giving the event loop a turn every `interval` ms. */
class Pacer {
    /** Runs all the steps of a generator at once. */
    static runSteps = runSteps;
    /** A promise that resolves after a turn of the event loop. */
    static nextTurn = nextTurn;

    /** The time (ms) between two turns of the event loop. */
    interval: number;
    /** When (Date.now()) the event loop last had a turn. */
    last: number;

    constructor(interval = PAUSE_INTERVAL) {
        this.interval = interval;
        this.last = Date.now();
    }

    /**
     * A promise that resolves after a turn of the event loop when interval ms have passed since the last one, else
     * undefined (awaiting it would cost a microtask per step).
     */
    pause(): Promise<void> | undefined {
        if (Date.now() - this.last < this.interval) {
            return undefined;
        }
        return nextTurn().then(() => {
            this.last = Date.now();
        });
    }

    /** Runs the steps of a generator, pausing between them. */
    async run<R>(steps: Steps<R>): Promise<R> {
        let step = steps.next();
        while (!step.done) {
            const pause = this.pause();
            if (pause) {
                await pause;
            }
            step = steps.next();
        }
        return step.value;
    }
}

export { PAUSE_INTERVAL, nextTurn, runSteps };
export default Pacer;
