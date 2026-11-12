// Long reads and writes of a workbook give the event loop a turn about every PAUSE_INTERVAL ms, so that timers and I/O
// (e.g. the other requests of a server) are not blocked for the whole operation. The work is split into steps: generators
// that yield between items (rows, chunks); run synchronously by runSteps(), or with pauses by Pacer#run().

const PAUSE_INTERVAL = 25;

// resolves after a turn of the event loop (timers and I/O run): setImmediate in Node.js and Bun, a message in browsers
// (setTimeout(0) is clamped to 4 ms or more there)
function nextTurn() {
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

// runs all the steps of a generator at once
function runSteps(steps) {
    for (let step = steps.next(); !step.done; step = steps.next());
}

class Pacer {
    constructor(interval = PAUSE_INTERVAL) {
        this.interval = interval;
        this.last = Date.now();
    }

    // a promise that resolves after a turn of the event loop when interval ms have passed since the last one, else
    // undefined (awaiting it would cost a microtask per step)
    pause() {
        if (Date.now() - this.last < this.interval) {
            return undefined;
        }
        return nextTurn().then(() => {
            this.last = Date.now();
        });
    }

    // runs the steps of a generator, pausing between them
    async run(steps) {
        for (let step = steps.next(); !step.done; step = steps.next()) {
            const pause = this.pause();
            if (pause) {
                await pause;
            }
        }
    }
}

Pacer.runSteps = runSteps;
Pacer.nextTurn = nextTurn;

module.exports = Pacer;
