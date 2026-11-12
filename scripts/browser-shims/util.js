"use strict";

// Browser replacement for the subset of node:util used by the bundled dependencies (readable-stream: debuglog, inspect;
// @fast-csv/format: promisify). Only bundled into dist/; Node and Bun keep the native module. Smaller than the util
// polyfill, and it does not reach for globalThis.localStorage (see util-deprecate.js).
function promisify(fn) {
    return function promisified(...args) {
        return new Promise((resolve, reject) => {
            fn.call(this, ...args, (error, value) => (error ? reject(error) : resolve(value)));
        });
    };
}

// readable-stream only calls it from its own util.inspect.custom method, which browsers never call
function inspect(value) {
    return String(value);
}
inspect.custom = Symbol.for("nodejs.util.inspect.custom");

// NODE_DEBUG logging is not available in browsers
function debuglog() {
    return () => {};
}

module.exports = { promisify, inspect, debuglog };
