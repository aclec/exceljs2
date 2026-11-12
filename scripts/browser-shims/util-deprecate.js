"use strict";

// Browser replacement for util-deprecate (used by readable-stream): warns once when the deprecated function is called,
// without reading its flags from globalThis.localStorage, which Node.js 25+ warns about (ExperimentalWarning) as soon as
// the bundle is loaded there. Only bundled into dist/.
module.exports = function deprecate(fn, message) {
    let warned = false;
    return function deprecated(...args) {
        if (!warned) {
            warned = true;
            console.warn(message);
        }
        return fn.apply(this, args);
    };
};
