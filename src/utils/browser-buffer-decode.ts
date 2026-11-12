/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
const textDecoder = typeof TextDecoder === "undefined" ? null : new TextDecoder("utf-8");

function bufferToString(chunk) {
    if (typeof chunk === "string") {
        return chunk;
    }
    if (textDecoder) {
        return textDecoder.decode(chunk);
    }
    return chunk.toString();
}

export { bufferToString };
