const textDecoder = typeof TextDecoder === "undefined" ? null : new TextDecoder("utf-8");

/**
 * Decodes UTF-8 bytes (a Uint8Array or Buffer) to a string, with TextDecoder (else with the chunk's own toString, which
 * decodes a Buffer). A string is returned as is.
 */
function bufferToString(chunk: string | Uint8Array): string {
    if (typeof chunk === "string") {
        return chunk;
    }
    if (textDecoder) {
        return textDecoder.decode(chunk);
    }
    return chunk.toString();
}

export { bufferToString };
