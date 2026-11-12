const { SaxesParser } = require("saxes");
const { PassThrough } = require("stream");

// saxes keeps an object per open element: a document nested deeper than this fails instead of exhausting memory (the
// parts of an xlsx file are nested less than 30 deep)
const MAX_DEPTH = 1000;

module.exports = async function* (iterable) {
    // TODO: Remove once node v8 is deprecated
    // Detect and upgrade old streams
    if (iterable.pipe && !iterable[Symbol.asyncIterator]) {
        iterable = iterable.pipe(new PassThrough());
    }
    const saxesParser = new SaxesParser();
    let error;
    saxesParser.on("error", (err) => {
        error = err;
    });
    let events = [];
    let depth = 0;
    saxesParser.on("opentag", (value) => {
        if (++depth > MAX_DEPTH) {
            // thrown out of saxesParser.write: stops the parse at once
            throw new Error(`XML elements are nested more than ${MAX_DEPTH} deep`);
        }
        events.push({ eventType: "opentag", value });
    });
    saxesParser.on("text", (value) => events.push({ eventType: "text", value }));
    saxesParser.on("closetag", (value) => {
        // self-closing elements too
        depth--;
        events.push({ eventType: "closetag", value });
    });
    // one decoder per parse: in streaming mode it keeps the bytes of a multi-byte character
    // split across two chunks, so it must not be shared between concurrent parses

    const decoder = new TextDecoder("utf-8");
    for await (const chunk of iterable) {
        saxesParser.write(typeof chunk === "string" ? chunk : decoder.decode(chunk, { stream: true }));
        // saxesParser.write and saxesParser.on() are synchronous,
        // so we can only reach the below line once all events have been emitted
        if (error) throw error;
        // As a performance optimization, we gather all events instead of passing
        // them one by one, which would cause each event to go through the event queue
        yield events;
        events = [];
    }
    // flush an incomplete trailing sequence (decoded as U+FFFD)
    const tail = decoder.decode();
    if (tail) {
        saxesParser.write(tail);
        if (error) throw error;
        yield events;
    }
};
