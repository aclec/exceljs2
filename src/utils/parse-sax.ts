import { SaxesParser, type SaxesTagPlain } from "saxes";
import { PassThrough } from "stream";

// saxes keeps an object per open element: a document nested deeper than this fails instead of exhausting memory (the
// parts of an xlsx file are nested less than 30 deep)
const MAX_DEPTH = 1000;

/**
 * An XML element as reported by the SAX parser (start and end tags, also of self-closing elements).
 *
 * `name` is the qualified name as written (`x14:cfRule`), namespaces are not resolved. Attribute values are decoded strings
 * from the file: untrusted data. An attribute that is absent is `undefined` (with `noUncheckedIndexedAccess`).
 */
export type SaxTag = SaxesTagPlain;

/** One SAX event: an element start, a chunk of text (entities decoded; a text node may come in several chunks), or an element end. */
export type SaxEvent =
    | { readonly eventType: "opentag"; readonly value: SaxTag }
    | { readonly eventType: "text"; readonly value: string }
    | { readonly eventType: "closetag"; readonly value: SaxTag };

/** A chunk of an XML document: text, or UTF-8 bytes (a multi-byte character may be split across two chunks). */
export type SaxChunk = string | ArrayBuffer | ArrayBufferView;

/** A pre-async-iterator Node.js stream (`pipe()` only): read through a PassThrough. */
interface LegacyReadable {
    pipe(destination: PassThrough): PassThrough;
}

/**
 * An XML document, chunk by chunk: an (async) iterable such as a Node.js Readable, a zip entry stream, an async generator,
 * or an array of chunks.
 */
export type SaxInput = AsyncIterable<SaxChunk> | Iterable<SaxChunk> | LegacyReadable;

// property reads rather than `in`, which throws on a string (a string is an Iterable<string>)
function isLegacyReadable(input: SaxInput): input is LegacyReadable {
    const candidate = input as Partial<LegacyReadable & AsyncIterable<SaxChunk>>;
    return Boolean(candidate.pipe) && !candidate[Symbol.asyncIterator];
}

/**
 * Parses an XML document with saxes, yielding its events in batches: one array per input chunk (iterating the events of a
 * batch synchronously is much cheaper than one await per event).
 *
 * Protections against hostile documents: an element nested more than 1000 deep fails at once; a document that ends early
 * (an element left open, no root element) fails instead of being read as complete; the first well-formedness error fails.
 * Bytes are decoded with one streaming UTF-8 TextDecoder per parse, which keeps a multi-byte character split across two
 * chunks (an incomplete trailing sequence decodes as U+FFFD).
 *
 * @param iterable - the document; byte chunks must be UTF-8
 * @returns batches of events; each array is new (the consumer may keep it)
 * @throws Error on malformed XML (the first saxes error), on nesting deeper than 1000 elements, or on a truncated document
 */
export default async function* parseSax(iterable: SaxInput): AsyncGenerator<SaxEvent[], void, undefined> {
    // a stream without Symbol.asyncIterator (old readable-stream versions) is read through a PassThrough, which has one
    const input: AsyncIterable<SaxChunk> | Iterable<SaxChunk> = isLegacyReadable(iterable) ? iterable.pipe(new PassThrough()) : iterable;
    const saxesParser = new SaxesParser();
    let error: Error | undefined;
    saxesParser.on("error", (err) => {
        // the first error: saxes goes on after it, and the next ones follow from it
        error = error || err;
    });
    let events: SaxEvent[] = [];
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
    for await (const chunk of input) {
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
        events = [];
    }
    // the end of the document: a truncated one (an element left open, or no root element) fails instead of being read as
    // if it were complete
    saxesParser.close();
    if (error) throw error;
    if (events.length) {
        yield events;
    }
}
