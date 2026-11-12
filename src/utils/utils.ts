import * as nodeFs from "fs";

/** A constructor function, as `inherits` sees it: its prototype is replaced. */
interface InheritingConstructor {
    prototype: unknown;
    super_?: unknown;
}

/**
 * Makes cls inherit from superCtor, the ES5 way (prototype-based classes not converted to `class` yet): cls.prototype
 * becomes an object inheriting from superCtor.prototype, with the properties of prototype, and cls gets the properties
 * of statics. `cls.super_` is set to superCtor.
 *
 * Called with three arguments, the third one is the prototype (no statics).
 * @deprecated Use an ES `class ... extends`.
 */
function inherits(
    cls: InheritingConstructor,
    superCtor: { prototype: object | null },
    statics?: object | null,
    prototype?: object | null
): void {
    cls.super_ = superCtor;

    if (!prototype) {
        prototype = statics;
        statics = null;
    }

    if (statics) {
        // the own enumerable string keys, as Object.keys(statics)
        for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(statics))) {
            if (descriptor.enumerable) Object.defineProperty(cls, key, descriptor);
        }
    }

    const properties: PropertyDescriptorMap = {
        constructor: {
            value: cls,
            enumerable: false,
            writable: false,
            configurable: true,
        },
    };
    if (prototype) {
        for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(prototype))) {
            if (descriptor.enumerable) properties[key] = descriptor;
        }
    }

    cls.prototype = Object.create(superCtor.prototype, properties);
}

// eslint-disable-next-line no-control-regex -- the characters that XML 1.0 does not allow
const xmlDecodeRegex = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]/;
// Removes each open...close section (the first close after an open), like
// text.replace(/\[[^\]]*]/g, '') but in linear time: that regular expression rescans the rest of the
// text for every unclosed '['. An open without a close after it is kept.
function removeDelimited(text: string, open: string, close: string): string {
    let result = "";
    let from = 0;
    for (;;) {
        const start = text.indexOf(open, from);
        const end = start === -1 ? -1 : text.indexOf(close, start + 1);
        if (end === -1) {
            return result + text.slice(from);
        }
        result += text.slice(from, start);
        from = end + 1;
    }
}

// the format codes of a number format without its literal text: quoted ("text"), escaped (\m),
// padding (_m) and fill (*m) characters. A quote without its closing quote is kept, like removeDelimited.
function removeLiterals(fmt: string): string {
    let result = "";
    let i = 0;
    while (i < fmt.length) {
        const char = fmt[i];
        const end = char === '"' ? fmt.indexOf('"', i + 1) : -1;
        if (end !== -1) {
            i = end + 1;
        } else if (char === "\\" || char === "_" || char === "*") {
            i += 2;
        } else {
            result += char;
            i++;
        }
    }
    return result;
}

const DATE_FMT_CACHE_SIZE = 1000;
const dateFmtCache = new Map<string, boolean>();

/** Does nothing (a default callback). */
function nop(): void {}

/**
 * A promise that resolves to value after a turn of the event loop (setImmediate, else a 1 ms timer).
 */
function promiseImmediate<T>(value: T): Promise<T> {
    return new Promise((resolve) => {
        if (typeof global.setImmediate === "function") {
            setImmediate(() => {
                resolve(value);
            });
        } else {
            // poorman's setImmediate - must wait at least 1ms
            setTimeout(() => {
                resolve(value);
            }, 1);
        }
    });
}

/**
 * The Excel serial number of a date: days since 1899-12-30 (or since 1904-01-01 with date1904), the time of day as
 * the fraction. The date is read in UTC.
 */
function dateToExcel(d: Date, date1904?: boolean): number {
    return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
}

/** The date of an Excel serial number (see {@link dateToExcel}), rounded to the millisecond, in UTC. */
function excelToDate(v: number, date1904?: boolean): Date {
    const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
    return new Date(millisecondSinceEpoch);
}

/** Splits a zip entry path at its last `/`: `"xl/worksheets/sheet1.xml"` → `{ path: "xl/worksheets", name: "sheet1.xml" }`. */
function parsePath(filepath: string): { path: string; name: string } {
    const last = filepath.lastIndexOf("/");
    return {
        path: filepath.substring(0, last),
        name: filepath.substring(last + 1),
    };
}

/** The path of the relationships part of a part: `"xl/workbook.xml"` → `"xl/_rels/workbook.xml.rels"`. */
function getRelsPath(filepath: string): string {
    const path = parsePath(filepath);
    return `${path.path}/_rels/${path.name}.rels`;
}

/**
 * Escapes text for XML content and attribute values (`&quot;`, `&amp;`, `&apos;`, `&lt;`, `&gt;`), and removes the
 * characters that XML 1.0 does not allow (control characters but tab / line feed / carriage return, DEL, lone
 * surrogates, U+FFFE and U+FFFF). Well-formed surrogate pairs are kept.
 */
function xmlEncode(text: string): string {
    const regexResult = xmlDecodeRegex.exec(text);
    if (!regexResult) return text;

    let result = "";
    let escape: string;
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < text.length; i++) {
        const charCode = text.charCodeAt(i);
        switch (charCode) {
            case 34: // "
                escape = "&quot;";
                break;
            case 38: // &
                escape = "&amp;";
                break;
            case 39: // '
                escape = "&apos;";
                break;
            case 60: // <
                escape = "&lt;";
                break;
            case 62: // >
                escape = "&gt;";
                break;
            case 127:
            case 0xfffe: // not allowed in XML 1.0
            case 0xffff:
                escape = "";
                break;
            default: {
                if (charCode >= 0xd800 && charCode <= 0xdfff) {
                    const next = text.charCodeAt(i + 1);
                    if (charCode <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
                        i++; // well-formed surrogate pair (e.g. emoji): keep both halves
                        continue;
                    }
                    escape = ""; // lone surrogate: not allowed in XML 1.0
                    break;
                }
                if (charCode <= 31 && (charCode <= 8 || (charCode >= 11 && charCode !== 13))) {
                    escape = "";
                    break;
                }
                continue;
            }
        }
        if (lastIndex !== i) result += text.substring(lastIndex, i);
        lastIndex = i + 1;
        if (escape) result += escape;
    }
    if (lastIndex !== i) return result + text.substring(lastIndex, i);
    return result;
}

/** Decodes the five predefined XML entities (`&lt;` `&gt;` `&amp;` `&apos;` `&quot;`); other entities are kept. */
function xmlDecode(text: string): string {
    return text.replace(/&([a-z]*);/g, (c) => {
        switch (c) {
            case "&lt;":
                return "<";
            case "&gt;":
                return ">";
            case "&amp;":
                return "&";
            case "&apos;":
                return "'";
            case "&quot;":
                return '"';
            default:
                return c;
        }
    });
}

/** The integer that value starts with (parseInt, base 10), or 0 when there is none. */
function validInt(value: string | number | null | undefined): number {
    const i = parseInt(String(value), 10);
    return !Number.isNaN(i) ? i : 0;
}

/**
 * Whether a number format formats dates or times: it has date / time codes (y m d h s b) outside its literal text,
 * colours, conditions and locales. Runs for every numeric cell with a custom format: linear, and cached per format.
 */
function isDateFmt(fmt: string | null | undefined): boolean {
    if (!fmt) {
        return false;
    }
    let result = dateFmtCache.get(fmt);
    if (result === undefined) {
        // ignore the chars inside [] (colours, conditions, locales), then the literal ones
        const text = removeLiterals(removeDelimited(fmt, "[", "]"));
        // then check for date formatting chars
        result = /[ymdhMsb]/.test(text);
        if (dateFmtCache.size >= DATE_FMT_CACHE_SIZE) {
            dateFmtCache.clear();
        }
        dateFmtCache.set(fmt, result);
    }
    return result;
}

/** Whether a file exists (Node.js and Bun only). */
function exists(path: string): Promise<boolean> {
    return new Promise((resolve) => {
        nodeFs.access(path, nodeFs.constants.F_OK, (err) => {
            resolve(!err);
        });
    });
}

/** File system helpers (Node.js and Bun only). */
const fs = { exists };

/**
 * The date part of the ISO 8601 form of dt, in UTC: `"2024-01-31"`. (Upstream called `toIsoString().subsstr()`,
 * which do not exist, so it always threw a TypeError.)
 * @deprecated Use `date.toISOString().substring(0, 10)`.
 */
function toIsoDateString(dt: Date): string {
    return dt.toISOString().substring(0, 10);
}

/**
 * Reads an xsd:boolean: true for `true`, `"true"`, `1` and `"1"`, false for anything else.
 * @param absent - The value of a missing attribute (null or undefined): the schema default.
 */
function parseBoolean(value: string | number | boolean | null | undefined, absent = false): boolean {
    if (value === undefined || value === null) {
        return absent;
    }
    return value === true || value === "true" || value === 1 || value === "1";
}

/** The numbers from start (included) to stop (excluded), by step (negative to count down). */
function* range(start: number, stop: number, step = 1): Generator<number, void, undefined> {
    const compareOrder = step > 0 ? (a: number, b: number) => a < b : (a: number, b: number) => a > b;
    for (let value = start; compareOrder(value, stop); value += step) {
        yield value;
    }
}

const isFiniteNumber = <T>(item: T): item is T & number => Number.isFinite(item);

/** The values, sorted: numerically when they are all finite numbers, else as strings (Array#sort). */
function toSortedArray<T>(values: Iterable<T> | ArrayLike<T>): T[] {
    const result = Array.from(values);

    // Note: per default, `Array.prototype.sort()` converts values
    // to strings when comparing. Here, if we have numbers, we use
    // numeric sort.
    if (result.every(isFiniteNumber)) {
        const compareNumbers = (a: number, b: number) => a - b;
        return result.sort(compareNumbers);
    }

    return result.sort();
}

/**
 * An object with each of props set to value (null by default). It has no prototype: the props are header names from
 * the data, such as "constructor" or "__proto__".
 */
function objectFromProps<V = null>(props: readonly string[], value: V | null = null): Record<string, V | null> {
    // reduce, not for...of: holes of a sparse props array are skipped
    return props.reduce<Record<string, V | null>>((result, property) => {
        result[property] = value;
        return result;
    }, createDictionary<V | null>());
}

/**
 * A type guard for a union of string literals: `isOneOf(["a", "b"] as const)(x)` is true when x is `"a"` or `"b"`
 * (e.g. an attribute value read from a file).
 */
function isOneOf<T extends string>(values: readonly T[]): (value: unknown) => value is T {
    const set = new Set<unknown>(values);
    return (value: unknown): value is T => set.has(value);
}

/**
 * A new empty object without prototype, for keys that come from data (`"__proto__"`, `"constructor"`...), which a
 * plain `{}` would resolve on Object.prototype.
 */
function createDictionary<V>(): Record<string, V> {
    const dictionary: Record<string, V> = {};
    Object.setPrototypeOf(dictionary, null);
    return dictionary;
}

const utils = {
    nop,
    promiseImmediate,
    inherits,
    dateToExcel,
    excelToDate,
    parsePath,
    getRelsPath,
    xmlEncode,
    xmlDecode,
    validInt,
    isDateFmt,
    fs,
    toIsoDateString,
    parseBoolean,
    range,
    toSortedArray,
    objectFromProps,
    createDictionary,
    isOneOf,
};

export {
    nop,
    promiseImmediate,
    inherits,
    dateToExcel,
    excelToDate,
    parsePath,
    getRelsPath,
    xmlEncode,
    xmlDecode,
    validInt,
    isDateFmt,
    fs,
    exists,
    toIsoDateString,
    parseBoolean,
    range,
    toSortedArray,
    objectFromProps,
    createDictionary,
    isOneOf,
};

export default utils;
