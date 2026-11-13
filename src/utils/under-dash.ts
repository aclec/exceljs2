// Small lodash-like helpers. Arrays are iterated with their native methods (holes skipped, the callback gets
// (value, index)); other objects by their own enumerable string keys (Object.keys order), the callback gets (value, key).

import { createDictionary } from "./utils";

const escapeHtmlRegex = /["&<>]/;

/**
 * The callback of {@link each}, {@link some}, {@link every} and {@link map}: an entry of an array and its index, or a
 * property of an object and its key. The value is unknown: an object's own enumerable keys are not limited to the ones
 * its type declares, so these helpers cannot type their values.
 */
export type EntryCallback<R> = (value: unknown, key: number | string) => R;

/**
 * Calls cb for each entry of an array (skipping holes) or each own enumerable property of an object. Does nothing for
 * null or undefined.
 */
export function each(obj: object | null | undefined, cb: EntryCallback<void>): void {
    if (obj) {
        if (Array.isArray(obj)) {
            obj.forEach(cb);
        } else {
            const keys = Object.keys(obj);
            for (let i = 0; i < keys.length; i++) {
                const key = keys[i];
                if (key !== undefined) cb(Reflect.get(obj, key), key);
            }
        }
    }
}

/**
 * Whether cb returns a truthy value for an entry of an array or a property of an object (stops at the first one).
 * False for null or undefined.
 */
export function some(obj: object | null | undefined, cb: EntryCallback<unknown>): boolean {
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.some(cb);
        }
        return Object.keys(obj).some((key) => cb(Reflect.get(obj, key), key));
    }
    return false;
}

/**
 * Whether cb returns a truthy value for every entry of an array or property of an object (stops at the first falsy
 * one). True for null or undefined.
 */
export function every(obj: object | null | undefined, cb: EntryCallback<unknown>): boolean {
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.every(cb);
        }
        return Object.keys(obj).every((key) => cb(Reflect.get(obj, key), key));
    }
    return true;
}

/** The results of cb for each entry of an array or property of an object. An empty array for null or undefined. */
export function map<R>(obj: object | null | undefined, cb: EntryCallback<R>): R[] {
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.map(cb);
        }
        return Object.keys(obj).map((key) => cb(Reflect.get(obj, key), key));
    }
    return [];
}

/**
 * The items of a, by the (string) value of their property p; the last one wins. The result has no prototype: the keys
 * are data, e.g. "__proto__".
 */
export function keyBy<T, K extends keyof T>(a: readonly T[], p: K): Record<string, T> {
    return a.reduce<Record<string, T>>((o, v) => {
        o[String(v[p])] = v;
        return o;
    }, createDictionary<T>());
}

/**
 * Deep equality of plain values: arrays item by item, objects by their own enumerable keys and values, anything else
 * with `===`.
 */
export function isEqual(a: unknown, b: unknown): boolean {
    const aType = typeof a;
    const bType = typeof b;
    const aArray = Array.isArray(a);
    const bArray = Array.isArray(b);

    if (aType !== bType) {
        return false;
    }
    if (typeof a === "object" && typeof b === "object") {
        if (aArray || bArray) {
            if (Array.isArray(a) && Array.isArray(b)) {
                return (
                    a.length === b.length &&
                    a.every((aValue, index) => {
                        const bValue: unknown = b[index];
                        return isEqual(aValue, bValue);
                    })
                );
            }
            return false;
        }

        if (a === null || b === null) {
            return a === b;
        }

        // Compare object keys and values
        const keys = Object.keys(a);

        if (Object.keys(b).length !== keys.length) {
            return false;
        }

        for (const key of keys) {
            if (!Object.hasOwn(b, key)) {
                return false;
            }
        }

        for (const key of keys) {
            if (!isEqual(Reflect.get(a, key), Reflect.get(b, key))) {
                return false;
            }
        }
        return true;
    }
    return a === b;
}

/** Escapes `"`, `&`, `'`, `<` and `>` as HTML entities (a `'` before the first other one is left as is). */
export function escapeHtml(html: string): string {
    const regexResult = escapeHtmlRegex.exec(html);
    if (!regexResult) return html;

    let result = "";
    let escape: string;
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < html.length; i++) {
        switch (html.charAt(i)) {
            case '"':
                escape = "&quot;";
                break;
            case "&":
                escape = "&amp;";
                break;
            case "'":
                escape = "&apos;";
                break;
            case "<":
                escape = "&lt;";
                break;
            case ">":
                escape = "&gt;";
                break;
            default:
                continue;
        }
        if (lastIndex !== i) result += html.substring(lastIndex, i);
        lastIndex = i + 1;
        result += escape;
    }
    if (lastIndex !== i) return result + html.substring(lastIndex, i);
    return result;
}

/** A three-way comparison with `<` and `>`, for Array#sort. */
export function strcmp<T extends string | number>(a: T, b: T): -1 | 0 | 1 {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
}

/** Whether val is undefined. */
export function isUndefined(val: unknown): val is undefined {
    return Object.prototype.toString.call(val) === "[object Undefined]";
}

/**
 * Whether val's toString tag is `[object Object]`: plain and null-prototype objects, class instances; not arrays,
 * dates, null...
 */
export function isObject(val: unknown): val is Record<string, unknown> {
    return Object.prototype.toString.call(val) === "[object Object]";
}

/**
 * Merges the sources into target, deeply: plain objects and arrays are merged into the target's own (created when
 * missing), other values replace the target's; undefined values are skipped, and so are the keys `__proto__`,
 * `constructor` and `prototype`.
 * @param target - The object merged into, and returned; a new object when falsy.
 * @returns target. Its type, the intersection of the arguments' types, is an approximation: a property that a source
 * sets to undefined keeps the target's value.
 */
export function deepMerge<T extends object, A>(target: T, a: A): T & A;
export function deepMerge<T extends object, A, B>(target: T, a: A, b: B): T & A & B;
export function deepMerge<T extends object, A, B, C>(target: T, a: A, b: B, c: C): T & A & B & C;
export function deepMerge(target: object | null | undefined, ...sources: unknown[]): object;
export function deepMerge(target: object | null | undefined, ...sources: unknown[]): object {
    const result: object = target || {};

    function assignValue(val: unknown, key: string | number): void {
        // never merge into Object.prototype via JSON-sourced keys
        if (key === "__proto__" || key === "constructor" || key === "prototype") {
            return;
        }
        if (Array.isArray(val)) {
            const src: unknown = Reflect.get(result, key);
            Reflect.set(result, key, deepMerge(Array.isArray(src) ? src : [], val));
        } else if (isObject(val)) {
            const src: unknown = Reflect.get(result, key);
            Reflect.set(result, key, deepMerge(isObject(src) ? src : {}, val));
        } else if (!isUndefined(val)) {
            Reflect.set(result, key, val);
        }
    }

    function mergeFrom(source: unknown): void {
        if (Array.isArray(source)) {
            source.forEach(assignValue);
        } else if (source) {
            // any other truthy value, by its own enumerable keys: a string is merged by its characters
            for (const [key, val] of Object.entries(source)) assignValue(val, key);
        }
    }

    // the target is merged into itself first, as it always was (its nested objects are merged into themselves)
    mergeFrom(target);
    for (const source of sources) mergeFrom(source);
    return result;
}

const _ = {
    each,
    some,
    every,
    map,
    keyBy,
    isEqual,
    escapeHtml,
    strcmp,
    isUndefined,
    isObject,
    deepMerge,
};

export default _;
