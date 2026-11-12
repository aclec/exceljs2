// Small lodash-like helpers. Arrays are iterated with their native methods (holes skipped, the callback gets
// (value, index)); other objects by their own enumerable string keys (Object.keys order), the callback gets (value, key).

const escapeHtmlRegex = /["&<>]/;

/** The values of the own enumerable string-keyed properties of O. */
type ValueOf<O> = O[keyof O & string];
/** The own enumerable string keys of O. */
type KeyOf<O> = keyof O & string;

/** A loose view of any object, to read its properties by key. */
type Indexable = Record<PropertyKey, unknown>;

/**
 * The callback parameter of the implementation signatures below: the overloads give their callbacks precise parameter
 * types, so the implementation's must accept any of them.
 */
type AnyCallback = (value: never, key: never) => unknown;
/** How the implementations call the callbacks: with the values and keys of whatever they iterate. */
type LooseCallback<R> = (value: unknown, key: string | number) => R;

/**
 * Calls cb for each entry of an array (skipping holes) or each own enumerable property of an object. Does nothing for
 * null or undefined.
 */
export function each<T>(obj: readonly T[] | null | undefined, cb: (value: T, index: number) => void): void;
export function each<O extends object>(obj: O | null | undefined, cb: (value: ValueOf<O>, key: KeyOf<O>) => void): void;
export function each(obj: object | null | undefined, callback: AnyCallback): void {
    const cb = callback as LooseCallback<void>;
    if (obj) {
        if (Array.isArray(obj)) {
            obj.forEach(cb);
        } else {
            // an index loop: 10 to 30% faster than Object.keys().forEach (attributes of every XML element written)
            const keys = Object.keys(obj);
            for (let i = 0; i < keys.length; i++) {
                const key = keys[i]!;
                cb((obj as Indexable)[key], key);
            }
        }
    }
}

/**
 * Whether cb returns a truthy value for an entry of an array or a property of an object (stops at the first one).
 * False for null or undefined.
 */
export function some<T>(obj: readonly T[] | null | undefined, cb: (value: T, index: number) => unknown): boolean;
export function some<O extends object>(obj: O | null | undefined, cb: (value: ValueOf<O>, key: KeyOf<O>) => unknown): boolean;
export function some(obj: object | null | undefined, callback: AnyCallback): boolean {
    const cb = callback as LooseCallback<unknown>;
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.some(cb);
        }
        return Object.keys(obj).some((key) => cb((obj as Indexable)[key], key));
    }
    return false;
}

/**
 * Whether cb returns a truthy value for every entry of an array or property of an object (stops at the first falsy
 * one). True for null or undefined.
 */
export function every<T>(obj: readonly T[] | null | undefined, cb: (value: T, index: number) => unknown): boolean;
export function every<O extends object>(obj: O | null | undefined, cb: (value: ValueOf<O>, key: KeyOf<O>) => unknown): boolean;
export function every(obj: object | null | undefined, callback: AnyCallback): boolean {
    const cb = callback as LooseCallback<unknown>;
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.every(cb);
        }
        return Object.keys(obj).every((key) => cb((obj as Indexable)[key], key));
    }
    return true;
}

/** The results of cb for each entry of an array or property of an object. An empty array for null or undefined. */
export function map<T, R>(obj: readonly T[] | null | undefined, cb: (value: T, index: number) => R): R[];
export function map<O extends object, R>(obj: O | null | undefined, cb: (value: ValueOf<O>, key: KeyOf<O>) => R): R[];
export function map(obj: object | null | undefined, callback: AnyCallback): unknown[] {
    const cb = callback as LooseCallback<unknown>;
    if (obj) {
        if (Array.isArray(obj)) {
            return obj.map(cb);
        }
        return Object.keys(obj).map((key) => cb((obj as Indexable)[key], key));
    }
    return [];
}

/**
 * The items of a, by the (string) value of their property p; the last one wins. The result has no prototype: the keys
 * are data, e.g. "__proto__".
 */
export function keyBy<T, K extends keyof T>(a: readonly T[], p: K): Record<string, T> {
    return a.reduce<Record<string, T>>(
        (o, v) => {
            o[String(v[p])] = v;
            return o;
        },
        Object.create(null) as Record<string, T>
    );
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
            if (!isEqual((a as Indexable)[key], (b as Indexable)[key])) {
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

/** Whether val is a plain object (`[object Object]`: no array, date, null...). */
export function isObject(val: unknown): val is Record<string, unknown> {
    return Object.prototype.toString.call(val) === "[object Object]";
}

/**
 * Merges the sources into target, deeply: plain objects and arrays are merged into the target's own (created when
 * missing), other values replace the target's; undefined values are skipped, and so are the keys `__proto__`,
 * `constructor` and `prototype`.
 * @param target - The object merged into, and returned; a new object when falsy.
 * @returns target.
 */
export function deepMerge<T extends object, A>(target: T, a: A): T & A;
export function deepMerge<T extends object, A, B>(target: T, a: A, b: B): T & A & B;
export function deepMerge<T extends object, A, B, C>(target: T, a: A, b: B, c: C): T & A & B & C;
export function deepMerge(target: object | null | undefined, ...sources: unknown[]): object;
export function deepMerge(...args: unknown[]): object {
    const target = (args[0] || {}) as Indexable;
    const { length } = args;

    let src: unknown;
    let clone: unknown;
    let copyIsArray: boolean;

    function assignValue(val: unknown, key: string | number): void {
        // never merge into Object.prototype via JSON-sourced keys
        if (key === "__proto__" || key === "constructor" || key === "prototype") {
            return;
        }
        src = target[key];
        copyIsArray = Array.isArray(val);
        if (isObject(val) || copyIsArray) {
            if (copyIsArray) {
                copyIsArray = false;
                clone = src && Array.isArray(src) ? src : [];
            } else {
                clone = src && isObject(src) ? src : {};
            }
            target[key] = deepMerge(clone as object, val);
        } else if (!isUndefined(val)) {
            target[key] = val;
        }
    }

    for (let i = 0; i < length; i++) {
        // any value: a string is merged by its characters, as Object.keys sees it
        each(args[i] as object | null | undefined, assignValue);
    }
    return target;
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
