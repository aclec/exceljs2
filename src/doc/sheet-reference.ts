// The cell addresses and ranges that a worksheet takes (getCell, mergeCells, fillFormula...), decoded strictly: a
// reference that names a sheet must name the worksheet itself, as a merge or a formula range written with another
// sheet name (`<mergeCell ref="Sheet2!A1:B2">`) makes an invalid file.

import colCache, { type CellAddress } from "../utils/col-cache";
import Range, { type RangeArgs } from "./range";

// character codes of a cell address
const DOLLAR = 36;
const ZERO = 48;
const NINE = 57;

/** Whether a character code is an upper-case letter A-Z (a column letter). */
function isLetter(code: number): boolean {
    return code >= 65 && code <= 90;
}

/** The bounds of a reference (1-based, inclusive, normalised: top <= bottom, left <= right) and its sheet name. */
export interface SheetReference {
    /** The sheet name, unquoted; `undefined` when the reference has none. */
    readonly sheetName: string | undefined;
    readonly top: number;
    readonly left: number;
    readonly bottom: number;
    readonly right: number;
}

/** The arguments of a range that a worksheet takes: {@link RangeArgs} but nothing (an empty range is no range). */
export type SheetRangeArgs = Exclude<RangeArgs, []>;

/** Whether `n` is a row number of a worksheet: an integer from 1 to `colCache.MAX_ROWS`. */
export function isRowNumber(n: unknown): n is number {
    return typeof n === "number" && Number.isInteger(n) && n >= 1 && n <= colCache.MAX_ROWS;
}

/** Whether `n` is a column number of a worksheet: an integer from 1 to `colCache.MAX_COLS`. */
export function isColumnNumber(n: unknown): n is number {
    return typeof n === "number" && Number.isInteger(n) && n >= 1 && n <= colCache.MAX_COLS;
}

/**
 * A cell address (`"B3"`, `"$B$3"`, `"B$3"`, `"B03"`), without sheet name, decoded as its canonical form (`"B3"`).
 * @returns `undefined` for anything else, a row out of 1..1,048,576 or a column beyond `XFD`.
 */
function decodeCell(text: string): CellAddress | undefined {
    // a scan of the characters (getCell's hot path): [$] 1 to 3 letters A-Z, [$] 1 to 7 digits, nothing else
    const { length } = text;
    let i = text.charCodeAt(0) === DOLLAR ? 1 : 0;
    const lettersStart = i;
    while (i < length && i - lettersStart < 4 && isLetter(text.charCodeAt(i))) {
        i++;
    }
    const lettersEnd = i;
    if (lettersEnd === lettersStart || lettersEnd - lettersStart > 3) {
        return undefined;
    }
    if (text.charCodeAt(i) === DOLLAR) {
        i++;
    }
    const digitsStart = i;
    let row = 0;
    for (; i < length; i++) {
        const code = text.charCodeAt(i);
        if (code < ZERO || code > NINE || i - digitsStart >= 7) {
            return undefined;
        }
        row = row * 10 + code - ZERO;
    }
    if (i === digitsStart || !isRowNumber(row)) {
        return undefined;
    }
    const letters = text.slice(lettersStart, lettersEnd);
    // same length, upper case: the string order is the column order (colCache.l2n throws beyond XFD)
    if (letters.length === 3 && letters > "XFD") {
        return undefined;
    }
    // the canonical form (the common case) is decoded by colCache directly: cached for the top-left cells
    return lettersStart === 0 && digitsStart === lettersEnd && text.charCodeAt(digitsStart) !== ZERO
        ? colCache.decodeAddress(text)
        : colCache.getAddress(row, colCache.l2n(letters));
}

/**
 * The sheet name of a reference prefix (before `!`): `Sheet1`, or quoted `'My sheet'` (`''` for a quote).
 * @returns `undefined` for an empty name or a misplaced quote.
 */
function decodeSheetName(prefix: string): string | undefined {
    if (prefix.length > 2 && prefix.startsWith("'") && prefix.endsWith("'")) {
        const quoted = prefix.slice(1, -1);
        // every quote of a quoted name is doubled
        return quoted.replaceAll("''", "").includes("'") ? undefined : quoted.replaceAll("''", "'");
    }
    return prefix && !prefix.includes("'") ? prefix : undefined;
}

/**
 * Decodes a reference: a cell or a range (`A1`, `$A$1:$B$2`, the corners in any order), optionally preceded by a sheet
 * name (`Sheet1!A1:B2`, `'My sheet'!A1`). Strict: upper-case column letters, rows from 1 to 1,048,576, columns from
 * `A` to `XFD`, nothing else in the text. Linear in the length of the text.
 * @returns `undefined` for anything else
 */
export function decodeReference(text: string): SheetReference | undefined {
    // a sheet name may contain "!" (quoted), an address never does
    const bang = text.lastIndexOf("!");
    const sheetName = bang === -1 ? undefined : decodeSheetName(text.slice(0, bang));
    if (bang !== -1 && sheetName === undefined) {
        return undefined;
    }
    const address = text.slice(bang + 1);
    const colon = address.indexOf(":");
    const tl = decodeCell(colon === -1 ? address : address.slice(0, colon));
    const br = colon === -1 ? tl : decodeCell(address.slice(colon + 1));
    if (!tl || !br) {
        return undefined;
    }
    return {
        sheetName,
        top: Math.min(tl.row, br.row),
        left: Math.min(tl.col, br.col),
        bottom: Math.max(tl.row, br.row),
        right: Math.max(tl.col, br.col),
    };
}

/** Whether a sheet name designates the worksheet `worksheetName` (sheet names are case-insensitive, as in Excel). */
function isSheet(sheetName: string, worksheetName: string): boolean {
    return sheetName.toLowerCase() === worksheetName.toLowerCase();
}

/** The arguments as they were given, for an error message. */
function describeArgs(args: readonly unknown[], depth = 0): string {
    return args.map((arg) => describeArg(arg, depth)).join(", ");
}

/** The longest text of an argument quoted in an error message. */
const MAX_QUOTED = 64;

/** The location keys of an object argument shown in an error message. */
const LOCATION_KEYS = ["top", "left", "bottom", "right", "sheetName"] as const;

/**
 * One argument of a range as an error message shows it: never throws (a circular object or array is fine), and stays
 * short (a long text is cut, nested arrays are shown to a depth of 2, an object by its location keys).
 */
function describeArg(arg: unknown, depth: number): string {
    if (arg instanceof Range) {
        return arg.range;
    }
    if (Array.isArray(arg)) {
        return depth < 2 ? `[${describeArgs(arg, depth + 1)}]` : "[...]";
    }
    if (typeof arg === "string") {
        return JSON.stringify(arg.slice(0, MAX_QUOTED));
    }
    if (typeof arg === "object" && arg !== null) {
        const keys = LOCATION_KEYS.filter((key) => key in arg);
        const fields = keys.map((key) => {
            const value: unknown = Reflect.get(arg, key);
            return `${key}: ${describeValue(value)}`;
        });
        return keys.length ? `{ ${fields.join(", ")} }` : "an object";
    }
    return describeValue(arg);
}

/** A value that is not an object as an error message shows it (an object as its type: its `toString` may throw). */
function describeValue(value: unknown): string {
    switch (typeof value) {
        case "string":
            return JSON.stringify(value.slice(0, MAX_QUOTED));
        case "number":
        case "boolean":
        case "bigint":
        case "undefined":
            return String(value);
        case "object":
            return value === null ? "null" : "an object";
        case "symbol":
        case "function":
            return `a ${typeof value}`;
    }
}

/** A cell address of a worksheet, or why the text or the numbers are not one. */
type CellLookup = { readonly address: CellAddress; readonly error?: undefined } | { readonly address?: undefined; readonly error: string };

/**
 * The address of a cell of the worksheet `worksheetName`, given by its row and column numbers or by its address
 * (`"B2"`, `"$B$2"`, `"Sheet1!B2"`: a sheet name must be the worksheet's).
 */
function lookupCell(worksheetName: string, r: number | string, c: number | undefined): CellLookup {
    if (typeof r === "number" || c !== undefined) {
        if (isRowNumber(r) && isColumnNumber(c)) {
            return { address: colCache.getAddress(r, c) };
        }
        return {
            error: `Invalid cell (${describeArgs([r, c])}): the row must be an integer from 1 to ${colCache.MAX_ROWS} and the column from 1 to ${colCache.MAX_COLS}`,
        };
    }
    // the common case: an address without sheet name
    const cell = decodeCell(r);
    if (cell) {
        return { address: cell };
    }
    const reference = decodeReference(r);
    if (!reference || reference.top !== reference.bottom || reference.left !== reference.right) {
        return { error: `Invalid cell address: ${JSON.stringify(r)}` };
    }
    if (reference.sheetName !== undefined && !isSheet(reference.sheetName, worksheetName)) {
        return { error: `The cell ${r} is not on the worksheet ${JSON.stringify(worksheetName)}` };
    }
    return { address: colCache.getAddress(reference.top, reference.left) };
}

/**
 * The address of a cell of the worksheet `worksheetName` (see {@link lookupCell}).
 * @throws Error for an invalid address (row < 1 included), or a cell of another sheet
 */
export function resolveCellAddress(worksheetName: string, r: number | string, c?: number): CellAddress {
    const lookup = lookupCell(worksheetName, r, c);
    if (lookup.error !== undefined) {
        throw new Error(lookup.error);
    }
    return lookup.address;
}

/**
 * The address of a cell of the worksheet `worksheetName` (see {@link lookupCell}).
 * @returns `undefined` for an invalid address, or a cell of another sheet
 */
export function findCellAddress(worksheetName: string, r: number | string, c?: number): CellAddress | undefined {
    return lookupCell(worksheetName, r, c).address;
}

/** The deepest nesting of range arguments given as one array (`mergeCells([["A1:B2"]])`): a deeper one is no range. */
const MAX_ARGS_NESTING = 4;

/**
 * The bounds and the sheet names (`undefined` when not given) of range arguments; `undefined` when they are not a range
 * (also when nested deeper than {@link MAX_ARGS_NESTING}: a circular array is no range).
 */
function decodeRangeArgs(args: readonly unknown[], depth = 0): { bounds: SheetReference; sheetNames: (string | undefined)[] } | undefined {
    const [first, second, third, fourth, fifth] = args;
    switch (args.length) {
        case 1: {
            if (Array.isArray(first)) {
                return depth < MAX_ARGS_NESTING ? decodeRangeArgs(first, depth + 1) : undefined;
            }
            if (typeof first === "string") {
                const bounds = decodeReference(first);
                return bounds && { bounds, sheetNames: [bounds.sheetName] };
            }
            if (typeof first !== "object" || first === null) {
                return undefined;
            }
            // a Range (its model may hold the 0 or undefined bounds of an empty or error range), or a location
            const location: Partial<Record<"top" | "left" | "bottom" | "right" | "sheetName", unknown>> =
                first instanceof Range ? first.model : first;
            const { top, left, bottom, right, sheetName } = location;
            if (!isRowNumber(top) || !isRowNumber(bottom) || !isColumnNumber(left) || !isColumnNumber(right)) {
                return undefined;
            }
            if (sheetName !== undefined && typeof sheetName !== "string") {
                return undefined;
            }
            const bounds = {
                sheetName,
                top: Math.min(top, bottom),
                left: Math.min(left, right),
                bottom: Math.max(top, bottom),
                right: Math.max(left, right),
            };
            return { bounds, sheetNames: [sheetName] };
        }
        case 2:
        case 3: {
            // two cells, each with an optional sheet name, then an optional sheet name
            if (typeof first !== "string" || typeof second !== "string" || (third !== undefined && typeof third !== "string")) {
                return undefined;
            }
            const tl = decodeReference(first);
            const br = decodeReference(second);
            if (!tl || !br || tl.top !== tl.bottom || tl.left !== tl.right || br.top !== br.bottom || br.left !== br.right) {
                return undefined;
            }
            const bounds = {
                sheetName: third,
                top: Math.min(tl.top, br.top),
                left: Math.min(tl.left, br.left),
                bottom: Math.max(tl.top, br.top),
                right: Math.max(tl.left, br.left),
            };
            return { bounds, sheetNames: [tl.sheetName, br.sheetName, third] };
        }
        case 4:
        case 5: {
            if (!isRowNumber(first) || !isColumnNumber(second) || !isRowNumber(third) || !isColumnNumber(fourth)) {
                return undefined;
            }
            if (fifth !== undefined && typeof fifth !== "string") {
                return undefined;
            }
            const bounds = {
                sheetName: fifth,
                top: Math.min(first, third),
                left: Math.min(second, fourth),
                bottom: Math.max(first, third),
                right: Math.max(second, fourth),
            };
            return { bounds, sheetNames: [fifth] };
        }
        default:
            return undefined;
    }
}

/**
 * The range of cells of the worksheet `worksheetName` that range arguments give (see {@link RangeArgs}: a reference
 * `"A1:B2"`, two cells `"A1", "B2"`, four bounds `1, 1, 2, 2`, a Range or a location, or such a list as one array). A
 * sheet name, wherever given, must be the worksheet's (case-insensitive): the range has none.
 * @param action - the operation, for the error messages (`"merge"`)
 * @returns a new Range, without sheet name
 * @throws Error for no range (no arguments), an invalid one (bounds that are not rows 1..1,048,576 and columns 1..16,384,
 * an unparsable reference), or a range of another sheet
 */
export function resolveRange(worksheetName: string, action: string, args: Readonly<RangeArgs>): Range {
    const decoded = decodeRangeArgs(args);
    if (!decoded) {
        throw new Error(`Cannot ${action}: invalid range (${describeArgs(args)})`);
    }
    const { bounds, sheetNames } = decoded;
    const otherSheet = sheetNames.find((sheetName) => sheetName !== undefined && !isSheet(sheetName, worksheetName));
    if (otherSheet !== undefined) {
        throw new Error(
            `Cannot ${action} ${describeArgs(args)}: the range is on the worksheet ${JSON.stringify(otherSheet)}, not ${JSON.stringify(worksheetName)}`
        );
    }
    return new Range(bounds.top, bounds.left, bounds.bottom, bounds.right);
}
