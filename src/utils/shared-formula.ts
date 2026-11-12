import colCache, { type CellAddress, type DecodedRange } from "./col-cache";

// A candidate is an optional sheet name ([a-z_\-0-9]* then '!'), then an address part ([a-z0-9_$]{2,}), then an
// optional '(' (a function name, left alone). The sheet name is only looked for at the start of a token: a regular
// expression trying it at every position rescanned the rest of a run such as 'a-a-a-...' each time (quadratic).
const CRrx = /^([$])?([a-z]+)([$])?([1-9][0-9]*)$/i;

// character classes by char code (a per-character RegExp test was several times slower)
function charClass(chars: string): Uint8Array {
    const table = new Uint8Array(128);
    for (const c of chars) {
        table[c.charCodeAt(0)] = 1;
    }
    return table;
}
const ALNUM_ = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_";
const SHEET_CHAR = charClass(`${ALNUM_}-`);
const ADDRESS_CHAR = charClass(`${ALNUM_}$`);

const isIn = (table: Uint8Array, text: string, index: number): boolean => table[text.charCodeAt(index)] === 1;

function runEnd(text: string, start: number, table: Uint8Array): number {
    let end = start;
    while (end < text.length && isIn(table, text, end)) {
        end++;
    }
    return end;
}

/**
 * The Excel limit of the length of a formula (characters). Longer formulas are invalid xlsx: each shared-formula clone
 * would repeat the scan of the master.
 */
const MAX_FORMULA_LENGTH = 8192;

/**
 * Checks the length of a formula.
 * @throws Error when it is longer than {@link MAX_FORMULA_LENGTH}.
 */
function checkFormulaLength(formula: string | null | undefined): void {
    if (formula && formula.length > MAX_FORMULA_LENGTH) {
        throw new Error(`Formula exceeds the Excel limit of ${MAX_FORMULA_LENGTH} characters`);
    }
}

// the column / row of a decoded cell address: NaN for a range (a range has none: the shift is NaN, as it always was)
const colOf = (decoded: CellAddress | DecodedRange): number => ("col" in decoded ? decoded.col : NaN);
const rowOf = (decoded: CellAddress | DecodedRange): number => ("row" in decoded ? decoded.row : NaN);

/**
 * Translates the relative references of a formula written for cell fromCell to cell toCell, as Excel does for the
 * clones of a shared formula: `slideFormula("A1+$B$1", "C1", "C2")` → `"A2+$B$1"`; a reference slid off the grid is
 * `#REF!`, as in Excel. Linear in the formula length.
 * @param fromCell - The address of the cell the formula is written for (the master of a shared formula), e.g. "C1".
 * @param toCell - The address of the target cell, e.g. "C2".
 * @throws Error when the formula is longer than {@link MAX_FORMULA_LENGTH}.
 */
function slideFormula(formula: string, fromCell: string, toCell: string): string {
    checkFormulaLength(formula);
    const offset = colCache.decode(fromCell);
    const to = colCache.decode(toCell);
    const translate = (addrPart: string): string => {
        const match = CRrx.exec(addrPart);
        // groups 2 and 4 are not optional: a match has them
        const colLetters = match?.[2];
        const rowStr = match?.[4];
        if (!match || colLetters === undefined || rowStr === undefined) {
            return addrPart;
        }
        const colDollar = match[1];
        const colStr = colLetters.toUpperCase();
        const rowDollar = match[3];
        if (colStr.length > 3 || (colStr.length === 3 && colStr > "XFD")) {
            // > XFD is the highest col number in excel 2007 and beyond, so this is a named range
            return addrPart;
        }
        let col = colCache.l2n(colStr);
        let row = parseInt(rowStr, 10);
        if (!colDollar) {
            col += colOf(to) - colOf(offset);
        }
        if (!rowDollar) {
            row += rowOf(to) - rowOf(offset);
        }
        if (col < 1 || col > colCache.MAX_COLS || row < 1 || row > colCache.MAX_ROWS) {
            // slid off the grid: Excel writes #REF! for such a reference
            return "#REF!";
        }
        return (colDollar || "") + colCache.n2l(col) + (rowDollar || "") + row;
    };

    let result = "";
    let copied = 0; // formula[0, copied) is in result
    let i = 0;
    while (i < formula.length) {
        let addrStart = i;
        if (i === 0 || !isIn(SHEET_CHAR, formula, i - 1)) {
            const sheetEnd = runEnd(formula, i, SHEET_CHAR);
            if (formula[sheetEnd] === "!" && runEnd(formula, sheetEnd + 1, ADDRESS_CHAR) - sheetEnd - 1 >= 2) {
                addrStart = sheetEnd + 1;
            }
        }
        const addrEnd = runEnd(formula, addrStart, ADDRESS_CHAR);
        if (addrEnd - addrStart < 2) {
            i++;
        } else if (formula[addrEnd] === "(") {
            i = addrEnd + 1;
        } else {
            result += formula.slice(copied, addrStart) + translate(formula.slice(addrStart, addrEnd));
            copied = i = addrEnd;
        }
    }
    return result + formula.slice(copied);
}

export { MAX_FORMULA_LENGTH, checkFormulaLength, slideFormula };
