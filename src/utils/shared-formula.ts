/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import colCache from "./col-cache";

// A candidate is an optional sheet name ([a-z_\-0-9]* then '!'), then an address part ([a-z0-9_$]{2,}), then an
// optional '(' (a function name, left alone). The sheet name is only looked for at the start of a token: a regular
// expression trying it at every position rescanned the rest of a run such as 'a-a-a-...' each time (quadratic).
const CRrx = /^([$])?([a-z]+)([$])?([1-9][0-9]*)$/i;

// character classes by char code (a per-character RegExp test was several times slower)
function charClass(chars) {
    const table = new Uint8Array(128);
    for (const c of chars) {
        table[c.charCodeAt(0)] = 1;
    }
    return table;
}
const ALNUM_ = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_";
const SHEET_CHAR = charClass(`${ALNUM_}-`);
const ADDRESS_CHAR = charClass(`${ALNUM_}$`);

const isIn = (table, text, index) => table[text.charCodeAt(index)] === 1;

function runEnd(text, start, table) {
    let end = start;
    while (end < text.length && isIn(table, text, end)) {
        end++;
    }
    return end;
}

// Excel limit (characters). Longer formulas are invalid xlsx: each shared-formula clone would
// repeat the scan of the master.
const MAX_FORMULA_LENGTH = 8192;

function checkFormulaLength(formula) {
    if (formula && formula.length > MAX_FORMULA_LENGTH) {
        throw new Error(`Formula exceeds the Excel limit of ${MAX_FORMULA_LENGTH} characters`);
    }
}

function slideFormula(formula, fromCell, toCell) {
    checkFormulaLength(formula);
    const offset = colCache.decode(fromCell);
    const to = colCache.decode(toCell);
    const translate = (addrPart) => {
        const match = CRrx.exec(addrPart);
        if (!match) {
            return addrPart;
        }
        const colDollar = match[1];
        const colStr = match[2].toUpperCase();
        const rowDollar = match[3];
        const rowStr = match[4];
        if (colStr.length > 3 || (colStr.length === 3 && colStr > "XFD")) {
            // > XFD is the highest col number in excel 2007 and beyond, so this is a named range
            return addrPart;
        }
        let col = colCache.l2n(colStr);
        let row = parseInt(rowStr, 10);
        if (!colDollar) {
            col += to.col - offset.col;
        }
        if (!rowDollar) {
            row += to.row - offset.row;
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
