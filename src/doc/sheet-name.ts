// The rules of the name of a sheet (worksheet or chartsheet), as Excel applies them, and the options of a chartsheet
// added (`workbook.addChartsheet`, document and streaming).

import type { WorksheetState } from "../../index";

/** The states of a sheet. */
const SHEET_STATES: readonly WorksheetState[] = ["visible", "hidden", "veryHidden"];

/** The characters a sheet name cannot have (a single character class: linear). */
const INVALID_CHARACTERS = /[*?:/\\[\]]/;

/**
 * Checks a sheet name: a string, not empty, not "History", without any of `* ? : \ / [ ]`, not starting or ending with
 * a quote, not the name of another sheet of the workbook (case-insensitive). A name longer than 31 characters is
 * truncated (with a warning).
 *
 * @param isTaken - whether a name (in lower case) is the name of another sheet
 * @returns the name, truncated to 31 characters
 * @throws Error for an invalid name
 */
export function checkSheetName(name: unknown, isTaken: (lowerName: string) => boolean): string {
    if (typeof name !== "string") {
        throw new Error("The name has to be a string.");
    }

    if (name === "") {
        throw new Error("The name can't be empty.");
    }

    if (name === "History") {
        throw new Error('The name "History" is protected. Please use a different name.');
    }

    // Illegal character in worksheet name: asterisk (*), question mark (?),
    // colon (:), forward slash (/ \), or bracket ([])
    if (INVALID_CHARACTERS.test(name)) {
        throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
    }

    if (name.startsWith("'") || name.endsWith("'")) {
        throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
    }

    let checked = name;
    if (checked.length > 31) {
        console.warn(`Worksheet name ${checked} exceeds 31 chars. This will be truncated`);
        checked = checked.substring(0, 31);
    }

    if (isTaken(checked.toLowerCase())) {
        throw new Error(`Worksheet name already exists: ${checked}`);
    }
    return checked;
}

/**
 * The state of a chartsheet added (`options.state` of `addChartsheet`), `visible` by default.
 * @throws Error for options that are not an object, or a state that is not one of the states of a sheet
 */
export function chartsheetState(options: unknown): WorksheetState {
    if (options === undefined) {
        return "visible";
    }
    if (typeof options !== "object" || options === null) {
        throw new Error("Invalid chartsheet options: expected an object");
    }
    const state: unknown = Reflect.get(options, "state");
    if (state === undefined) {
        return "visible";
    }
    const found = SHEET_STATES.find((value) => value === state);
    if (!found) {
        const shown = typeof state === "string" ? JSON.stringify(state.slice(0, 64)) : typeof state;
        throw new Error(`Invalid chartsheet state: ${shown}, expected one of ${SHEET_STATES.join(", ")}`);
    }
    return found;
}
