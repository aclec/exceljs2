// Shared strings models: the shared strings table (`xl/sharedStrings.xml`, src/xlsx/xform/strings) and the managers that
// build it (SharedStringsXform for the document writer, utils/shared-strings for the stream writer).

import type { RichTextValue } from "./cell";

/**
 * A shared string: plain text, or rich text (`{ richText: [...] }`).
 *
 * Producer: shared-string-xform parse (an empty `<si/>`, or one holding only phonetic runs, is `""`; a leading `<t>` before
 * `<r>` runs becomes the first run), the cells (String / RichText / Hyperlink values) through {@link SharedStringRegistry.add}.
 * Consumer: shared-string-xform render, cell-xform reconcile (`getString`).
 */
export type SharedStringValue = string | RichTextValue;

/**
 * The model of the shared strings table (SharedStringsXform): `values` are the unique strings, `count` the number of
 * references to them (cells).
 */
export interface SharedStringsModel {
    values: SharedStringValue[];
    count: number;
}

/** The write side of a shared strings manager, passed as `options.sharedStrings` to the cell `prepare()`. */
export interface SharedStringRegistry {
    /**
     * Adds a reference to a string (counted), adding the string to the table if it is new.
     *
     * @returns its index in the table
     */
    add(value: SharedStringValue): number;
}

/** The read side of a shared strings manager, passed as `options.sharedStrings` to the cell `reconcile()`. */
export interface SharedStringLookup {
    /** @returns the string at an index of the table (`undefined` when out of range: the index comes from the file) */
    getString(index: number): SharedStringValue | undefined;
}
