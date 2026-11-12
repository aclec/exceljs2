import type { CellRichTextValue } from "../../index";
import { createDictionary } from "./utils";

/** A value of the shared strings table: a plain string or a rich text. */
export type SharedStringValue = string | CellRichTextValue;

/**
 * The shared strings table of a workbook being written (the streaming writer's): each distinct value is stored once,
 * and cells refer to it by its index.
 */
class SharedStrings {
    private _values: SharedStringValue[];
    private _totalRefs: number;
    // plain strings => index. No prototype: the keys are cell values.
    private _hash: Record<string, number>;
    // rich text values, keyed by the JSON of their runs. Kept apart from _hash: as an object
    // key every rich text would be "[object Object]", and a JSON key there could collide
    // with an identical plain string.
    private _richTextHash: Map<string, number>;

    constructor() {
        this._values = [];
        this._totalRefs = 0;
        this._hash = createDictionary<number>();
        this._richTextHash = new Map();
    }

    /** The number of distinct values. */
    get count(): number {
        return this._values.length;
    }

    /** The distinct values, by index. */
    get values(): SharedStringValue[] {
        return this._values;
    }

    /** The number of values added, duplicates included. */
    get totalRefs(): number {
        return this._totalRefs;
    }

    /** The value at an index, or undefined. */
    getString(index: number): SharedStringValue | undefined {
        return this._values[index];
    }

    /**
     * Adds a value (rich texts are compared by their runs, as JSON), and counts a reference to it.
     * @returns The index of the value.
     */
    add(value: SharedStringValue): number {
        let index: number | undefined;
        if (typeof value === "string") {
            index = this._hash[value];
            if (index === undefined) {
                index = this._hash[value] = this._values.length;
                this._values.push(value);
            }
        } else {
            const key = JSON.stringify(value.richText);
            index = this._richTextHash.get(key);
            if (index === undefined) {
                index = this._values.length;
                this._richTextHash.set(key, index);
                this._values.push(value);
            }
        }
        this._totalRefs++;
        return index;
    }
}

export default SharedStrings;
