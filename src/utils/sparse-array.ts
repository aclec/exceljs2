// Rows, cells and defined-name cells are kept in arrays indexed by row or column number, which a single far
// row (1,048,576) or cell (column XFD) makes that long. forEach and index loops visit every index below the
// length: about 14 ms for 1,048,576 indexes in V8 and JSC, even when one entry exists. Object.keys lists the
// existing entries only, in time proportional to them for these sparse (dictionary-mode) arrays, but it
// creates a string per entry: 10 to 70 times slower than a loop on a dense array. So the arrays are looped
// over until a run of MAX_GAP holes, and only a sparse array is enumerated by its keys.
const MAX_GAP = 32;

/** An array indexed by row or column number, with holes (or `undefined` entries) for the missing ones. */
export type SparseArray<T> = readonly (T | undefined)[];

/**
 * Calls predicate(value, index) for each entry of array from index start holding a value other than undefined,
 * in index order, until it returns a truthy value. Values are read when they are reached, as forEach does, but
 * entries added past the first run of MAX_GAP holes while it runs may not be visited.
 * @returns Whether the predicate returned a truthy value.
 */
export function someEntry<T>(array: SparseArray<T>, predicate: (value: T, index: number) => unknown, start = 0): boolean {
    const { length } = array;
    let gap = 0;
    for (let i = start; i < length; i++) {
        const value = array[i];
        if (value !== undefined) {
            if (predicate(value, i)) {
                return true;
            }
            gap = 0;
        } else if (++gap > MAX_GAP) {
            // sparse: the entries after i only
            const keys = Object.keys(array);
            for (let k = 0; k < keys.length; k++) {
                const index = Number(keys[k]);
                if (index > i && index < length) {
                    const entry = array[index];
                    if (entry !== undefined && predicate(entry, index)) {
                        return true;
                    }
                }
            }
            return false;
        }
    }
    return false;
}

/**
 * Calls callback(value, index) for each entry of array from index start holding a value other than undefined,
 * in index order (someEntry without its predicate: this runs for every row and cell).
 */
export function eachEntry<T>(array: SparseArray<T>, callback: (value: T, index: number) => void, start = 0): void {
    const { length } = array;
    let gap = 0;
    for (let i = start; i < length; i++) {
        const value = array[i];
        if (value !== undefined) {
            callback(value, i);
            gap = 0;
        } else if (++gap > MAX_GAP) {
            const keys = Object.keys(array);
            for (let k = 0; k < keys.length; k++) {
                const index = Number(keys[k]);
                if (index > i && index < length) {
                    const entry = array[index];
                    if (entry !== undefined) {
                        callback(entry, index);
                    }
                }
            }
            return;
        }
    }
}

/** The indices of the entries of array holding a value other than undefined, ascending. */
export function entryIndices(array: SparseArray<unknown>): number[] {
    const indices: number[] = [];
    eachEntry(array, (value, index) => {
        indices.push(index);
    });
    return indices;
}

/** The index of the last entry of array holding a value other than undefined, or -1. */
export function lastEntryIndex(array: SparseArray<unknown>): number {
    const { length } = array;
    const stop = Math.max(0, length - MAX_GAP);
    for (let i = length - 1; i >= stop; i--) {
        if (array[i] !== undefined) {
            return i;
        }
    }
    if (stop === 0) {
        return -1;
    }
    // sparse: the last of the keys below stop
    const keys = Object.keys(array);
    for (let k = keys.length - 1; k >= 0; k--) {
        const index = Number(keys[k]);
        if (index < stop && array[index] !== undefined) {
            return index;
        }
    }
    return -1;
}

/**
 * The steps i of a loop over first..last (any order) that moves the entry at i to i + offset, whose source i or
 * destination i + offset is in indices: the steps that are not no-ops.
 * @param indices - Sorted, ascending.
 * @returns The steps, ascending.
 */
export function shiftSteps(indices: readonly number[], first: number, last: number, offset: number): number[] {
    const steps: number[] = [];
    let s = 0; // next source
    let d = 0; // next destination
    for (;;) {
        // undefined past the end of indices
        let sourceIndex = indices[s];
        while (sourceIndex !== undefined && sourceIndex < first) {
            sourceIndex = indices[++s];
        }
        let destinationIndex = indices[d];
        while (destinationIndex !== undefined && destinationIndex - offset < first) {
            destinationIndex = indices[++d];
        }
        const source = sourceIndex !== undefined && sourceIndex <= last ? sourceIndex : Infinity;
        const destination = destinationIndex !== undefined && destinationIndex - offset <= last ? destinationIndex - offset : Infinity;
        const step = Math.min(source, destination);
        if (step === Infinity) {
            return steps;
        }
        steps.push(step);
        first = step + 1;
    }
}

/**
 * The steps of a loop over first..last moving the entry for number i to i + offset, where the entry of array at
 * index n holds number n + base: all of them for a short loop, else the {@link shiftSteps} of the entries.
 * @returns The steps, ascending.
 */
export function loopSteps(array: SparseArray<unknown>, base: number, first: number, last: number, offset: number): number[] {
    if (last - first < 2 * MAX_GAP) {
        const steps: number[] = [];
        for (let i = first; i <= last; i++) {
            steps.push(i);
        }
        return steps;
    }
    return shiftSteps(
        entryIndices(array).map((index) => index + base),
        first,
        last,
        offset
    );
}

/**
 * Like array.splice(start, numDelete, <numInsert holes>), returning a new array, in time proportional to the
 * existing entries (splice moves every index after start).
 */
export function spliceEntries<T>(array: SparseArray<T>, start: number, numDelete: number, numInsert: number): T[] {
    const result: T[] = [];
    eachEntry(array, (value, index) => {
        if (index < start) {
            result[index] = value;
        } else if (index >= start + numDelete) {
            result[index - numDelete + numInsert] = value;
        }
    });
    return result;
}

export default { someEntry, eachEntry, entryIndices, lastEntryIndex, shiftSteps, loopSteps, spliceEntries };
