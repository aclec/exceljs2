"use strict";

// Rows, cells and defined-name cells are kept in arrays indexed by row or column number, which a single far
// row (1,048,576) or cell (column XFD) makes that long. forEach and index loops visit every index below the
// length: about 14 ms for 1,048,576 indexes in V8 and JSC, even when one entry exists. Object.keys lists the
// existing entries only, in time proportional to them for these sparse (dictionary-mode) arrays, but it
// creates a string per entry: 10 to 70 times slower than a loop on a dense array. So the arrays are looped
// over until a run of MAX_GAP holes, and only a sparse array is enumerated by its keys.
const MAX_GAP = 32;

// Calls predicate(value, index) for each entry of array holding a value other than undefined, in index order,
// until it returns a truthy value. Returns whether it did. Values are read when they are reached, as forEach
// does, but entries added past the first run of MAX_GAP holes while it runs may not be visited.
function someEntry(array, predicate) {
    const { length } = array;
    let gap = 0;
    for (let i = 0; i < length; i++) {
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

// Calls callback(value, index) for each entry of array holding a value other than undefined, in index order
function eachEntry(array, callback) {
    someEntry(array, (value, index) => {
        callback(value, index);
        return false;
    });
}

// The indices of the entries of array holding a value other than undefined, ascending
function entryIndices(array) {
    const indices = [];
    eachEntry(array, (value, index) => {
        indices.push(index);
    });
    return indices;
}

// The index of the last entry of array holding a value other than undefined, or -1
function lastEntryIndex(array) {
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

// The steps i of a loop over first..last (any order) that moves the entry at i to i + offset, whose source i or
// destination i + offset is in indices (sorted, ascending): the steps that are not no-ops. Ascending.
function shiftSteps(indices, first, last, offset) {
    const steps = [];
    let s = 0; // next source
    let d = 0; // next destination
    for (;;) {
        while (s < indices.length && indices[s] < first) {
            s++;
        }
        while (d < indices.length && indices[d] - offset < first) {
            d++;
        }
        const source = s < indices.length && indices[s] <= last ? indices[s] : Infinity;
        const destination = d < indices.length && indices[d] - offset <= last ? indices[d] - offset : Infinity;
        const step = Math.min(source, destination);
        if (step === Infinity) {
            return steps;
        }
        steps.push(step);
        first = step + 1;
    }
}

// Like array.splice(start, numDelete, <numInsert holes>), returning a new array, in time proportional to the
// existing entries (splice moves every index after start)
function spliceEntries(array, start, numDelete, numInsert) {
    const result = [];
    eachEntry(array, (value, index) => {
        if (index < start) {
            result[index] = value;
        } else if (index >= start + numDelete) {
            result[index - numDelete + numInsert] = value;
        }
    });
    return result;
}

module.exports = { someEntry, eachEntry, entryIndices, lastEntryIndex, shiftSteps, spliceEntries };
