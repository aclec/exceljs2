const { eachEntry, someEntry, lastEntryIndex, entryIndices, shiftSteps, spliceEntries } = verquire("utils/sparse-array");

// the entries forEach visits, without its undefined values
function forEachEntries(array) {
    const entries = [];
    array.forEach((value, index) => {
        if (value !== undefined) {
            entries.push([index, value]);
        }
    });
    return entries;
}

function eachEntries(array) {
    const entries = [];
    eachEntry(array, (value, index) => {
        entries.push([index, value]);
    });
    return entries;
}

describe("sparse-array", () => {
    const dense = ["a", "b", undefined, "d", null, 0, ""];
    const holey = [];
    holey[3] = "x";
    holey[40] = undefined;
    holey[41] = "y";
    holey[200] = "z";
    holey[201] = undefined;
    const far = [];
    far[0] = "first";
    far[1048575] = "last";

    describe("eachEntry", () => {
        it("visits the entries that hold a value other than undefined, in index order, like forEach", () => {
            [dense, holey, far, [], [undefined]].forEach((array) => {
                expect(eachEntries(array)).to.deep.equal(forEachEntries(array));
            });
            expect(eachEntries(far)).to.deep.equal([
                [0, "first"],
                [1048575, "last"],
            ]);
        });

        it("visits the entries of an array whose holes come before its first entry", () => {
            const array = [];
            array[500] = 1;
            array[100000] = 2;
            expect(eachEntries(array)).to.deep.equal([
                [500, 1],
                [100000, 2],
            ]);
        });

        it("reads the values when it reaches them, like forEach", () => {
            const array = [1, 2, 3, 4];
            const seen = [];
            eachEntry(array, (value, index) => {
                seen.push(value);
                array[index + 1] = undefined;
            });
            expect(seen).to.deep.equal([1, 3]);
        });
    });

    describe("someEntry", () => {
        it("stops at the first entry the predicate accepts", () => {
            const seen = [];
            const found = someEntry(holey, (value, index) => {
                seen.push(index);
                return value === "y";
            });
            expect(found).to.equal(true);
            expect(seen).to.deep.equal([3, 41]);
            expect(someEntry(far, (value) => value === "last")).to.equal(true);
            expect(someEntry(far, (value) => value === "none")).to.equal(false);
            expect(someEntry([], () => true)).to.equal(false);
        });
    });

    describe("lastEntryIndex", () => {
        it("returns the index of the last value other than undefined, or -1", () => {
            expect(lastEntryIndex(dense)).to.equal(6);
            expect(lastEntryIndex(holey)).to.equal(200);
            expect(lastEntryIndex(far)).to.equal(1048575);
            const cleared = far.slice(0, 1);
            cleared[1048575] = undefined;
            expect(lastEntryIndex(cleared)).to.equal(0);
            expect(lastEntryIndex([])).to.equal(-1);
            expect(lastEntryIndex([undefined, undefined])).to.equal(-1);
        });
    });

    describe("entryIndices", () => {
        it("lists the indices of the entries", () => {
            expect(entryIndices(holey)).to.deep.equal([3, 41, 200]);
            expect(entryIndices(far)).to.deep.equal([0, 1048575]);
        });
    });

    describe("shiftSteps", () => {
        // the steps of a loop over first..last moving array[i] to array[i + offset] that read or write an entry
        function naiveSteps(indices, first, last, offset) {
            const set = new Set(indices);
            const steps = [];
            for (let i = first; i <= last; i++) {
                if (set.has(i) || set.has(i + offset)) {
                    steps.push(i);
                }
            }
            return steps;
        }

        it("lists, in order, the steps whose source or destination holds an entry", () => {
            const indices = [0, 2, 3, 7, 9, 15];
            [
                [0, 20, 2],
                [3, 12, -2],
                [5, 9, 1],
                [0, 15, 0],
                [10, 30, -5],
            ].forEach(([first, last, offset]) => {
                expect(shiftSteps(indices, first, last, offset)).to.deep.equal(naiveSteps(indices, first, last, offset));
            });
            expect(shiftSteps([0, 1048575], 1, 1048575, -1)).to.deep.equal([1, 1048575]);
            expect(shiftSteps([], 0, 10, 1)).to.deep.equal([]);
        });
    });

    describe("spliceEntries", () => {
        it("removes and inserts holes in time proportional to the entries", () => {
            const array = [];
            array[1] = "a";
            array[5] = "b";
            array[1048575] = "c";
            const spliced = spliceEntries(array, 2, 3, 1);
            expect(eachEntries(spliced)).to.deep.equal([
                [1, "a"],
                [3, "b"],
                [1048573, "c"],
            ]);
        });
    });
});
