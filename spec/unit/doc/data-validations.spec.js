const DataValidations = verquire("doc/data-validations");
const colCache = verquire("utils/col-cache");

const LIST = { type: "list", formulae: ['"a,b"'] };
const WHOLE = { type: "whole", operator: "greaterThan", formulae: [0] };

describe("DataValidations", () => {
    it("adds validations to many cells in linear time", () => {
        const dataValidations = new DataValidations({ "B1:B1048576": LIST });
        const start = Date.now();
        for (let row = 1; row <= 5000; row++) {
            const address = `C${row}`;
            if (!dataValidations.find(address)) {
                dataValidations.add(address, WHOLE);
            }
        }
        // the range index used to be rebuilt from every key after each add: quadratic
        expect(Date.now() - start).to.be.below(250);
        expect(dataValidations.find("C5000")).to.equal(WHOLE);
        expect(dataValidations.find("B5000")).to.equal(LIST);
    });

    it("finds range keys added or removed after a lookup", () => {
        const dataValidations = new DataValidations();
        expect(dataValidations.find("D4")).to.equal(undefined);
        dataValidations.add("C3:E5", LIST);
        expect(dataValidations.find("D4")).to.equal(LIST);
        dataValidations.remove("C3:E5");
        expect(dataValidations.find("D4")).to.equal(undefined);
    });

    it("sees a range key written directly into the model after a lookup", () => {
        const dataValidations = new DataValidations();
        expect(dataValidations.find("B5")).to.equal(undefined);
        dataValidations.model["B1:B10"] = WHOLE;
        expect(dataValidations.find("B5")).to.equal(WHOLE);
        delete dataValidations.model["B1:B10"];
        expect(dataValidations.find("B5")).to.equal(undefined);
    });

    it("applies the first of overlapping range keys", () => {
        const dataValidations = new DataValidations({ "A1:A10": LIST, "A5:A20": WHOLE });
        expect(dataValidations.find("A7")).to.equal(LIST);
        expect(dataValidations.find("A15")).to.equal(WHOLE);
        dataValidations.remove("A1:A10");
        expect(dataValidations.find("A7")).to.equal(WHOLE);
        dataValidations.add("A1:XFD8", LIST);
        expect(dataValidations.find("A7")).to.equal(WHOLE);
        expect(dataValidations.find("Z7")).to.equal(LIST);
    });

    it("finds cells among many range keys without scanning them all", () => {
        // 250,000 range keys, as kept from a file once 100,000 cells were expanded
        const model = {};
        for (let i = 0; i < 250000; i++) {
            model[`C${2 * i + 1}:C${2 * i + 2}`] = WHOLE;
        }
        const dataValidations = new DataValidations(model);
        expect(dataValidations.find("C1")).to.equal(WHOLE); // builds the index
        const start = Date.now();
        for (let row = 1; row <= 5000; row++) {
            expect(dataValidations.find(`A${row}`)).to.equal(undefined);
            expect(dataValidations.find(`C${row * 97}`)).to.equal(WHOLE);
        }
        // a scan of every range key took about 0.75 ms per lookup
        expect(Date.now() - start).to.be.below(250);
    });

    it("finds the same validations as a scan of every range key", () => {
        let seed = 1;
        const random = (n) => {
            seed = (seed * 16807) % 2147483647;
            return seed % n;
        };
        const sizes = [1, 2, 3, 7, 64, 1000, 16384, 1048576];
        const dataValidations = new DataValidations();
        const keys = [];
        for (let i = 0; i < 400; i++) {
            const height = Math.min(sizes[random(sizes.length)], 2000);
            const width = Math.min(sizes[random(sizes.length)], 300);
            const top = 1 + random(2000);
            const left = 1 + random(300);
            const key = `${colCache.n2l(left)}${top}:${colCache.n2l(left + width - 1)}${top + height}`;
            dataValidations.add(key, { id: i });
            keys.push(key);
            if (i % 5 === 0) {
                dataValidations.remove(keys[random(keys.length)]);
            }
            if (i % 7 === 0) {
                dataValidations.find("A1"); // lookups between the changes
            }
        }
        const ranges = Object.keys(dataValidations.model)
            .filter((key) => dataValidations.model[key])
            .map((key) => ({ key, ...colCache.decode(key) }));
        for (let i = 0; i < 3000; i++) {
            const row = 1 + random(2100);
            const col = 1 + random(320);
            const match = ranges.find((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right);
            const address = colCache.encodeAddress(row, col);
            expect(dataValidations.find(address), address).to.equal(match && dataValidations.model[match.key]);
        }
    });

    describe("a cell inside a range key", () => {
        it("is removed by splitting the range around it", () => {
            const dataValidations = new DataValidations({ "B2:D4": LIST });
            dataValidations.remove("C3");
            expect(dataValidations.find("C3")).to.equal(undefined);
            ["B2", "C2", "D2", "B3", "D3", "B4", "C4", "D4"].forEach((address) => {
                expect(dataValidations.find(address), address).to.equal(LIST);
            });
            // no key covers C3 any more, so the removal survives a write
            expect(Object.keys(dataValidations.model).sort()).to.deep.equal(["B2:D2", "B3", "B4:D4", "C3", "D3"]);
        });

        it("keeps the existing keys that a split would overwrite", () => {
            const OTHER = { type: "any" };
            const dataValidations = new DataValidations();
            dataValidations.add("B1", WHOLE);
            dataValidations.add("A1:B1", LIST);
            dataValidations.add("A1", OTHER);
            expect(dataValidations.find("B1")).to.equal(WHOLE);
            expect(dataValidations.find("A1")).to.equal(OTHER);
            expect(dataValidations.model).to.deep.equal({ B1: WHOLE, A1: OTHER });
        });

        it("overrides many cells of a whole-column range in linear time", () => {
            const dataValidations = new DataValidations({ "A1:A1048576": LIST });
            const start = Date.now();
            for (let row = 1; row <= 20000; row++) {
                dataValidations.add(`A${row * 7}`, WHOLE); // leaves a range key above each cell
            }
            expect(Date.now() - start).to.be.below(250);
            expect(dataValidations.find("A140000")).to.equal(WHOLE);
            expect(dataValidations.find("A139999")).to.equal(LIST);
            expect(dataValidations.find("A140001")).to.equal(LIST);
        });

        it("is overridden by splitting the range around it", () => {
            const dataValidations = new DataValidations({ "A1:A1048576": LIST });
            dataValidations.add("A7", WHOLE);
            expect(dataValidations.find("A7")).to.equal(WHOLE);
            expect(dataValidations.find("A6")).to.equal(LIST);
            expect(dataValidations.find("A8")).to.equal(LIST);
            expect(Object.keys(dataValidations.model).sort()).to.deep.equal(["A1:A6", "A7", "A8:A1048576"]);
        });
    });
});
