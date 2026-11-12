const DataValidations = verquire("doc/data-validations");

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
