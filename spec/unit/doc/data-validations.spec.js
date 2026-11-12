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
});
