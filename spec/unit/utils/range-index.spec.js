const RangeIndex = verquire("utils/range-index");

describe("RangeIndex", () => {
    it("finds the rectangle added first among those containing a cell", () => {
        const index = new RangeIndex();
        const big = { top: 1, left: 1, bottom: 10, right: 10 };
        const small = { top: 2, left: 2, bottom: 3, right: 3 };
        index.add(small, 1);
        index.add(big, 0);
        expect(index.find(2, 2)).to.equal(big);
        expect(index.find(20, 20)).to.be.undefined();
        index.remove(big);
        expect(index.find(2, 2)).to.equal(small);
        expect(index.find(5, 5)).to.be.undefined();
    });

    it("throws a clear TypeError when removing a range added to another index", () => {
        const range = { top: 1, left: 1, bottom: 1, right: 1 };
        new RangeIndex().add(range, 0);
        expect(() => new RangeIndex().remove(range)).to.throw(TypeError, "RangeIndex: this range was not added");
        expect(() => new RangeIndex().remove({ top: 1, left: 1, bottom: 1, right: 1 })).to.throw(
            TypeError,
            "RangeIndex: this range was not added"
        );
    });
});
