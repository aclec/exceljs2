const _ = verquire("utils/under-dash");
const util = require("util");

describe("under-dash", () => {
    describe("isEqual", () => {
        const values = [
            0,
            1,
            true,
            false,
            "string",
            "foobar",
            "other string",
            [],
            ["array"],
            ["array", "foobar"],
            ["array2"],
            ["array2", "foobar"],
            {},
            { object: 1 },
            { object: 2 },
            { object: 1, foobar: "quux" },
            { object: 2, foobar: "quux" },
            null,
            undefined,
            () => {},
            () => {},
            Symbol("foo"),
            Symbol("foo"),
            Symbol("bar"),
        ];

        function showVal(o) {
            return util.inspect(o, { compact: true });
        }

        it("works on simple values", () => {
            for (let i = 0; i < values.length; i++) {
                for (let j = 0; j < values.length; j++) {
                    const a = values[i];
                    const b = values[j];

                    const assertion = `${showVal(a)} ${i === j ? "==" : "!="} ${showVal(b)}`;
                    expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
                }
            }
        });

        it("works on complex arrays", () => {
            for (let i = 0; i < values.length; i++) {
                for (let j = 0; j < values.length; j++) {
                    const a = [values[i]];
                    const b = [values[j]];

                    const assertion = `${showVal(a)} ${i === j ? "==" : "!="} ${showVal(b)}`;
                    expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
                }
            }
        });

        it("works on complex objects", () => {
            for (let i = 0; i < values.length; i++) {
                for (let j = 0; j < values.length; j++) {
                    const a = { key: values[i] };
                    const b = { key: values[j] };

                    const assertion = `${showVal(a)} ${i === j ? "==" : "!="} ${showVal(b)}`;
                    expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
                }
            }
        });
    });

    describe("deepMerge", () => {
        it("merges objects and arrays deeply, skips undefined values and prototype keys", () => {
            const target = { a: { x: 1 }, list: [1, 2] };
            const hostile = JSON.parse('{"__proto__": {"polluted": 1}, "b": 2}');
            const result = _.deepMerge(target, { a: { y: 2 }, list: [3], u: undefined }, hostile);
            expect(result).to.equal(target);
            expect(result).to.deep.equal({ a: { x: 1, y: 2 }, list: [3, 2], b: 2 });
            expect({}.polluted).to.equal(undefined);
        });

        it("creates the target when it is falsy and merges a string by its characters", () => {
            expect(_.deepMerge(null, "ab", 5, null)).to.deep.equal({ 0: "a", 1: "b" });
        });
    });

    describe("each, some, every, map", () => {
        it("iterate arrays by index (skipping holes) and objects by key", () => {
            const seen = [];
            const sparse = ["a", "b", "c"];
            delete sparse[1];
            _.each(sparse, (value, index) => seen.push([value, index]));
            _.each({ x: 1 }, (value, key) => seen.push([value, key]));
            _.each(null, () => seen.push("never"));
            expect(seen).to.deep.equal([
                ["a", 0],
                ["c", 2],
                [1, "x"],
            ]);
            expect(_.some({ a: 1, b: 2 }, (v) => v === 2)).to.equal(true);
            expect(_.every([1, 2], (v) => v > 1)).to.equal(false);
            expect(_.map({ a: 1, b: 2 }, (v, k) => k + v)).to.deep.equal(["a1", "b2"]);
            expect(_.map(undefined, () => 1)).to.deep.equal([]);
        });
    });
});
