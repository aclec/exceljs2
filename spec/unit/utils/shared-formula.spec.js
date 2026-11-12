const { slideFormula } = verquire("utils/shared-formula");
const colCache = verquire("utils/col-cache");

// the former implementation, quadratic on 'a-a-a-...' (only run on short formulas here)
function referenceSlideFormula(formula, fromCell, toCell) {
    const offset = colCache.decode(fromCell);
    const to = colCache.decode(toCell);
    return formula.replace(/(([a-z_\-0-9]*)!)?([a-z0-9_$]{2,})([(])?/gi, (refMatch, sheet, sheetMaybe, addrPart, trailingParen) => {
        if (trailingParen) {
            return refMatch;
        }
        const match = /^([$])?([a-z]+)([$])?([1-9][0-9]*)$/i.exec(addrPart);
        if (!match) {
            return refMatch;
        }
        const colStr = match[2].toUpperCase();
        if (colStr.length > 3 || (colStr.length === 3 && colStr > "XFD")) {
            return refMatch;
        }
        const col = colCache.l2n(colStr) + (match[1] ? 0 : to.col - offset.col);
        const row = parseInt(match[4], 10) + (match[3] ? 0 : to.row - offset.row);
        return (sheet || "") + (match[1] || "") + colCache.n2l(col) + (match[3] || "") + row;
    });
}

describe("shared-formula", () => {
    describe("slideFormula", () => {
        const expectations = [
            { args: ["A1+1", "A2", "A3"], result: "A2+1" },
            { args: ["A1+1", "A2", "B2"], result: "B1+1" },
            { args: ["SUM(A1:A10)", "A11", "B11"], result: "SUM(B1:B10)" },
            { args: ["$A$1+A1", "A2", "A3"], result: "$A$1+A2" },
            { args: ["$A$1+A1", "A2", "B2"], result: "$A$1+B1" },
            { args: ["$A1+A1", "A2", "A3"], result: "$A2+A2" },
            { args: ["$A1+A1", "A2", "B2"], result: "$A1+B1" },
            { args: ["A$1+A1", "A2", "A3"], result: "A$1+A2" },
            { args: ["A$1+A1", "A2", "B2"], result: "B$1+B1" },
        ];
        expectations.forEach(({ args, result }) => {
            it(`${args[0]} from ${args[1]} to ${args[2]}`, () => {
                expect(slideFormula(...args)).to.equal(result);
            });
        });

        it("translates like the regular expression implementation", () => {
            // a reference slid off the grid throws in both
            const outcome = (fn, ...args) => {
                try {
                    return fn(...args);
                } catch (error) {
                    return `throws ${error.message}`;
                }
            };
            const fixed = ["Sheet1!A1+B2", "a-Sheet1!A1", "A1-x!B2", "A1-B2", "SUM(A1:B2)", "x_1!$A$1*A1", "-x!A1", "a-bc!+C3", "A1$-x!B2"];
            fixed.forEach((formula) =>
                expect(outcome(slideFormula, formula, "A1", "C4"), formula).to.equal(outcome(referenceSlideFormula, formula, "A1", "C4"))
            );
            // a fixed pseudo-random sequence over the characters that matter
            const alphabet = ["A", "b", "1", "2", "-", "!", "$", "_", "(", "+", ":", "XFD", "XFE", " "];
            let seed = 7;
            for (let i = 0; i < 20000; i++) {
                let formula = "";
                const length = i % 14;
                for (let j = 0; j < length; j++) {
                    seed = (seed * 1103515245 + 12345) % 2147483648;
                    formula += alphabet[seed % alphabet.length];
                }
                expect(outcome(slideFormula, formula, "B2", "C5"), formula).to.equal(outcome(referenceSlideFormula, formula, "B2", "C5"));
            }
        });

        it("translates a long formula in linear time", () => {
            const hostile = "a-".repeat(4096);
            const start = Date.now();
            for (let i = 0; i < 100; i++) {
                expect(slideFormula(hostile, "A1", `A${i + 2}`)).to.equal(hostile);
            }
            expect(Date.now() - start).to.be.below(300); // was ~30 ms per call
        });
    });
});
