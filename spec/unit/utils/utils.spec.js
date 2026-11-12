const utils = verquire("utils/utils");

describe("utils", () => {
    describe("xmlEncode", () => {
        it("encodes xml text", () => {
            expect(utils.xmlEncode("<")).to.equal("&lt;");
            expect(utils.xmlEncode(">")).to.equal("&gt;");
            expect(utils.xmlEncode("&")).to.equal("&amp;");
            expect(utils.xmlEncode('"')).to.equal("&quot;");
            expect(utils.xmlEncode("'")).to.equal("&apos;");

            expect(
                utils.xmlEncode(
                    "abc\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x20abc\x7f"
                )
            ).to.equal("abc abc");

            expect(utils.xmlEncode('<a href="www.whatever.com">Talk to the H&</a>')).to.equal(
                "&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;"
            );

            expect(utils.xmlEncode("new\x0aline")).to.equal("new\x0aline");
        });
    });
    describe("isDateFmt", () => {
        // an escaped quote does not start a literal text
        ["yyyy-mm-dd", "d\\ mmmm\\ yyyy", "[$-40C]d\\-mmm", '\\"yyyy\\"'].forEach((fmt) => {
            it(`'${fmt}' a date`, () => {
                expect(utils.isDateFmt(fmt)).to.be.true();
            });
        });

        // escaped (\m), padding (_m) and fill (*m) characters are literal text
        ["", "[Green]#,##0 ;[Red](#,##0)", "0.0\\ \\m", "#,##0_m", "0*m", '0"*" "pts"', '"a_" 0 "d"', '0"\\" "m"'].forEach((fmt) => {
            it(`'${fmt}' is not a date`, () => {
                expect(utils.isDateFmt(fmt)).to.be.false();
            });
        });

        // the former implementation, quadratic on '[[[[...' (only run on short strings here)
        const reference = (fmt) => {
            if (!fmt) {
                return false;
            }
            return /[ymdhMsb]+/.test(fmt.replace(/\[[^\]]*]/g, "").replace(/"[^"]*"/g, ""));
        };

        it("matches the regular expression implementation", () => {
            const fixed = [
                "[h]:mm",
                '"d"0',
                "[$-409]d-mmm",
                '0"[y]"',
                '"["y"]"',
                '[y"]"d',
                '["y]"',
                "[[y]",
                "[y",
                '"y',
                "]y]",
                "0.00",
                "General",
            ];
            fixed.forEach((fmt) => expect(utils.isDateFmt(fmt), fmt).to.equal(reference(fmt)));
            // a fixed pseudo-random sequence over the characters that matter
            const alphabet = '[]"yd0 ';
            let seed = 42;
            for (let i = 0; i < 5000; i++) {
                let fmt = "";
                const length = i % 12;
                for (let j = 0; j < length; j++) {
                    seed = (seed * 1103515245 + 12345) % 2147483648;
                    fmt += alphabet[seed % alphabet.length];
                }
                expect(utils.isDateFmt(fmt), fmt).to.equal(reference(fmt));
            }
        });

        it("runs in linear time on unbalanced brackets and quotes", () => {
            const start = Date.now();
            expect(utils.isDateFmt("[".repeat(100000))).to.be.false();
            expect(utils.isDateFmt('"'.repeat(100000))).to.be.false();
            expect(utils.isDateFmt(`${"[".repeat(100000)}y`)).to.be.true();
            expect(Date.now() - start).to.be.below(300);
        });
    });

    describe("dateToExcel", () => {
        it("should convert date to excel properly", () => {
            const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));

            const excelDate = utils.dateToExcel(myDate, false);

            expect(excelDate).to.equal(43084.70833333333);
        });
    });

    describe("excelToDate", () => {
        it("should round to the nearest millisecond when parsing excel date", () => {
            const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));
            const excelDate = utils.dateToExcel(myDate, false);

            const dateConverted = utils.excelToDate(excelDate, false);

            expect(dateConverted).to.deep.equal(myDate);
        });
        it("should not lost millisecond precision when parsing excel date", () => {
            const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));
            const excelDate = utils.dateToExcel(myDate, false);

            const dateConverted = utils.excelToDate(excelDate, false);

            expect(dateConverted).to.deep.equal(myDate);
        });
    });

    describe("parseBoolean", () => {
        it("reads xsd:boolean values", () => {
            ["1", "true", 1, true].forEach((value) => expect(utils.parseBoolean(value)).to.equal(true));
            ["0", "false", 0, false].forEach((value) => expect(utils.parseBoolean(value)).to.equal(false));
        });

        it("returns the given default for an absent attribute", () => {
            expect(utils.parseBoolean(undefined)).to.equal(false);
            expect(utils.parseBoolean(undefined, true)).to.equal(true);
            expect(utils.parseBoolean("0", true)).to.equal(false);
            expect(utils.parseBoolean("false", true)).to.equal(false);
        });
    });

    describe("toIsoDateString", () => {
        it("gives the UTC date part of the ISO form (it always threw a TypeError)", () => {
            expect(utils.toIsoDateString(new Date(Date.UTC(2024, 0, 31, 23, 59)))).to.equal("2024-01-31");
        });
    });

    describe("createDictionary", () => {
        it("creates an empty object without prototype", () => {
            const dictionary = utils.createDictionary();
            expect(Object.getPrototypeOf(dictionary)).to.equal(null);
            expect(Object.keys(dictionary)).to.deep.equal([]);
            dictionary.__proto__ = 1;
            expect(dictionary.__proto__).to.equal(1);
            expect(Object.keys(dictionary)).to.deep.equal(["__proto__"]);
        });
    });

    describe("inherits", () => {
        it("copies the own enumerable statics and prototype properties", () => {
            function Base() {}
            Base.prototype.base = () => "base";
            function Derived() {}
            const statics = { s: 1 };
            Object.defineProperty(statics, "hidden", { value: 2, enumerable: false });
            const prototype = {
                get g() {
                    return "g";
                },
            };
            Object.defineProperty(prototype, "hiddenMethod", { value: () => 0, enumerable: false });
            utils.inherits(Derived, Base, statics, prototype);
            expect(Derived.s).to.equal(1);
            expect(Object.hasOwn(Derived, "hidden")).to.equal(false);
            const derived = new Derived();
            expect(derived.g).to.equal("g");
            expect(derived.base()).to.equal("base");
            expect(derived.hiddenMethod).to.equal(undefined);
            expect(Derived.super_).to.equal(Base);
        });
    });
});
