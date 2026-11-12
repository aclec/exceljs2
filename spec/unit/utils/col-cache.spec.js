const colCache = verquire("utils/col-cache");

describe("colCache", () => {
    it("caches values", () => {
        expect(colCache.l2n("A")).to.equal(1);
        expect(colCache._l2n.A).to.equal(1);
        expect(colCache._n2l[1]).to.equal("A");

        // also, because of the fill heuristic A-Z will be there too
        const dic = [
            "A",
            "B",
            "C",
            "D",
            "E",
            "F",
            "G",
            "H",
            "I",
            "J",
            "K",
            "L",
            "M",
            "N",
            "O",
            "P",
            "Q",
            "R",
            "S",
            "T",
            "U",
            "V",
            "W",
            "X",
            "Y",
            "Z",
        ];
        dic.forEach((letter, index) => {
            expect(colCache._l2n[letter]).to.equal(index + 1);
            expect(colCache._n2l[index + 1]).to.equal(letter);
        });

        // next level
        expect(colCache.n2l(27)).to.equal("AA");
        expect(colCache._l2n.AB).to.equal(28);
        expect(colCache._n2l[28]).to.equal("AB");
    });

    it("converts numbers to letters", () => {
        expect(colCache.n2l(1)).to.equal("A");
        expect(colCache.n2l(26)).to.equal("Z");
        expect(colCache.n2l(27)).to.equal("AA");
        expect(colCache.n2l(702)).to.equal("ZZ");
        expect(colCache.n2l(703)).to.equal("AAA");
    });
    it("converts letters to numbers", () => {
        expect(colCache.l2n("A")).to.equal(1);
        expect(colCache.l2n("Z")).to.equal(26);
        expect(colCache.l2n("AA")).to.equal(27);
        expect(colCache.l2n("ZZ")).to.equal(702);
        expect(colCache.l2n("AAA")).to.equal(703);
    });

    it("throws when out of bounds", () => {
        expect(() => {
            colCache.n2l(0);
        }).to.throw(Error);
        expect(() => {
            colCache.n2l(-1);
        }).to.throw(Error);
        expect(() => {
            colCache.n2l(16385);
        }).to.throw(Error);

        expect(() => {
            colCache.l2n("");
        }).to.throw(Error);
        expect(() => {
            colCache.l2n("AAAA");
        }).to.throw(Error);
        expect(() => {
            colCache.l2n(16385);
        }).to.throw(Error);
    });

    it("validates addresses properly", () => {
        expect(colCache.validateAddress("A1")).to.be.ok();
        expect(colCache.validateAddress("AA10")).to.be.ok();
        expect(colCache.validateAddress("ABC100000")).to.be.ok();

        expect(() => {
            colCache.validateAddress("A");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("1");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("1A");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("A 1");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("A1A");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("1A1");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("a1");
        }).to.throw(Error);
        expect(() => {
            colCache.validateAddress("a");
        }).to.throw(Error);
    });

    it("decodes addresses", () => {
        expect(colCache.decodeAddress("A1")).to.deep.equal({
            address: "A1",
            col: 1,
            row: 1,
            $col$row: "$A$1",
        });
        expect(colCache.decodeAddress("AA11")).to.deep.equal({
            address: "AA11",
            col: 27,
            row: 11,
            $col$row: "$AA$11",
        });
    });

    describe("with a malformed address", () => {
        it("decodes a side of a reference without row number (a column)", () => {
            expect(colCache.decodeReference("$B")).to.deep.equal({
                address: "B",
                col: 2,
                row: undefined,
                $col$row: "$B$",
            });
        });

        it("decodes a side of a reference without column letters (a row)", () => {
            expect(colCache.decodeReference("$2")).to.deep.equal({
                address: "2",
                col: undefined,
                row: 2,
                $col$row: "$$2",
            });
        });

        it("rejects an address without row number or column letters (it gave undefined, typed number)", () => {
            for (const value of ["$B", "B", "$2", "12", "", "$", "a1"]) {
                expect(() => colCache.decodeAddress(value), value).to.throw(Error, `Invalid Address: ${value}`);
            }
        });
    });

    it("freezes the decoded addresses it shares (the cached top-left cells)", () => {
        const shared = colCache.decodeAddress("C3");
        expect(colCache.decodeAddress("$C$3")).to.equal(shared);
        expect(Object.isFrozen(shared)).to.equal(true);
        expect(Reflect.set(shared, "row", 4)).to.equal(false);
        expect(shared.row).to.equal(3);
        // an address out of the cache is a new object of the caller
        expect(Object.isFrozen(colCache.decodeAddress("C300"))).to.equal(false);
    });

    it("rejects a column number that is not an integer (it gave undefined, typed string)", () => {
        for (const n of [1.5, NaN]) {
            expect(() => colCache.n2l(n), String(n)).to.throw(Error, "is not a column number");
        }
    });

    it("convert [sheetName!][$]col[$]row[[$]col[$]row] into address or range structures", () => {
        expect(colCache.decodeEx("Sheet1!$H$1")).to.deep.equal({
            $col$row: "$H$1",
            address: "H1",
            col: 8,
            row: 1,
            sheetName: "Sheet1",
        });
        expect(colCache.decodeEx("'Sheet 1'!$H$1")).to.deep.equal({
            $col$row: "$H$1",
            address: "H1",
            col: 8,
            row: 1,
            sheetName: "Sheet 1",
        });
        expect(colCache.decodeEx("'Sheet !$:1'!$H$1")).to.deep.equal({
            $col$row: "$H$1",
            address: "H1",
            col: 8,
            row: 1,
            sheetName: "Sheet !$:1",
        });
        expect(colCache.decodeEx("'Sheet !$:1'!#REF!")).to.deep.equal({
            sheetName: "Sheet !$:1",
            error: "#REF!",
        });
    });

    it("gets address structures (and caches them)", () => {
        let addr = colCache.getAddress("D5");
        expect(addr.address).to.equal("D5");
        expect(addr.row).to.equal(5);
        expect(addr.col).to.equal(4);
        expect(colCache.getAddress("D5")).to.equal(addr);
        expect(colCache.getAddress(5, 4)).to.equal(addr);

        addr = colCache.getAddress("E4");
        expect(addr.address).to.equal("E4");
        expect(addr.row).to.equal(4);
        expect(addr.col).to.equal(5);
        expect(colCache.getAddress("E4")).to.equal(addr);
        expect(colCache.getAddress(4, 5)).to.equal(addr);
    });

    it("throws for a column 0 given by number (it decoded the row number as an address)", () => {
        expect(() => colCache.getAddress(5, 0)).to.throw("0 is out of bounds");
    });

    it("decodes whole rows and whole columns as the rectangles they cover (print titles; they gave NaN bounds)", () => {
        const rows = colCache.decodeEx("Sheet1!$2:$1");
        expect(rows).to.deep.include({ sheetName: "Sheet1", top: 1, left: 1, bottom: 2, right: 16384, dimensions: "A1:XFD2" });
        expect(rows.tl).to.deep.equal({ address: "A1", col: 1, row: 1, $col$row: "$A$1", sheetName: "Sheet1" });
        expect(rows.br).to.deep.equal({ address: "XFD2", col: 16384, row: 2, $col$row: "$XFD$2", sheetName: "Sheet1" });
        const columns = colCache.decodeEx("'My Sheet'!$A:$B");
        expect(columns).to.deep.include({ sheetName: "My Sheet", top: 1, left: 1, bottom: 1048576, right: 2, dimensions: "A1:B1048576" });
        expect(colCache.decode("C:B")).to.deep.include({ top: 1, left: 2, bottom: 1048576, right: 3, dimensions: "B1:C1048576" });
        expect(colCache.decode("3:3")).to.deep.include({ top: 3, left: 1, bottom: 3, right: 16384, dimensions: "A3:XFD3" });
        expect(colCache.decode("A1:B2:C3")).to.include({ address: "A1" });
        expect(colCache.decodeEx("A1:B2:C3")).to.include({ dimensions: "A1:B2" });
    });

    it("rejects a range whose sides are of different kinds, or neither a row nor a column", () => {
        for (const value of ["A1:3", "A:3", "1:B2", "A1:", ":", "Sheet1!$A$1:$"]) {
            expect(() => colCache.decodeEx(value), value).to.throw(Error, "Invalid range");
        }
        expect(() => colCache.decode("A:1"), "A:1").to.throw(Error, "Invalid range");
    });

    it("decodes an error reference without sheet (#REF! was read as the sheet #REF)", () => {
        expect(colCache.decodeEx("#REF!")).to.deep.equal({ error: "#REF!" });
        expect(colCache.decodeEx("Sheet1!#REF!")).to.deep.equal({ sheetName: "Sheet1", error: "#REF!" });
    });

    it("decodes addresses and ranges", () => {
        // address
        expect(colCache.decode("A1")).to.deep.equal({
            address: "A1",
            col: 1,
            row: 1,
            $col$row: "$A$1",
        });
        expect(colCache.decode("AA11")).to.deep.equal({
            address: "AA11",
            col: 27,
            row: 11,
            $col$row: "$AA$11",
        });

        // range
        expect(colCache.decode("A1:B2")).to.deep.equal({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2,
        });

        // wonky ranges
        expect(colCache.decode("A2:B1")).to.deep.equal({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2,
        });
        expect(colCache.decode("B2:A1")).to.deep.equal({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2,
        });
    });
});
