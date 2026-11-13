const ExcelJS = verquire("exceljs");

const { headerFooter } = ExcelJS;

describe("headerFooter helpers", () => {
    describe("build", () => {
        it("writes the sections in the order left, center, right, leaving out the absent and empty ones", () => {
            expect(headerFooter.build({ right: "r", left: "l", center: "c" })).to.equal("&Ll&Cc&Rr");
            expect(headerFooter.build({ center: "Title" })).to.equal("&CTitle");
            expect(headerFooter.build({ left: "", center: undefined, right: "r" })).to.equal("&Rr");
            expect(headerFooter.build({})).to.equal("");
        });

        it("turns the fields of Excel's dialog into their codes", () => {
            expect(
                headerFooter.build({
                    left: "&[Page]/&[Pages]",
                    center: "&[Date] &[Time]",
                    right: "&[Path]&[File] (&[Tab])",
                })
            ).to.equal("&L&P/&N&C&D &T&R&Z&F (&A)");
            expect(headerFooter.build({ right: "Page &[Page]+1 of &[Pages]" })).to.equal("&RPage &P+1 of &N");
        });

        it("doubles the literal ampersands", () => {
            expect(headerFooter.build({ center: "Fish & Chips" })).to.equal("&CFish && Chips");
            expect(headerFooter.build({ center: "&" })).to.equal("&C&&");
            expect(headerFooter.build({ center: "a&" })).to.equal("&Ca&&");
            expect(headerFooter.build({ center: "&&&" })).to.equal("&C&&&&");
            expect(headerFooter.build({ center: "x&;y&=z" })).to.equal("&Cx&&;y&&=z");
            expect(headerFooter.build({ center: "&p &b" })).to.equal("&C&&p &&b");
            expect(headerFooter.build({ center: "&[Unknown] &[page] &[Page" })).to.equal("&C&&[Unknown] &&[page] &&[Page");
        });

        it("keeps the formatting codes, the field codes and the escaped ampersands", () => {
            expect(headerFooter.build({ left: '&B&I&"Arial,Bold Italic"&14&KFF0000Red&B' })).to.equal(
                '&L&B&I&"Arial,Bold Italic"&14&KFF0000Red&B'
            );
            expect(headerFooter.build({ center: "&U&E&S&X&Y&O&H&K05+000x" })).to.equal("&C&U&E&S&X&Y&O&H&K05+000x");
            expect(headerFooter.build({ right: "Page &P of &N, &D &T, &Z&F, &A" })).to.equal("&RPage &P of &N, &D &T, &Z&F, &A");
            expect(headerFooter.build({ center: "Tom && Jerry" })).to.equal("&CTom && Jerry");
        });

        it("keeps a text through escape as it is shown", () => {
            expect(headerFooter.escape("AT&T")).to.equal("AT&&T");
            expect(headerFooter.escape("&[Page] && &")).to.equal("&&[Page] &&&& &&");
            expect(headerFooter.escape("no ampersand")).to.equal("no ampersand");
            expect(headerFooter.build({ center: `${headerFooter.escape("R&D, AT&T, &[Page]")} &[Page]` })).to.equal(
                "&CR&&D, AT&&T, &&[Page] &P"
            );
        });

        it("rejects a picture, which the library does not write", () => {
            expect(() => headerFooter.build({ left: "&[Picture]" })).to.throw(
                Error,
                "Unsupported header/footer field in left: &[Picture] (pictures are not written)"
            );
            expect(() => headerFooter.build({ right: "x&G" })).to.throw(
                Error,
                "Unsupported header/footer code in right: &G (pictures are not written)"
            );
            // an escaped one is text
            expect(headerFooter.build({ center: "&&[Picture]&&G" })).to.equal("&C&&[Picture]&&G");
        });

        it("rejects a section code inside a section", () => {
            for (const code of ["L", "C", "R"]) {
                expect(() => headerFooter.build({ center: `a&${code}b` })).to.throw(
                    Error,
                    `Invalid header/footer code in center: &${code} (give the sections as left, center and right)`
                );
            }
        });

        it("rejects a section that is not a string", () => {
            for (const value of [1, null, {}, ["a"]]) {
                expect(() => headerFooter.build({ left: value })).to.throw(
                    Error,
                    "Invalid header/footer section left: a string is expected"
                );
            }
        });

        it("rejects a text longer than 255 characters (the limit of Excel)", () => {
            expect(headerFooter.build({ center: "x".repeat(253) })).to.have.length(255);
            expect(() => headerFooter.build({ center: "x".repeat(254) })).to.throw(
                Error,
                "Header/footer too long: 256 characters (Excel allows at most 255)"
            );
            // the length of the text written: the escaped ampersands and the codes count
            expect(() => headerFooter.build({ center: "a&".repeat(85) })).to.throw(Error, /^Header\/footer too long: 257 characters/);
            expect(headerFooter.build({ center: "&[Pages]".repeat(126) })).to.have.length(254);
        });

        it("is linear in the size of the text", () => {
            const started = Date.now();
            for (const text of ["&[".repeat(100000), "&[Page".repeat(50000), "&".repeat(200000), "&[Pages]".repeat(30000)]) {
                expect(() => headerFooter.build({ center: text })).to.throw(Error, /^Header\/footer too long/);
            }
            expect(Date.now() - started).to.be.below(500);
        });
    });

    describe("parse", () => {
        it("splits the sections, the text before any section code being the center", () => {
            expect(headerFooter.parse("&Ll&Cc&Rr")).to.deep.equal({ left: "l", center: "c", right: "r" });
            expect(headerFooter.parse("Page 1")).to.deep.equal({ center: "Page 1" });
            expect(headerFooter.parse("before&Rright")).to.deep.equal({ center: "before", right: "right" });
            expect(headerFooter.parse("&Rr&Ll")).to.deep.equal({ left: "l", right: "r" });
            expect(headerFooter.parse("&La&Cb&Lc")).to.deep.equal({ left: "ac", center: "b" });
            expect(headerFooter.parse("&L&C&R")).to.deep.equal({});
            expect(headerFooter.parse("")).to.deep.equal({});
        });

        it("gives the field codes as the fields of Excel's dialog", () => {
            expect(headerFooter.parse("&LPage &P of &N&C&D &T&R&Z&F &A")).to.deep.equal({
                left: "Page &[Page] of &[Pages]",
                center: "&[Date] &[Time]",
                right: "&[Path]&[File] &[Tab]",
            });
            expect(headerFooter.parse("&G")).to.deep.equal({ center: "&[Picture]" });
        });

        it("keeps the escaped ampersands and the formatting codes", () => {
            expect(headerFooter.parse('&L&"Algerian,Italic"&14&S&U&X&K05+000Formatted&CTom && Jerry&R&&P')).to.deep.equal({
                left: '&"Algerian,Italic"&14&S&U&X&K05+000Formatted',
                center: "Tom && Jerry",
                right: "&&P",
            });
            expect(headerFooter.parse("&&&&L")).to.deep.equal({ center: "&&&&L" });
            expect(headerFooter.parse("a&")).to.deep.equal({ center: "a&" });
        });

        it("keeps the line breaks", () => {
            expect(headerFooter.parse("&L&P of &N&C&D &T\n&Z&F\n&A")).to.deep.equal({
                left: "&[Page] of &[Pages]",
                center: "&[Date] &[Time]\n&[Path]&[File]\n&[Tab]",
            });
        });

        it("gives back what build takes", () => {
            for (const sections of [
                { left: "Fish && Chips", center: "&[Page] of &[Pages]", right: '&B&"Arial,Bold"&12Title' },
                { center: "&KFF0000Red &[Date] &[Time]" },
                { right: "&&[Page]" },
            ]) {
                const text = headerFooter.build(sections);
                expect(headerFooter.parse(text)).to.deep.equal(sections);
                expect(headerFooter.build(headerFooter.parse(text))).to.equal(text);
            }
        });

        it("reads any length, in linear time", () => {
            const started = Date.now();
            expect(headerFooter.parse("&".repeat(200000)).center).to.have.length(200000);
            expect(headerFooter.parse("&L".repeat(100000))).to.deep.equal({});
            expect(headerFooter.parse("&P".repeat(100000)).center).to.have.length(700000);
            expect(Date.now() - started).to.be.below(500);
        });
    });

    it("sets a header written to and read from a file", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("s");
        ws.headerFooter.oddHeader = headerFooter.build({ left: "Q&&A", right: "Page &[Page] of &[Pages]" });
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(await wb.xlsx.writeBuffer());
        const { oddHeader } = loaded.getWorksheet("s").headerFooter;
        expect(oddHeader).to.equal("&LQ&&A&RPage &P of &N");
        expect(headerFooter.parse(oddHeader)).to.deep.equal({ left: "Q&&A", right: "Page &[Page] of &[Pages]" });
    });
});
