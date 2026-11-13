// The authors of the notes (comments): <authors> of the comments part, <comment authorId>, Comment.author.
const { PassThrough } = require("stream");
const { baseWorkbookBuffer, readPart, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb.getWorksheet("Sheet1");
}

// the notes of a streamed workbook, written into memory
async function streamBuffer(build) {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    const ws = wb.addWorksheet("Sheet1");
    build(ws);
    ws.commit();
    await wb.commit();
    return Buffer.concat(chunks);
}

function authorsOf(xml) {
    const authors = /<authors>(.*)<\/authors>/.exec(xml);
    return [...authors[1].matchAll(/<author>([^<]*)<\/author>|<author\/>/g)].map((match) => match[1] ?? "");
}

function authorIdsOf(xml) {
    return [...xml.matchAll(/<comment ref="([^"]+)" authorId="([^"]+)"/g)].map((match) => `${match[1]}:${match[2]}`);
}

// a workbook whose comments part has the given <authors> content and authorId on its notes (B1, B2, ...)
async function withAuthors(authorsXml, authorIds) {
    const buffer = await baseWorkbookBuffer((wb, ws) => {
        authorIds.forEach((authorId, index) => {
            ws.getCell(`B${index + 1}`).note = `note ${index + 1}`;
        });
    });
    return tamper(buffer, {
        "xl/comments1.xml": (xml) => {
            let index = 0;
            return xml
                .replace("<authors><author>Author</author></authors>", `<authors>${authorsXml}</authors>`)
                .replace(/authorId="0"/g, () => `authorId="${authorIds[index++]}"`);
        },
    });
}

function buildAuthoredNotes(ws) {
    ws.getCell("A1").note = { texts: [{ text: "first" }], author: "Alice" };
    ws.getCell("A2").note = "no author";
    ws.getCell("A3").note = { texts: [{ font: { bold: true }, text: "Bob:" }, { text: "\nsecond" }], author: "Bob" };
    ws.getCell("A4").note = { texts: [{ text: "third" }], author: "Alice" };
}

describe("Note authors", () => {
    it("writes the placeholder author when no note has one (unchanged output)", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("A1").note = "one";
            ws.getCell("B2").note = { texts: [{ text: "two" }] };
        });
        const xml = await readPart(buffer, "xl/comments1.xml");
        expect(xml).to.contain("<authors><author>Author</author></authors>");
        expect(authorIdsOf(xml)).to.deep.equal(["A1:0", "B2:0"]);

        // and reads it back as no author: a string note stays a string
        const ws = await load(buffer);
        expect(ws.getCell("A1").note).to.equal("one");
        expect(ws.getCell("B2").note).to.equal("two");
    });

    it("writes the distinct authors in the order of their first note, with the authorId of each note", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => buildAuthoredNotes(ws));
        const xml = await readPart(buffer, "xl/comments1.xml");
        expect(authorsOf(xml)).to.deep.equal(["Alice", "Author", "Bob"]);
        expect(authorIdsOf(xml)).to.deep.equal(["A1:0", "A2:1", "A3:2", "A4:0"]);
        // the text is not changed
        expect(xml).to.contain("<t>first</t>");
    });

    it("reads the author of each note back (round trip)", async () => {
        const ws = await load(await baseWorkbookBuffer((wb, sheet) => buildAuthoredNotes(sheet)));
        expect(ws.getCell("A1").note).to.include({ author: "Alice" });
        expect(ws.getCell("A1").note.texts).to.deep.equal([{ text: "first" }]);
        expect(ws.getCell("A2").note).to.equal("no author");
        expect(ws.getCell("A3").note.author).to.equal("Bob");
        expect(ws.getCell("A3").note.texts).to.have.length(2);
        expect(ws.getCell("A4").note.author).to.equal("Alice");
        expect(ws.getCell("A1").model.comment.note.author).to.equal("Alice");

        // and once more
        const again = await load(await ws.workbook.xlsx.writeBuffer());
        expect(again.getCell("A4").note.author).to.equal("Alice");
        expect(authorsOf(await readPart(await again.workbook.xlsx.writeBuffer(), "xl/comments1.xml"))).to.deep.equal([
            "Alice",
            "Author",
            "Bob",
        ]);
    });

    it("keeps the author when the notes are copied through the model", async () => {
        const ws = await load(await baseWorkbookBuffer((wb, sheet) => buildAuthoredNotes(sheet)));
        const wb2 = new ExcelJS.Workbook();
        wb2.model = ws.workbook.model;
        expect(wb2.getWorksheet("Sheet1").getCell("A3").note.author).to.equal("Bob");
    });

    it("escapes the author (XML and _xHHHH_) and reads it back as written", async () => {
        const author = 'a <b> & "c" _x0041_ é東';
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("A1").note = { texts: [{ text: "n" }], author };
        });
        const xml = await readPart(buffer, "xl/comments1.xml");
        expect(xml).to.contain("<author>a &lt;b&gt; &amp; &quot;c&quot; _x005F_x0041_ é東</author>");
        expect((await load(buffer)).getCell("A1").note.author).to.equal(author);
    });

    it("keeps an empty author", async () => {
        const ws = await load(await withAuthors("<author></author><author>Author</author>", [0, 1]));
        expect(ws.getCell("B1").note).to.deep.include({ author: "", texts: [{ text: "note 1" }] });
        expect(ws.getCell("B2").note).to.equal("note 2");
    });

    it("rejects an author that is not a string", async () => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet("Sheet1").getCell("A1").note = { texts: [{ text: "n" }], author: 42 };
        let error;
        try {
            await wb.xlsx.writeBuffer();
        } catch (err) {
            error = err;
        }
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("The author of a note must be a string (note of A1)");
    });

    describe("from a file", () => {
        it("resolves the authorId of each note", async () => {
            const ws = await load(await withAuthors("<author>Ann</author><author>Bea</author><author>Cy</author>", [2, 0, 1, 2]));
            expect(["B1", "B2", "B3", "B4"].map((address) => ws.getCell(address).note.author)).to.deep.equal(["Cy", "Ann", "Bea", "Cy"]);
        });

        it("reads an author split by an entity, and its _xHHHH_ escapes", async () => {
            const ws = await load(await withAuthors("<author>T&amp;J _x00E9_</author>", [0]));
            expect(ws.getCell("B1").note.author).to.equal("T&J é");
        });

        it("gives no author for an authorId that is not the index of an author (never throws)", async () => {
            const ids = ["2", "-1", "x", "", "1.0", "1e0", "0x1", " 1", "99999999999999999999999", "constructor", "__proto__", "length"];
            const ws = await load(await withAuthors("<author>Ann</author><author>Bea</author>", ids));
            ids.forEach((id, index) => {
                expect(ws.getCell(`B${index + 1}`).note, `authorId "${id}"`).to.equal(`note ${index + 1}`);
            });
        });

        it("gives no author without <authors>, or with <authors> after the notes", async () => {
            const noAuthors = await tamper(await withAuthors("<author>Ann</author>", [0]), {
                "xl/comments1.xml": (xml) => xml.replace("<authors><author>Ann</author></authors>", ""),
            });
            expect((await load(noAuthors)).getCell("B1").note).to.equal("note 1");

            const authorsLast = await tamper(await withAuthors("<author>Ann</author>", [0]), {
                "xl/comments1.xml": (xml) =>
                    xml
                        .replace("<authors><author>Ann</author></authors>", "")
                        .replace("</comments>", "<authors><author>Ann</author></authors></comments>"),
            });
            expect((await load(authorsLast)).getCell("B1").note).to.equal("note 1");
        });

        it("ignores an <author> outside <authors>, and <authors> elements inside a note", async () => {
            const buffer = await tamper(await withAuthors("<author>Ann</author>", [0, 1]), {
                "xl/comments1.xml": (xml) =>
                    xml
                        .replace("<authors>", "<author>Stray</author><authors>")
                        .replace("<text>", "<authors><author>Inner</author></authors><text>"),
            });
            const ws = await load(buffer);
            expect(ws.getCell("B1").note.author).to.equal("Ann");
            expect(ws.getCell("B2").note).to.equal("note 2");
        });

        it("reads many authors in linear time", async function () {
            this.timeout(20000);
            const count = 100000;
            const authors = Array.from({ length: count }, (_, index) => `<author>A${index}</author>`).join("");
            const ws = await load(await withAuthors(authors, [count - 1, 0, count]));
            expect(ws.getCell("B1").note.author).to.equal(`A${count - 1}`);
            expect(ws.getCell("B2").note.author).to.equal("A0");
            expect(ws.getCell("B3").note).to.equal("note 3");
        });
    });

    describe("streaming writer", () => {
        it("writes the same authors and authorIds as the document writer", async () => {
            const buffer = await streamBuffer(buildAuthoredNotes);
            const xml = await readPart(buffer, "xl/comments1.xml");
            expect(authorsOf(xml)).to.deep.equal(["Alice", "Author", "Bob"]);
            expect(authorIdsOf(xml)).to.deep.equal(["A1:0", "A2:1", "A3:2", "A4:0"]);

            const ws = await load(buffer);
            expect(ws.getCell("A1").note.author).to.equal("Alice");
            expect(ws.getCell("A2").note).to.equal("no author");
            expect(ws.getCell("A3").note.author).to.equal("Bob");
        });

        it("writes the placeholder author when no note has one", async () => {
            const xml = await readPart(
                await streamBuffer((ws) => {
                    ws.getCell("A1").note = "one";
                }),
                "xl/comments1.xml"
            );
            expect(xml).to.contain("<authors><author>Author</author></authors>");
            expect(authorIdsOf(xml)).to.deep.equal(["A1:0"]);
        });

        it("rejects an author that is not a string when the row is committed", async () => {
            const stream = new PassThrough();
            stream.resume();
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
            const ws = wb.addWorksheet("Sheet1");
            ws.getCell("A1").note = { texts: [{ text: "n" }], author: null };
            expect(() => ws.getRow(1).commit()).to.throw(Error, "The author of a note must be a string (note of A1)");
        });
    });
});
