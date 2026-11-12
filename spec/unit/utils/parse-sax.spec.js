const parseSax = verquire("utils/parse-sax");

// 2-byte (é), 3-byte (あ, €) and 4-byte (😀, 𝄞) UTF-8 sequences
const TEXT = "é-あ-€-😀-𝄞";
const XML = `<?xml version="1.0" encoding="UTF-8"?><root a="${TEXT}"><t>${TEXT}</t></root>`;

async function* chunks(buffer, size) {
    for (let i = 0; i < buffer.length; i += size) {
        // give other parses a chance to run between chunks

        await new Promise((resolve) => setImmediate(resolve));
        yield buffer.subarray(i, i + size);
    }
}

async function parse(iterable) {
    let text = "";
    let attribute;
    for await (const events of parseSax(iterable)) {
        for (const { eventType, value } of events) {
            if (eventType === "text") text += value;
            if (eventType === "opentag" && value.name === "root") attribute = value.attributes.a;
        }
    }
    return { text, attribute };
}

describe("parseSax", () => {
    const buffer = Buffer.from(XML, "utf8");

    it("rejects a document that ends before its root element is closed", async () => {
        for (const xml of ["<root><a>text", "<root><a", "<root>", '<?xml version="1.0"?>', ""]) {
            const error = await parse([Buffer.from(xml)]).then(
                () => undefined,
                (e) => e
            );
            expect(error, JSON.stringify(xml)).to.be.an("error");
            expect(error.message, JSON.stringify(xml)).to.match(/unclosed tag|unexpected end|root element/);
        }
        expect(await parse([Buffer.from("<root/>  ")])).to.deep.equal({ text: "  ", attribute: undefined });
    });

    [1, 2, 3, 5, 7].forEach((size) => {
        it(`decodes multi-byte characters split across ${size}-byte chunks`, async () => {
            const result = await parse(chunks(buffer, size));
            expect(result).to.deep.equal({ text: TEXT, attribute: TEXT });
        });
    });

    it("splits inside every 2, 3 and 4-byte sequence", async () => {
        const start = buffer.indexOf(Buffer.from(TEXT, "utf8"));
        for (let offset = start; offset < start + Buffer.byteLength(TEXT); offset++) {
            const iterable = [buffer.subarray(0, offset), buffer.subarray(offset)];

            const result = await parse(iterable);
            expect(result.attribute).to.equal(TEXT, `split at byte ${offset}`);
        }
    });

    it("keeps concurrent parses independent", async () => {
        const other = Buffer.from('<root a="x"><t>€€€€€€€€</t></root>', "utf8");
        const results = await Promise.all([
            parse(chunks(buffer, 1)),
            parse(chunks(other, 1)),
            parse(chunks(buffer, 2)),
            parse(chunks(other, 2)),
        ]);
        expect(results).to.deep.equal([
            { text: TEXT, attribute: TEXT },
            { text: "€€€€€€€€", attribute: "x" },
            { text: TEXT, attribute: TEXT },
            { text: "€€€€€€€€", attribute: "x" },
        ]);
    });

    describe("nesting depth", () => {
        const nested = (depth) => `${"<a>".repeat(depth)}${"</a>".repeat(depth)}`;

        it("reads elements nested 1000 deep", async () => {
            let opened = 0;
            for await (const events of parseSax([nested(1000)])) {
                opened += events.filter(({ eventType }) => eventType === "opentag").length;
            }
            expect(opened).to.equal(1000);
        });

        it("rejects deeper elements, whatever the chunks", async () => {
            for (const iterable of [[nested(1001)], chunks(Buffer.from(`<root>${"<a/>".repeat(5000)}${nested(1000)}</root>`), 7)]) {
                let error;
                try {
                    // eslint-disable-next-line no-unused-vars
                    for await (const events of parseSax(iterable)) {
                        // parse
                    }
                } catch (err) {
                    error = err;
                }
                expect(error).to.be.an.instanceOf(Error);
                expect(error.message).to.equal("XML elements are nested more than 1000 deep");
            }
        });
    });

    it("accepts Uint8Array and string chunks", async () => {
        const bytes = new Uint8Array(buffer);
        const result = await parse([bytes.subarray(0, 50), bytes.subarray(50), "<!-- trailing -->"]);
        expect(result).to.deep.equal({ text: TEXT, attribute: TEXT });
    });
});
