const { parseIsoDate } = verquire("utils/iso-date");

const iso = (text) => {
    const date = parseIsoDate(text);
    return date && date.toISOString();
};

describe("parseIsoDate", () => {
    it("reads dates and date-times without an offset as UTC", () => {
        expect(iso("2024-02-29")).to.equal("2024-02-29T00:00:00.000Z");
        expect(iso("2024-02-29T13:14")).to.equal("2024-02-29T13:14:00.000Z");
        expect(iso("2024-02-29T13:14:15")).to.equal("2024-02-29T13:14:15.000Z");
        expect(iso("2024-02-29T13:14:15.5")).to.equal("2024-02-29T13:14:15.500Z");
        expect(iso("2024-02-29T13:14:15.123456")).to.equal("2024-02-29T13:14:15.123Z");
        expect(iso("2024-02-29T13:14:15.007")).to.equal("2024-02-29T13:14:15.007Z");
    });

    it("applies Z and +hh:mm / -hh:mm offsets", () => {
        expect(iso("2024-02-29T13:14:15Z")).to.equal("2024-02-29T13:14:15.000Z");
        expect(iso("2024-02-29T13:14:15+02:00")).to.equal("2024-02-29T11:14:15.000Z");
        expect(iso("2024-02-29T01:00-05:30")).to.equal("2024-02-29T06:30:00.000Z");
        expect(iso("2024-03-01T00:30+01:00")).to.equal("2024-02-29T23:30:00.000Z");
        expect(iso("2024-02-29T13:14+14:00")).to.equal("2024-02-28T23:14:00.000Z");
    });

    it("reads T24:00 as the start of the next day (ISO 8601 / XML Schema end of day)", () => {
        expect(iso("2024-02-29T24:00")).to.equal("2024-03-01T00:00:00.000Z");
        expect(iso("2024-12-31T24:00:00.000Z")).to.equal("2025-01-01T00:00:00.000Z");
        expect(iso("2024-02-29T24:00:00+01:00")).to.equal("2024-02-29T23:00:00.000Z");
        expect(parseIsoDate("2024-02-29T24:00:01")).to.equal(undefined);
        expect(parseIsoDate("2024-02-29T24:01")).to.equal(undefined);
        expect(parseIsoDate("2024-02-29T24:00:00.5")).to.equal(undefined);
    });

    it("keeps years 0-99 (no mapping to 1900-1999)", () => {
        expect(iso("0099-12-31")).to.equal("0099-12-31T00:00:00.000Z");
        expect(iso("0000-01-01T12:00")).to.equal("0000-01-01T12:00:00.000Z");
        expect(iso("0050-06-15T10:00+01:00")).to.equal("0050-06-15T09:00:00.000Z");
    });

    it("returns undefined for anything else", () => {
        [
            undefined,
            null,
            42,
            "",
            "not a date",
            "2024-02-30",
            "2023-02-29",
            "2024-13-01",
            "2024-00-10",
            "2024-02-29T25:00",
            "2024-02-29T12:60",
            "2024-02-29T12:30:60",
            "2024-02-29T12:30+15:00",
            "2024-02-29T12:30+02:60",
            "2024-02-29T12:30+0200",
            "2024-02-29T",
            "2024-02-29 12:30",
            "2024-2-29",
            "+002024-02-29",
            " 2024-02-29",
            "2024-02-29Z",
        ].forEach((text) => expect(parseIsoDate(text), String(text)).to.equal(undefined));
    });
});
