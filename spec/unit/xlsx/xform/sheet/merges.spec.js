const Merges = verquire("xlsx/xform/sheet/merges");

describe("Merges", () => {
    it("collects the merged ranges by master cell", () => {
        const merges = new Merges();
        merges.add({ address: "B1", master: "A1" });
        merges.add({ address: "A2", master: "A1" });
        merges.add({ address: "B2", master: "A1" });
        merges.add({ address: "D4", master: "C3" });
        expect(merges.mergeCells).to.deep.equal(["A1:B2", "C3:D4"]);
    });

    it("finds the master address of a merged cell", () => {
        const merges = new Merges();
        merges.add({ address: "B2", master: "A1" });
        merges.add({ address: "D4", master: "C3" });
        expect(merges.getMasterAddress("B1")).to.equal("A1");
        expect(merges.getMasterAddress("A1")).to.equal("A1");
        expect(merges.getMasterAddress("D3")).to.equal("C3");
        expect(merges.getMasterAddress("E5")).to.equal(undefined);
    });
});
