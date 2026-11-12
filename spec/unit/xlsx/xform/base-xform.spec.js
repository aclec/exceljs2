const BaseXform = verquire("xlsx/xform/base-xform");

describe("BaseXform", () => {
    describe("toBoolValue (xsd:boolean)", () => {
        it('reads "1" and "true" as true', () => {
            expect(BaseXform.toBoolValue("1")).to.equal(true);
            expect(BaseXform.toBoolValue("true")).to.equal(true);
            expect(BaseXform.toBoolValue("true", false)).to.equal(true);
        });

        it('reads "0" and "false" as false', () => {
            expect(BaseXform.toBoolValue("0")).to.equal(false);
            expect(BaseXform.toBoolValue("false")).to.equal(false);
            expect(BaseXform.toBoolValue("false", true)).to.equal(false);
        });

        it("gives the default when the attribute is absent", () => {
            expect(BaseXform.toBoolValue(undefined)).to.equal(undefined);
            expect(BaseXform.toBoolValue(undefined, true)).to.equal(true);
        });
    });
});
