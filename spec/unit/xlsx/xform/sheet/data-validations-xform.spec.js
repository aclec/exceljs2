const testXformHelper = require("../test-xform-helper");

const DataValidationsXform = verquire("xlsx/xform/sheet/data-validations-xform");
const parseSax = verquire("utils/parse-sax");

async function parse(xml) {
    const xform = new DataValidationsXform();
    return xform.parse(parseSax([xml]));
}

const expectations = [
    {
        title: "list type",
        create: () => new DataValidationsXform(),
        preparedModel: {
            E1: {
                type: "list",
                allowBlank: true,
                showInputMessage: true,
                showErrorMessage: true,
                formulae: ["Ducks"],
            },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="1">
        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1">
          <formula1>Ducks</formula1>
        </dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "whole type",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: {
                type: "whole",
                operator: "between",
                allowBlank: true,
                showInputMessage: true,
                showErrorMessage: true,
                formulae: [5, 10],
            },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "decimal type",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: {
                type: "decimal",
                operator: "notBetween",
                allowBlank: true,
                showInputMessage: true,
                showErrorMessage: true,
                formulae: [5, 10],
            },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="1">
        <dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "custom type",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: {
                type: "custom",
                allowBlank: true,
                showInputMessage: true,
                showErrorMessage: true,
                formulae: ["OR(C21=5,C21=7)"],
            },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="1">
        <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>OR(C21=5,C21=7)</formula1>
        </dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "parse open office",
        create: () => new DataValidationsXform(),
        xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="true" showInputMessage="false" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
        parsedModel: {
            A1: {
                type: "whole",
                operator: "between",
                allowBlank: true,
                showInputMessage: false,
                formulae: [5, 10],
            },
        },
        tests: ["parse"],
    },
    {
        title: "optimised",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: { type: "whole", operator: "between", formulae: [5, 10] },
            A2: { type: "whole", operator: "between", formulae: [5, 10] },
            B1: { type: "whole", operator: "between", formulae: [5, 10] },
            B2: { type: "whole", operator: "between", formulae: [5, 10] },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="1">
        <dataValidation type="whole" sqref="A1:B2">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
        tests: ["render", "parse"],
    },
    {
        title: "formulas that are not literal values",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: { type: "whole", operator: "greaterThan", formulae: ["$B$1"] },
            A2: { type: "decimal", operator: "between", formulae: ["A1+1", "MAX(C:C)"] },
            A3: { type: "textLength", operator: "lessThan", formulae: ["LEN(B3)"] },
            A4: { type: "date", operator: "greaterThan", formulae: ["TODAY()"] },
        },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: `
      <dataValidations count="4">
        <dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>$B$1</formula1></dataValidation>
        <dataValidation type="decimal" sqref="A2"><formula1>A1+1</formula1><formula2>MAX(C:C)</formula2></dataValidation>
        <dataValidation type="textLength" operator="lessThan" sqref="A3"><formula1>LEN(B3)</formula1></dataValidation>
        <dataValidation type="date" operator="greaterThan" sqref="A4"><formula1>TODAY()</formula1></dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "a date",
        create: () => new DataValidationsXform(),
        preparedModel: { A1: { type: "date", operator: "lessThan", formulae: [new Date(Date.UTC(2016, 0, 1))] } },
        // the serial date is read as a Date by reconcile, which knows the date system of the workbook
        parsedModel: { A1: { type: "date", operator: "lessThan", formulae: ["42370"] } },
        get reconciledModel() {
            return this.preparedModel;
        },
        options: {},
        xml: `
      <dataValidations count="1">
        <dataValidation type="date" operator="lessThan" sqref="A1"><formula1>42370</formula1></dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn", "parse", "reconcile"],
    },
    {
        // an ISO date string was converted to a date before 5.0: written as a formula it would be 2016 - 1 - 1
        title: "a date given as an ISO 8601 string",
        create: () => new DataValidationsXform(),
        preparedModel: {
            A1: { type: "date", operator: "lessThan", formulae: ["2016-01-01"] },
            A2: { type: "date", operator: "lessThan", formulae: ["2016-01-01T00:00:00Z"] },
        },
        xml: `
      <dataValidations count="2">
        <dataValidation type="date" operator="lessThan" sqref="A1"><formula1>42370</formula1></dataValidation>
        <dataValidation type="date" operator="lessThan" sqref="A2"><formula1>42370</formula1></dataValidation>
      </dataValidations>
    `,
        tests: ["render", "renderIn"],
    },
];

describe("DataValidationsXform", () => {
    testXformHelper(expectations);

    it("rejects a key that is not a cell address or a range in its normal form", () => {
        // these keys made the writer fail with a TypeError
        for (const key of ["$A$1:B2", "B2:A1", "a1", "foo", "#REF!", "Sheet1!#REF!"]) {
            const xform = new DataValidationsXform();
            expect(() => xform.toXml({ [key]: { type: "list", formulae: ['"a"'] } }), key).to.throw(/Invalid data validation address/);
        }
    });
});

describe("DataValidationsXform formulae", () => {
    it("reads the number literals of whole, decimal and textLength rules as numbers", async () => {
        const model = await parse(
            '<dataValidations count="3">' +
                '<dataValidation type="decimal" sqref="A1"><formula1>-1.5E-3</formula1><formula2>+.5</formula2></dataValidation>' +
                '<dataValidation type="whole" sqref="A2"><formula1>-7</formula1><formula2>1e3</formula2></dataValidation>' +
                '<dataValidation type="textLength" sqref="A3"><formula1>2.</formula1><formula2>0x10</formula2></dataValidation>' +
                "</dataValidations>"
        );
        expect(model.A1.formulae).to.deep.equal([-0.0015, 0.5]);
        expect(model.A2.formulae).to.deep.equal([-7, 1000]);
        // 0x10 is not a number literal of a formula
        expect(model.A3.formulae).to.deep.equal([2, "0x10"]);
    });

    it("reads a formula of a number rule in linear time", async () => {
        // a long run of digits that is not a number literal in the end: the literal test must not backtrack
        const formula = `${"1".repeat(100000)}x`;
        const start = Date.now();
        const model = await parse(
            `<dataValidations count="1"><dataValidation type="decimal" sqref="A1"><formula1>${formula}</formula1></dataValidation></dataValidations>`
        );
        expect(Date.now() - start).to.be.below(200);
        expect(model.A1.formulae).to.deep.equal([formula]);
    });

    it("ignores a formula outside of a <dataValidation>", async () => {
        // was a TypeError (or a formula added to the previous validation)
        const model = await parse(
            '<dataValidations count="1"><formula1>1</formula1>' +
                '<dataValidation type="whole" sqref="A1"><formula1>5</formula1></dataValidation><formula2>9</formula2></dataValidations>'
        );
        expect(model).to.deep.equal({ A1: { type: "whole", operator: "between", formulae: [5] } });
    });

    it("reads an unknown type as any, and an unknown operator as between", async () => {
        const model = await parse(
            '<dataValidations count="4">' +
                '<dataValidation type="none" showInputMessage="1" prompt="p" sqref="A1"/>' +
                '<dataValidation type="integer" showInputMessage="1" prompt="p" sqref="A2"/>' +
                '<dataValidation type="whole" operator="greaterThanOrEquals" sqref="A3"><formula1>5</formula1></dataValidation>' +
                '<dataValidation type="time" operator="lessThan" sqref="A4"><formula1>0.5</formula1></dataValidation>' +
                "</dataValidations>"
        );
        expect(model).to.deep.equal({
            A1: { type: "any", showInputMessage: true, prompt: "p" },
            A2: { type: "any", showInputMessage: true, prompt: "p" },
            A3: { type: "whole", operator: "between", formulae: [5] },
            // ST_DataValidationType time: its serial times are kept as text
            A4: { type: "time", operator: "lessThan", formulae: ["0.5"] },
        });
    });

    it("rejects a date or a number that is not a value", () => {
        const cases = [
            [{ type: "date", formulae: [new Date(NaN)] }, /Invalid date/],
            [{ type: "date", formulae: [NaN] }, /Invalid date/],
            [{ type: "whole", formulae: [NaN] }, /not a finite number/],
            [{ type: "decimal", formulae: [Infinity] }, /not a finite number/],
            // date strings that 4.x converted with new Date(): written as formulas they would be wrong (2020/01/05 is 404)
            [{ type: "date", formulae: ["2020/01/05"] }, /not an ISO 8601 date/],
            [{ type: "date", formulae: ["2020-1-5"] }, /not an ISO 8601 date/],
            [{ type: "date", formulae: ["01.05.2020"] }, /not an ISO 8601 date/],
            [{ type: "date", formulae: ["January 5, 2020"] }, /not an ISO 8601 date/],
            [{ type: "date", formulae: ["5 Jan 2020"] }, /not an ISO 8601 date/],
            [{ type: "date", formulae: [String(new Date(Date.UTC(2020, 0, 5)))] }, /not an ISO 8601 date/],
        ];
        for (const [dataValidation, error] of cases) {
            const xform = new DataValidationsXform();
            expect(() => xform.toXml({ A1: dataValidation }), String(dataValidation.formulae[0])).to.throw(Error, error);
        }
    });

    it("squeezes the cells of a column or a block into ranges that do not overlap (rows beyond 9 included)", () => {
        const dv = { type: "list", formulae: ['"a,b"'] };
        const other = { type: "whole", operator: "between", formulae: [1, 2] };
        const cells = (from, to, cols, validation) => {
            const model = {};
            for (let r = from; r <= to; r++) {
                for (const c of cols) {
                    model[`${c}${r}`] = validation;
                }
            }
            return model;
        };
        const sqrefs = (model) => [...new DataValidationsXform().toXml(model).matchAll(/sqref="([^"]*)"/g)].map((m) => m[1]).sort();

        expect(sqrefs(cells(2, 20, ["G"], dv))).to.deep.equal(["G2:G20"]);
        expect(sqrefs(cells(3, 100, ["H"], dv))).to.deep.equal(["H3:H100"]);
        expect(sqrefs(cells(9, 12, ["A", "B"], dv))).to.deep.equal(["A9:B12"]);
        // two validations side by side, and a block with a hole: every cell is in exactly one range
        const mixed = { ...cells(1, 12, ["A"], dv), ...cells(1, 12, ["B"], other), ...cells(1, 12, ["C"], dv) };
        delete mixed.C5;
        expect(sqrefs(mixed)).to.deep.equal(["A1:A12", "B1:B12", "C1:C4", "C6:C12"]);
        // the last column and the last row of the grid
        expect(sqrefs({ XFD1: dv, XFD2: dv, A1048576: dv, XFD1048576: dv })).to.deep.equal(["A1048576", "XFD1048576", "XFD1:XFD2"]);
    });

    it("writes the formulas of a date rule that are not dates as they are", () => {
        for (const formula of ["TODAY()", "$B$1", "DATE(2020,1,5)", "B1+30", "Sheet1!A1", "mayday"]) {
            const xml = new DataValidationsXform().toXml({ A1: { type: "date", operator: "greaterThan", formulae: [formula] } });
            expect(xml, formula).to.include(`<formula1>${formula}</formula1>`);
        }
    });
});
