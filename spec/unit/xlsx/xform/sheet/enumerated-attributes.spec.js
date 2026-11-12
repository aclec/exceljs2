// The enumerated attributes of the worksheet elements (ECMA-376 / MS-XLSX simple types) are validated when read: a value
// of the enumeration is kept, an unknown one reads as the default (it was kept as is, against the types, and written back
// as an invalid file).
const parseSax = verquire("utils/parse-sax");
const { cloneDeep } = require("../../../../utils/under-dash");

const PageSetupXform = verquire("xlsx/xform/sheet/page-setup-xform");
const SheetViewXform = verquire("xlsx/xform/sheet/sheet-view-xform");
const CfvoXform = verquire("xlsx/xform/sheet/cf/cfvo-xform");
const CfvoExtXform = verquire("xlsx/xform/sheet/cf-ext/cfvo-ext-xform");
const CfRuleXform = verquire("xlsx/xform/sheet/cf/cf-rule-xform");
const CfRuleExtXform = verquire("xlsx/xform/sheet/cf-ext/cf-rule-ext-xform");
const ConditionalFormattingsXform = verquire("xlsx/xform/sheet/cf/conditional-formattings-xform");
const IconSetXform = verquire("xlsx/xform/sheet/cf/icon-set-xform");
const IconSetExtXform = verquire("xlsx/xform/sheet/cf-ext/icon-set-ext-xform");
const CfIconExtXform = verquire("xlsx/xform/sheet/cf-ext/cf-icon-ext-xform");
const DatabarExtXform = verquire("xlsx/xform/sheet/cf-ext/databar-ext-xform");
const CellXform = verquire("xlsx/xform/sheet/cell-xform");

async function parse(xform, xml) {
    // without the undefined keys
    return cloneDeep(await xform.parse(parseSax([xml])), false);
}

describe("Enumerated attributes of the worksheet elements", () => {
    it("<pageSetup>: orientation, pageOrder, cellComments, errors", async () => {
        const invalid = await parse(
            new PageSetupXform(),
            '<pageSetup orientation="upright" pageOrder="up" cellComments="x" errors="hide"/>'
        );
        expect(invalid).to.deep.include({ orientation: "portrait", pageOrder: "downThenOver", cellComments: "None", errors: "displayed" });
        // ST_Orientation default: the orientation of the printer
        const valid = await parse(
            new PageSetupXform(),
            '<pageSetup orientation="default" pageOrder="overThenDown" cellComments="atEnd" errors="NA"/>'
        );
        expect(valid).to.deep.include({ orientation: "default", pageOrder: "overThenDown", cellComments: "atEnd", errors: "NA" });
    });

    it("<sheetView view> and <pane activePane>", async () => {
        // normal is the default presentation: no style
        expect(await parse(new SheetViewXform(), '<sheetView view="normal" workbookViewId="0"/>')).not.to.have.property("style");
        expect(await parse(new SheetViewXform(), '<sheetView view="zoomed" workbookViewId="0"/>')).not.to.have.property("style");
        expect(await parse(new SheetViewXform(), '<sheetView view="pageLayout" workbookViewId="0"/>')).to.include({ style: "pageLayout" });
        const split = await parse(
            new SheetViewXform(),
            '<sheetView workbookViewId="0"><pane xSplit="2000" activePane="middle"/></sheetView>'
        );
        expect(split).to.include({ state: "split", activePane: "topLeft" });
    });

    it("<cfvo type> and <x14:cfvo type>: num when absent or unknown", async () => {
        expect(await parse(new CfvoXform(), '<cfvo type="median" val="5"/>')).to.deep.equal({ type: "num", value: 5 });
        expect(await parse(new CfvoXform(), '<cfvo val="5"/>')).to.deep.equal({ type: "num", value: 5 });
        expect(await parse(new CfvoXform(), '<cfvo type="percentile" val="50"/>')).to.deep.equal({ type: "percentile", value: 50 });
        expect(await parse(new CfvoExtXform(), '<x14:cfvo type="median"><xm:f>5</xm:f></x14:cfvo>')).to.deep.equal({
            type: "num",
            value: 5,
        });
        expect(await parse(new CfvoExtXform(), '<x14:cfvo type="autoMin"/>')).to.deep.equal({ type: "autoMin" });
    });

    it("<cfRule>: a cellIs rule keeps every operator of ST_ConditionalFormattingOperator", async () => {
        for (const operator of ["containsText", "notContains", "beginsWith", "endsWith", "notBetween"]) {
            const xml = `<cfRule type="cellIs" dxfId="0" priority="1" operator="${operator}"><formula>"a"</formula></cfRule>`;
            expect(await parse(new CfRuleXform(), xml), operator).to.deep.equal({
                type: "cellIs",
                dxfId: 0,
                priority: 1,
                operator,
                formulae: ['"a"'],
            });
        }
    });

    it("<cfRule>: a rule of the variant of its type, an unknown operator or time period is absent", async () => {
        expect(
            await parse(
                new CfRuleXform(),
                '<cfRule type="cellIs" dxfId="0" priority="1" operator="largerThan"><formula>1</formula></cfRule>'
            )
        ).to.deep.equal({ type: "cellIs", dxfId: 0, priority: 1, formulae: ["1"] });
        expect(await parse(new CfRuleXform(), '<cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="nextYear"/>')).to.deep.equal({
            type: "timePeriod",
            dxfId: 0,
            priority: 1,
        });
        expect(await parse(new CfRuleExtXform(), '<x14:cfRule type="cellIs" priority="1" operator="largerThan" id="{A}"/>')).to.deep.equal({
            type: "cellIs",
            priority: 1,
            x14Id: "{A}",
        });
    });

    it("<cfRule>: a rule of a type that the model does not have is not read", async () => {
        // it was kept with its type, and written back as an empty <conditionalFormatting> (a corrupt file)
        expect(await new CfRuleXform().parse(parseSax(['<cfRule type="duplicateValues" dxfId="0" priority="1"/>']))).to.equal(null);
        expect(await new CfRuleExtXform().parse(parseSax(['<x14:cfRule type="uniqueValues" priority="1" id="{A}"/>']))).to.equal(null);
        const xform = new ConditionalFormattingsXform();
        xform.reset();
        await xform.parse(
            parseSax([
                '<conditionalFormatting sqref="A1:A5"><cfRule type="duplicateValues" dxfId="0" priority="1"/>' +
                    '<cfRule type="expression" dxfId="0" priority="2"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>',
            ])
        );
        expect(cloneDeep(xform.model, false)).to.deep.equal([
            { ref: "A1:A5", rules: [{ type: "expression", dxfId: 0, priority: 2, formulae: ["A1>1"] }] },
        ]);
    });

    it("<conditionalFormatting> without sqref is not read (its rules apply to no cell)", async () => {
        const xform = new ConditionalFormattingsXform();
        xform.reset();
        await xform.parse(
            parseSax([
                '<conditionalFormatting><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>',
            ])
        );
        expect(xform.model).to.deep.equal([]);
    });

    it("<cfRule>: the settings of a child of another type (malformed) are not merged", async () => {
        const rule = await parse(
            new CfRuleXform(),
            '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
                "<formula>1</formula></cfRule>"
        );
        expect(rule).to.deep.equal({ type: "cellIs", dxfId: 0, priority: 1, operator: "equal", formulae: ["1"] });
    });

    it("<iconSet iconSet>, <x14:iconSet iconSet> and <x14:cfIcon iconSet>", async () => {
        expect(await parse(new IconSetXform(), '<iconSet iconSet="3Hearts"><cfvo type="percent" val="0"/></iconSet>')).to.deep.equal({
            iconSet: "3TrafficLights1",
            cfvo: [{ type: "percent", value: 0 }],
        });
        expect(await parse(new IconSetExtXform(), '<x14:iconSet iconSet="3Hearts"/>')).to.include({ iconSet: "3TrafficLights1" });
        // an icon of an unknown set is no icon (it made the write fail)
        expect(await parse(new CfIconExtXform(), '<x14:cfIcon iconSet="3Hearts" iconId="2"/>')).to.deep.equal({
            iconSet: "NoIcons",
            iconId: 0,
        });
        expect(await parse(new CfIconExtXform(), '<x14:cfIcon iconSet="3Stars" iconId="2"/>')).to.deep.equal({
            iconSet: "3Stars",
            iconId: 2,
        });
    });

    it("<x14:dataBar axisPosition direction>", async () => {
        expect(await parse(new DatabarExtXform(), '<x14:dataBar axisPosition="left" direction="up"/>')).to.include({
            axisPosition: "auto",
            direction: "context",
        });
        expect(await parse(new DatabarExtXform(), '<x14:dataBar axisPosition="middle" direction="rightToLeft"/>')).to.include({
            axisPosition: "middle",
            direction: "rightToLeft",
        });
    });

    it("<c>: the error code of the file and the text of a t=d result are kept as text", async () => {
        // recent error codes are kept as written
        expect(await parse(new CellXform(), '<c r="A1" t="e"><v>#SPILL!</v></c>')).to.deep.include({ value: { error: "#SPILL!" } });
        // a malformed t="d" formula cell with runs: its text, as for the other types (it was the runs object)
        const cell = await parse(new CellXform(), '<c r="A1" t="d"><f>B1</f><is><r><t>x</t></r></is></c>');
        expect(cell.result).to.be.a("string");
    });
});
