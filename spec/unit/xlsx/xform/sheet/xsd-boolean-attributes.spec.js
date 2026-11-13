// The boolean attributes of the worksheet elements are xsd:boolean: "true" / "false" are as valid as "1" / "0" (LibreOffice
// and other producers write them). They were read with `=== "1"` (or as numbers), so that "true" was false.
const parseSax = verquire("utils/parse-sax");
const { cloneDeep } = require("../../../../utils/under-dash");

const PageSetupPropertiesXform = verquire("xlsx/xform/sheet/page-setup-properties-xform");
const SheetProtectionXform = verquire("xlsx/xform/sheet/sheet-protection-xform");
const PrintOptionsXform = verquire("xlsx/xform/sheet/print-options-xform");
const SheetViewXform = verquire("xlsx/xform/sheet/sheet-view-xform");
const PageSetupXform = verquire("xlsx/xform/sheet/page-setup-xform");
const HeaderFooterXform = verquire("xlsx/xform/sheet/header-footer-xform");
const OutlinePropertiesXform = verquire("xlsx/xform/sheet/outline-properties-xform");
const CfRuleXform = verquire("xlsx/xform/sheet/cf/cf-rule-xform");
const DatabarXform = verquire("xlsx/xform/sheet/cf/databar-xform");
const IconSetXform = verquire("xlsx/xform/sheet/cf/icon-set-xform");
const CfRuleExtXform = verquire("xlsx/xform/sheet/cf-ext/cf-rule-ext-xform");
const DatabarExtXform = verquire("xlsx/xform/sheet/cf-ext/databar-ext-xform");
const IconSetExtXform = verquire("xlsx/xform/sheet/cf-ext/icon-set-ext-xform");

async function parse(Xform, xml) {
    // without the undefined keys
    return cloneDeep(await new Xform().parse(parseSax([xml])), false);
}

describe("xsd:boolean attributes of the worksheet elements", () => {
    it("<pageSetUpPr fitToPage>", async () => {
        expect(await parse(PageSetupPropertiesXform, '<pageSetUpPr fitToPage="true"/>')).to.deep.equal({ fitToPage: true });
        expect(await parse(PageSetupPropertiesXform, '<pageSetUpPr fitToPage="false"/>')).to.deep.equal({ fitToPage: false });
    });

    it("<sheetProtection>", async () => {
        const spelled = await parse(
            SheetProtectionXform,
            '<sheetProtection sheet="true" objects="true" scenarios="true" selectLockedCells="true" selectUnlockedCells="true" ' +
                'formatCells="false" formatColumns="false" formatRows="false" insertColumns="false" insertRows="false" ' +
                'insertHyperlinks="false" deleteColumns="false" deleteRows="false" sort="false" autoFilter="false" pivotTables="false"/>'
        );
        const digits = await parse(
            SheetProtectionXform,
            '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1" selectUnlockedCells="1" ' +
                'formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" ' +
                'insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"/>'
        );
        expect(spelled).to.deep.equal(digits);
        expect(spelled).to.deep.include({ sheet: true, objects: false, selectLockedCells: false, formatCells: true, pivotTables: true });
        // the defaults spelled out are the defaults
        expect(await parse(SheetProtectionXform, '<sheetProtection sheet="false" objects="false" formatCells="true"/>')).to.deep.equal({});
    });

    it("<printOptions>", async () => {
        expect(
            await parse(
                PrintOptionsXform,
                '<printOptions headings="true" gridLines="true" horizontalCentered="true" verticalCentered="true"/>'
            )
        ).to.deep.equal({ showRowColHeaders: true, showGridLines: true, horizontalCentered: true, verticalCentered: true });
    });

    it("<sheetView>", async () => {
        const model = await parse(
            SheetViewXform,
            '<sheetView rightToLeft="true" showRuler="false" showRowColHeaders="false" showGridLines="false" workbookViewId="0"/>'
        );
        expect(model).to.deep.include({ rightToLeft: true, showRuler: false, showRowColHeaders: false, showGridLines: false });
        const defaults = await parse(
            SheetViewXform,
            '<sheetView rightToLeft="false" showRuler="true" showRowColHeaders="true" showGridLines="true" workbookViewId="0"/>'
        );
        expect(defaults).to.deep.include({ rightToLeft: false, showRuler: true, showRowColHeaders: true, showGridLines: true });
    });

    it("<pageSetup>", async () => {
        const model = await parse(PageSetupXform, '<pageSetup blackAndWhite="true" draft="true" usePrinterDefaults="true"/>');
        expect(model).to.deep.include({ blackAndWhite: true, draft: true, usePrinterDefaults: true });
    });

    it("<headerFooter>", async () => {
        expect(await parse(HeaderFooterXform, '<headerFooter differentFirst="true" differentOddEven="true"/>')).to.deep.equal({
            differentFirst: true,
            differentOddEven: true,
        });
    });

    it("<outlinePr>", async () => {
        expect(await parse(OutlinePropertiesXform, '<outlinePr summaryBelow="true" summaryRight="false"/>')).to.deep.equal({
            summaryBelow: true,
            summaryRight: false,
        });
    });

    it("<cfRule>", async () => {
        expect(
            await parse(
                CfRuleXform,
                '<cfRule type="top10" dxfId="0" priority="1" stopIfTrue="true" percent="true" bottom="true" rank="5"/>'
            )
        ).to.deep.equal({ type: "top10", dxfId: 0, priority: 1, stopIfTrue: true, percent: true, bottom: true, rank: 5 });
        expect(await parse(CfRuleXform, '<cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="true"/>')).to.deep.equal({
            type: "aboveAverage",
            dxfId: 0,
            priority: 1,
            aboveAverage: true,
        });
    });

    it("<dataBar> and <iconSet>", async () => {
        expect(await parse(DatabarXform, '<dataBar showValue="true"><cfvo type="min"/><cfvo type="max"/></dataBar>')).to.deep.equal({
            cfvo: [{ type: "min" }, { type: "max" }],
            showValue: true,
        });
        expect(
            await parse(IconSetXform, '<iconSet reverse="true" showValue="true"><cfvo type="percent" val="0"/></iconSet>')
        ).to.deep.equal({
            iconSet: "3TrafficLights1",
            reverse: true,
            showValue: true,
            cfvo: [{ type: "percent", value: 0 }],
        });
    });

    it("<x14:cfRule>, <x14:dataBar> and <x14:iconSet>", async () => {
        expect(await parse(CfRuleExtXform, '<x14:cfRule type="expression" priority="1" stopIfTrue="true" id="{A}"/>')).to.deep.include({
            type: "expression",
            stopIfTrue: true,
        });
        const dataBar = await parse(
            DatabarExtXform,
            '<x14:dataBar border="true" gradient="false" negativeBarColorSameAsPositive="true" negativeBarBorderColorSameAsPositive="false"/>'
        );
        expect(dataBar).to.deep.include({
            border: true,
            gradient: false,
            negativeBarColorSameAsPositive: true,
            negativeBarBorderColorSameAsPositive: false,
        });
        expect(
            await parse(IconSetExtXform, '<x14:iconSet iconSet="3Arrows" reverse="true" showValue="false" custom="true"/>')
        ).to.deep.include({
            reverse: true,
            showValue: false,
            custom: true,
        });
    });
});
