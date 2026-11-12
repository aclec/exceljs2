// Compile-only checks (bun run test:types) of the index.d.ts types of styles, conditional formatting, tables and data
// validations, aligned with the runtime. Nothing here is executed. Lines marked @ts-expect-error must stay errors.

import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

declare const worksheet: ExcelJS.Worksheet;
declare const cell: ExcelJS.Cell;

// --- colors, fonts, borders --------------------------------------------------------------------
export function colorsFontsBorders(): void {
    cell.font = { color: { theme: 4, tint: -0.25 }, condense: true, extend: false, shadow: true };
    cell.font = { color: { indexed: 64 } };
    cell.border = { top: { style: "thin" }, bottom: { style: "double" }, color: { argb: "FFFF0000" } };
    worksheet.properties.tabColor = { theme: 5, tint: 0.4 };
    cell.font = { size: 10.5, color: { auto: true } };
    const tint: number | undefined = cell.font?.color?.tint;
    const shadow: boolean | undefined = cell.font?.shadow;
    void tint;
    void shadow;
}

// --- fills ---------------------------------------------------------------------------------------
export function fills(fill: ExcelJS.Fill): void {
    // a gradient read without degree: neither `gradient` nor `degree`
    cell.fill = { type: "gradient", stops: [{ position: 0, color: { argb: "FF0000FF" } }, { position: 1 }] };
    cell.fill = { type: "gradient", gradient: "path", center: { left: 0.2, top: 0.2, right: 0.8, bottom: 0.8 }, stops: [] };
    // @ts-expect-error a path gradient has a center
    cell.fill = { type: "gradient", gradient: "path", stops: [] };
    // @ts-expect-error a cell fill has a pattern
    cell.fill = { type: "pattern", fgColor: { argb: "FF000000" } };
    if (fill.type === "gradient") {
        if (fill.gradient === "path") {
            const right: number | undefined = fill.center.right;
            void right;
        } else {
            const degree: true = {} as Equals<typeof fill.degree, number | undefined>;
            void degree;
        }
        const stopColor: true = {} as Equals<ExcelJS.GradientStop["color"], Partial<ExcelJS.Color> | undefined>;
        void stopColor;
    }
}

// --- differential formats: conditional formatting and table columns --------------------------------
export function differentialStyles(table: ExcelJS.Table): void {
    const style = table.getColumn(0).style;
    const font: true = {} as Equals<NonNullable<typeof style>["font"], Partial<ExcelJS.Font> | null | undefined>;
    void font;
    // the pattern fill of a dxf may have no pattern; `{}` is a fill without pattern nor gradient
    const dxf: ExcelJS.DifferentialStyle = { fill: { type: "pattern", bgColor: { argb: "FFFFC7CE" } }, numFmt: null, border: null };
    const empty: ExcelJS.DifferentialStyle = { fill: {} };
    const tableStyle: ExcelJS.TableColumnStyle = dxf;
    // a cell style is a differential format
    const fromStyle: ExcelJS.DifferentialStyle = cell.style;
    void empty;
    void tableStyle;
    void fromStyle;
}

// --- conditional formatting rules ------------------------------------------------------------------
export function conditionalFormatting(): void {
    worksheet.addConditionalFormatting({
        ref: "A1:E7",
        rules: [
            { type: "cellIs", operator: "between", formulae: [1, "$B$1"], style: { font: { bold: true } } },
            { type: "expression", formulae: ["MOD(ROW(),2)=0"], style: { fill: { type: "pattern", bgColor: { argb: "FFFFC7CE" } } } },
            { type: "top10" },
            { type: "top10", rank: 5, percent: true, bottom: true },
            { type: "aboveAverage" },
            {
                type: "colorScale",
                cfvo: [{ type: "min" }, { type: "max" }],
                color: [{ argb: "FFF8696B" }, { theme: 6, tint: 0.4 }],
            },
            {
                type: "iconSet",
                custom: true,
                cfvo: [
                    { type: "percent", value: 0 },
                    { type: "percent", value: 33 },
                    { type: "percent", value: 67 },
                ],
                icons: [{ iconSet: "3Flags", iconId: 0 }, { iconSet: "NoIcons" }, { iconSet: "3Stars", iconId: 2 }],
            },
            { type: "dataBar", color: { argb: "FF638EC6" } },
        ],
    });

    // @ts-expect-error a color scale needs its thresholds (writing it throws)
    const noCfvo: ExcelJS.ConditionalFormattingRule = { type: "colorScale", color: [] };
    // @ts-expect-error a color scale needs its colors (writing it throws)
    const noColor: ExcelJS.ConditionalFormattingRule = { type: "colorScale", cfvo: [] };
    // @ts-expect-error an icon set needs its thresholds (writing it throws)
    const noIconCfvo: ExcelJS.ConditionalFormattingRule = { type: "iconSet", iconSet: "3Arrows" };
    const legacyIconSet: ExcelJS.ConditionalFormattingRule = {
        type: "iconSet",
        // @ts-expect-error not an icon set of the file format (the default is 3TrafficLights1)
        iconSet: "3TrafficLights",
        cfvo: [],
    };
    void legacyIconSet;
    // @ts-expect-error a formula is a text, a number or a boolean
    const dateFormula: ExcelJS.ConditionalFormattingRule = { type: "cellIs", formulae: [new Date()] };
    void noCfvo;
    void noColor;
    void noIconCfvo;
    void dateFormula;

    const rank: true = {} as Equals<ExcelJS.Top10RuleType["rank"], number | undefined>;
    const scaleColor: true = {} as Equals<ExcelJS.ColorScaleRuleType["color"], (Partial<ExcelJS.Color> | undefined)[]>;
    const formulae: true = {} as Equals<ExcelJS.CellIsRuleType["formulae"], ExcelJS.CfFormula[] | undefined>;
    void rank;
    void scaleColor;
    void formulae;
}

// --- tables and data validations ---------------------------------------------------------------------
export function tablesAndValidations(): void {
    worksheet.addTable({
        name: "T",
        ref: "A1",
        columns: [{ name: "Date" }, { name: "Amount", style: { numFmt: "0.00" } }],
        rows: [
            [new Date(), 1.5],
            [null, { formula: "B2*2" }],
        ],
    });
    worksheet.addTable({
        name: "U",
        ref: "A1",
        columns: [{ name: "A" }],
        // @ts-expect-error a table row holds cell values
        rows: [[Symbol("not a value")]],
    });
    cell.dataValidation = { type: "whole", operator: "between", formulae: [1, 10], errorStyle: "warning", showErrorMessage: true };
    cell.dataValidation = { type: "date", operator: "greaterThan", formulae: ["TODAY()"] };
}
