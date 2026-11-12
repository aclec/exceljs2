// Compile-only checks (bun run test:types) of the index.d.ts declarations of cells, rows and columns aligned with the
// runtime. Nothing here is executed. Lines marked @ts-expect-error must stay errors.

import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type Assert<T extends true> = T;

declare const cell: ExcelJS.Cell;
declare const row: ExcelJS.Row;
declare const column: ExcelJS.Column;

// --- cell values -----------------------------------------------------------------------------------------------------
export function cellValues(): void {
    // a hyperlink shows any value (a loaded one shows the value of its cell)
    cell.value = { text: 42, hyperlink: "https://example.com" };
    cell.value = { text: { richText: [{ text: "x", font: { bold: true } }] }, hyperlink: "#Sheet1!A1", tooltip: "t" };
    type HyperlinkText = Assert<
        Equals<ExcelJS.CellHyperlinkValue["text"], string | number | boolean | Date | ExcelJS.CellErrorValue | ExcelJS.CellRichTextValue>
    >;

    // a formula master: shared or array, with its range
    cell.value = { formula: "A1*2", result: 2, shareType: "shared", ref: "B1:B3" };
    cell.value = { sharedFormula: "B1", result: 4 };
    // @ts-expect-error date1904 is not read from a formula value
    cell.value = { formula: "A1*2", date1904: true };
    // @ts-expect-error a shareType is "shared" or "array"
    cell.value = { formula: "A1*2", shareType: "normal" };
    type SharedResult = Assert<Equals<ExcelJS.CellSharedFormulaValue["result"], ExcelJS.CellFormulaResult | undefined>>;

    // an error code of the file (recent versions of Excel), as read; the codes of ErrorValue are offered
    cell.value = { error: "#SPILL!" };
    cell.value = { error: ExcelJS.ErrorValue.NotApplicable };

    // undefined empties the cell
    cell.value = undefined;
    // @ts-expect-error an object that is not a cell value is not typed (the runtime holds it as a JSON value)
    cell.value = { foo: 1 };

    type Html = Assert<Equals<typeof cell.html, string>>;
    // @ts-expect-error html is read-only
    cell.html = "x";
    cell.names = ["a", "b"] as const;
    type Names = Assert<Equals<typeof cell.names, string[]>>;
}

// --- notes -----------------------------------------------------------------------------------------------------------
export function notes(): void {
    const note = cell.note;
    if (note && typeof note !== "string") {
        type Anchor = Assert<Equals<typeof note.anchor, string | undefined>>;
        // @ts-expect-error the anchor of a note read from a file is informational
        note.anchor = "0, 0, 0, 0, 1, 0, 1, 0";
    }
    type Inset = Assert<Equals<ExcelJS.CommentMargins["inset"], number[]>>;
    cell.note = { texts: [{ text: "n" }], margins: { insetmode: "custom", inset: [0.1, 0.1, 0.2, 0.2] }, editAs: "oneCells" };
}

// --- rows ------------------------------------------------------------------------------------------------------------
export function rows(): void {
    type Style = Assert<Equals<typeof row.style, Partial<ExcelJS.Style>>>;
    type OutlineLevel = Assert<Equals<typeof row.outlineLevel, number>>;
    row.outlineLevel = undefined;
    row.hidden = undefined;
    type Hidden = Assert<Equals<typeof row.hidden, boolean>>;

    type Model = Assert<Equals<typeof row.model, ExcelJS.RowModel | null>>;
    // @ts-expect-error setting a null model throws
    row.model = null;

    row.values = [1, "a", undefined, { formula: "A1" }] as const;
    row.values = { name: "x", age: 3 };
    row.values = null;
    type Values = Assert<Equals<typeof row.values, ExcelJS.CellValue[]>>;
    const values: ExcelJS.RowValues = ["a", 1] as const;
    row.values = values;

    const c1: ExcelJS.Cell = row.getCellEx({ address: "B1", col: 2 });
    const c2: ExcelJS.Cell | undefined = row.findCell(2);
    row.eachCell(null, (c, colNumber) => c.address + colNumber);
    row.eachCell({}, (c) => c.value);
    row.eachCell({ includeEmpty: true }, (c) => c.value);

    row.splice(2, 1, "a", 1, undefined, null, { richText: [{ text: "r" }] });
    // @ts-expect-error spliced values are cell values
    row.splice(2, 1, () => 1);
}

// --- columns ---------------------------------------------------------------------------------------------------------
export function columns(): void {
    type NumFmt = Assert<Equals<typeof column.numFmt, string | undefined>>;
    type Font = Assert<Equals<typeof column.font, Partial<ExcelJS.Font> | undefined>>;
    type OutlineLevel = Assert<Equals<typeof column.outlineLevel, number>>;
    column.outlineLevel = undefined;

    type Values = Assert<Equals<typeof column.values, ExcelJS.CellValue[]>>;
    column.values = [undefined, "header", 1, 2] as const;
    column.values = null;

    column.eachCell(undefined, (c, rowNumber) => c.address + rowNumber);
    column.eachCell({ includeEmpty: false }, (c) => c.value);
    column.defn = null;
}
