// Compile-only checks (bun run test:types) of the index.d.ts declarations corrected when the conformance check
// (src/types/conformance-helpers.ts) was made stricter. Nothing here is executed. Lines marked @ts-expect-error must stay
// errors.

import type { Stream } from "stream";
import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type Assert<T extends true> = T;

declare const workbook: ExcelJS.Workbook;
declare const worksheet: ExcelJS.Worksheet;
declare const container: ExcelJS.ModelContainer;
declare const legacy: Stream;

// --- worksheet / workbook models: what the setters read --------------------------------------------------------------
export function modelSetters(): void {
    // the getter gives every field
    const model = worksheet.model;
    type Dimensions = Assert<Equals<typeof model.dimensions, ExcelJS.Range>>;
    type Merges = Assert<Equals<typeof model.merges, string[]>>;
    // a model is set back as is
    worksheet.model = model;
    // the setter does not read the getter-only fields: they may be absent
    const { id, state, rowBreaks, dimensions, merges, ...rest } = model;
    worksheet.model = rest;
    worksheet.model = { ...rest, dimensions: "A1:B2" };
    type Input = Assert<Equals<ExcelJS.WorksheetModelInput["state"], ExcelJS.WorksheetState | undefined>>;
    // @ts-expect-error the setter reads the rows
    worksheet.model = { ...rest, rows: undefined };

    // a workbook model is set back as is, or from a container that read a file
    workbook.model = workbook.model;
    if (container.model) {
        workbook.model = container.model;
    }
    // @ts-expect-error the worksheet models of a workbook model need their id
    workbook.model = { ...workbook.model, worksheets: [rest] };
    type Ids = Assert<Equals<ExcelJS.WorkbookModelInput["worksheets"][number]["id"], number>>;
}

// --- a model read by a ModelContainer --------------------------------------------------------------------------------
export function containerModel(): void {
    type Model = Assert<Equals<typeof container.model, ExcelJS.WorkbookModelInput | undefined>>;
    const sheet = container.model?.worksheets[0];
    // a file read into a container gives no rowBreaks nor merges, and the dimensions as text
    type RowBreaks = Assert<Equals<NonNullable<typeof sheet>["rowBreaks"], ExcelJS.RowBreak[] | undefined>>;
    type Dimensions = Assert<Equals<NonNullable<typeof sheet>["dimensions"], ExcelJS.Range | string | undefined>>;
}

// --- the columns of a table in a worksheet model ---------------------------------------------------------------------
export function tableColumnModel(): void {
    const column = worksheet.model.tables[0]?.columns[0];
    // a table read from a file has the differential format of its columns: null for the parts the file does not set
    type Style = Assert<Equals<NonNullable<typeof column>["style"], ExcelJS.TableColumnStyle | undefined>>;
    const font = column?.style?.font;
    type Font = Assert<Equals<typeof font, Partial<ExcelJS.Font> | null | undefined>>;
    // @ts-expect-error null when the file does not set it
    const declared: Partial<ExcelJS.Font> | undefined = font;
    // the columns given to addTable still fit
    const properties: ExcelJS.TableColumnProperties = { name: "A", style: { font: { bold: true } } };
    const columnModel: ExcelJS.WorksheetTableColumnModel = properties;
}

// --- the xlsx output and input ---------------------------------------------------------------------------------------
export async function xlsxStreams(): Promise<void> {
    // the chunks given to an output are Node Buffers (the buffer of any ArrayBufferLike)
    type Chunk = Assert<Equals<Parameters<ExcelJS.XlsxOutput["write"]>[0], globalThis.Buffer>>;
    const chunks: Uint8Array[] = [];
    await workbook.xlsx.write({
        write(chunk: Uint8Array) {
            chunks.push(chunk);
        },
        end() {},
    });
    // an old-style stream (anything with pipe) is read
    await workbook.xlsx.read(legacy);
    // writeBuffer gives a Buffer over an ArrayBuffer
    type Written = Assert<Equals<Awaited<ReturnType<typeof workbook.xlsx.writeBuffer>>, ExcelJS.Buffer>>;
}
