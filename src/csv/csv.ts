import * as fs from "fs";
import { pipeline } from "stream";
import * as fastCsv from "fast-csv";
import customParseFormat from "dayjs/plugin/customParseFormat";
import utc from "dayjs/plugin/utc";
import dayjs from "dayjs";
import StreamBuf from "../utils/stream-buf";
import colCache from "../utils/col-cache";
import { exists } from "../utils/utils";
import type { CellErrorValue, CellValue, CsvReadOptions, CsvWriteOptions } from "../../index";
import type Workbook from "../doc/workbook";
import type { CellValue as DocCellValue } from "../doc/cell";
import type Worksheet from "../doc/worksheet";

dayjs.extend(customParseFormat);
dayjs.extend(utc);

/** A field of a CSV file mapped to a cell value (the default `map` of {@link CSV.read}). */
export type CsvCellValue = CellValue | Date | CellErrorValue;

// the CSV texts of booleans and errors (a Map: a field '__proto__' or 'constructor' is plain text)
const SpecialValues: ReadonlyMap<string, boolean | CellErrorValue> = new Map<string, boolean | CellErrorValue>([
    ["true", true],
    ["false", false],
    ["#N/A", { error: "#N/A" }],
    ["#REF!", { error: "#REF!" }],
    ["#NAME?", { error: "#NAME?" }],
    ["#DIV/0!", { error: "#DIV/0!" }],
    ["#NULL!", { error: "#NULL!" }],
    ["#VALUE!", { error: "#VALUE!" }],
    ["#NUM!", { error: "#NUM!" }],
]);

// plain integer literal (no decimal point or exponent)
const INTEGER = /^\s*[+-]?\d+\s*$/;

// The default date formats. A strict dayjs parse only accepts the text that the format would print, so a date in
// one of them is 10 ("YYYY-MM-DD", "MM-DD-YYYY") to 25 ("2020-01-05T10:20:30+01:00") characters long and starts
// with two digits. Other fields skip the four parses, which made reading text fields ~100 times slower.
const DEFAULT_DATE_FORMATS: readonly string[] = ["YYYY-MM-DD[T]HH:mm:ssZ", "YYYY-MM-DD[T]HH:mm:ss", "MM-DD-YYYY", "YYYY-MM-DD"];
const DEFAULT_DATE_SHAPE = /^\d\d[\s\S]{8,23}$/;

/** A header name of the parser (fast-csv): `undefined` skips the column. */
type HeaderName = string | null | undefined;

// [name, key] per header name, or undefined for a column to skip (an undefined name, as in fast-csv). Keys are unique:
// a repeated name (or empty name) is numbered "name_2", "name_3"... skipping names already in use.
function uniqueKeys(names: readonly HeaderName[]): ([name: string | null, key: string] | undefined)[] {
    const used = new Set(names);
    const next = new Map<string, number>();
    const taken = new Set<string>();
    return names.map((name) => {
        if (name === undefined) {
            return undefined;
        }
        let key = String(name);
        if (taken.has(key)) {
            let n = next.get(key) || 2;
            while (used.has(`${key}_${n}`) || taken.has(`${key}_${n}`)) {
                n++;
            }
            next.set(key, n + 1);
            key = `${key}_${n}`;
        }
        taken.add(key);
        return [name, key];
    });
}

/**
 * Reads a CSV file into a worksheet of the workbook, and writes a worksheet as CSV (with fast-csv). Node.js and Bun
 * only for the files; streams and buffers anywhere.
 */
class CSV {
    /** The workbook read into and written from. */
    readonly workbook: Workbook;
    /** Never set (always null): kept for compatibility with the JavaScript version. */
    worksheet: Worksheet | null;

    constructor(workbook: Workbook) {
        this.workbook = workbook;
        this.worksheet = null;
    }

    /**
     * Reads a CSV file into a new worksheet.
     *
     * @throws (rejects) when the file does not exist, or as {@link read}
     */
    async readFile(filename: string, options?: Partial<CsvReadOptions>): Promise<Worksheet> {
        options = options || {};
        if (!(await exists(filename))) {
            throw new Error(`File not found: ${filename}`);
        }
        const stream = fs.createReadStream(filename);
        const worksheet = await this.read(stream, options);
        stream.close();
        return worksheet;
    }

    /**
     * Reads CSV text into a new worksheet (`sheetName`). Each field is mapped to a cell value by `map`; by default an
     * empty field is null, a number a number (a long integer that would lose digits stays text), a date in one of
     * `dateFormats` a Date, `true` / `false` a boolean, `#N/A`... an error value, anything else text. With
     * `parserOptions.headers`, the header names become the first row and the column keys (made unique).
     *
     * @throws (rejects) on an error of the input or of the parser, or a header of more than 16,384 names
     */
    read(stream: NodeJS.ReadableStream, options?: Partial<CsvReadOptions>): Promise<Worksheet> {
        const readOptions = options || {};

        return new Promise((resolve, reject) => {
            const worksheet = this.workbook.addWorksheet(readOptions.sheetName);

            const dateFormats = readOptions.dateFormats || DEFAULT_DATE_FORMATS;
            const mayBeDate = (datum: string): boolean => dateFormats !== DEFAULT_DATE_FORMATS || DEFAULT_DATE_SHAPE.test(datum);
            // the value of the cell (a user map may return anything: the Cell setter stores objects as JSON values;
            // undefined leaves the cell empty)
            const map: (datum: string, index: number) => DocCellValue | undefined =
                readOptions.map ||
                function (datum: string): CsvCellValue {
                    if (datum === "") {
                        return null;
                    }
                    if (datum.trim() === "") {
                        // Number('  ') is 0
                        return datum;
                    }
                    const datumNumber = Number(datum);
                    if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
                        if (!Number.isSafeInteger(datumNumber) && INTEGER.test(datum)) {
                            // e.g. a 20-digit id: a Number would silently lose digits
                            return datum;
                        }
                        return datumNumber;
                    }
                    if (mayBeDate(datum)) {
                        for (const dateFormat of dateFormats) {
                            const dayjsObj = dayjs(datum, dateFormat, true);
                            if (dayjsObj.isValid()) {
                                return new Date(dayjsObj.valueOf());
                            }
                        }
                    }
                    const special = SpecialValues.get(datum);
                    if (special !== undefined) {
                        return special;
                    }
                    return datum;
                };

            // With parserOptions.headers, fast-csv turns each row into an object keyed by the header names, so values
            // under repeated (or empty, or "__proto__") names would be lost. Hand it the column indexes as names
            // instead (it still applies its column count checks) and keep the real names here.
            const parserOptions: fastCsv.ParserOptionsArgs = { ...readOptions.parserOptions };
            let headers: ReturnType<typeof uniqueKeys> | undefined; // per column: [name, key], or undefined for a column to skip
            const setHeaders = (names: readonly HeaderName[]): (string | undefined)[] => {
                if (names.length > colCache.MAX_COLS) {
                    throw new Error(`CSV header has ${names.length} names: a worksheet has at most ${colCache.MAX_COLS} columns`);
                }
                const columns = uniqueKeys(names);
                headers = columns;
                return columns.map((header, index) => header && String(index));
            };
            const requestedHeaders = parserOptions.headers;
            if (requestedHeaders === true) {
                parserOptions.headers = setHeaders;
            } else if (typeof requestedHeaders === "function") {
                parserOptions.headers = (row) => setHeaders(requestedHeaders(row));
            } else if (Array.isArray(requestedHeaders)) {
                parserOptions.headers = setHeaders(requestedHeaders);
            }

            const csvStream = fastCsv.parse(parserOptions);
            // an error thrown by a listener would escape fast-csv (e.g. from its 'headers' event) and crash the
            // process: fail the stream instead, which rejects the promise
            const guard =
                <T>(listener: (arg: T) => void) =>
                (arg: T): void => {
                    try {
                        listener(arg);
                    } catch (error) {
                        csvStream.destroy(error as Error);
                    }
                };
            csvStream
                .on(
                    "headers",
                    guard(() => {
                        // the header names become the first row, and the column keys (the parser emits 'headers' only
                        // with the headers option, which setHeaders handles)
                        // a null header name is kept as is (index.d.ts types Column.header as string)
                        worksheet.columns = headers!.map((header) => (header ? { header: header[0] as string, key: header[1] } : {}));
                    })
                )
                .on(
                    "data",
                    guard((data: Record<string, string> | string[]) => {
                        if (headers) {
                            // the row is keyed by the column indexes (see setHeaders); fast-csv fills a missing field with ""
                            const row = data as Record<string, string>;
                            worksheet.addRow(headers.map((header, index) => (header ? map(row[index]!, index) : undefined)));
                        } else {
                            worksheet.addRow((data as string[]).map(map));
                        }
                    })
                )
                .on("end", () => resolve(worksheet))
                .on("error", reject);

            // unlike pipe(), pipeline() forwards errors of the input stream
            pipeline(stream, csvStream, (error) => {
                if (error) {
                    reject(error);
                }
            });
        });
    }

    /**
     * @deprecated since version 4.0. You should use `CSV#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
     * @throws Error always
     */
    createInputStream(): never {
        throw new Error(
            "`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md"
        );
    }

    /**
     * Writes a worksheet (`sheetName` or `sheetId`, else the first one) as CSV. Each cell value is mapped to a field by
     * `map`; by default a hyperlink is its link, a formula its result, a date ISO 8601 (or `dateFormat`, local time or
     * `dateUTC`), an error its text, another object its JSON. Empty rows are written unless `includeEmptyRows` is false.
     *
     * @throws (rejects) on an error of the output stream or of the formatter
     */
    write(stream: NodeJS.WritableStream, options?: Partial<CsvWriteOptions>): Promise<void> {
        return new Promise((resolve, reject) => {
            const writeOptions = options || {};
            // const encoding = options.encoding || 'utf8';
            // const separator = options.separator || ',';
            // const quoteChar = options.quoteChar || '\'';

            const worksheet = this.workbook.getWorksheet(writeOptions.sheetName || writeOptions.sheetId);

            // index.d.ts FastCsvFormatterOptionsArgs.transform is typed apart from fast-csv's own
            const csvStream = fastCsv.format(
                writeOptions.formatterOptions as fastCsv.FormatterOptionsArgs<fastCsv.FormatterRow, fastCsv.FormatterRow>
            );
            stream.on("finish", () => {
                resolve();
            });
            // without a listener, an error of the destination (ENOENT, ENOSPC...) crashed the process
            stream.on("error", reject);
            csvStream.on("error", reject);
            csvStream.pipe(stream);

            const { dateFormat, dateUTC } = writeOptions;
            // index.d.ts types the values given to a user map as its CellValue
            const map: (value: DocCellValue, index: number) => unknown =
                (writeOptions.map as ((value: DocCellValue, index: number) => unknown) | undefined) ||
                ((value: DocCellValue): unknown => {
                    if (value && typeof value === "object") {
                        // the properties of the value objects of index.d.ts (hyperlink, formula, error)
                        const object = value as {
                            text?: unknown;
                            hyperlink?: unknown;
                            formula?: unknown;
                            result?: unknown;
                            error?: unknown;
                        };
                        if (object.text || object.hyperlink) {
                            return object.hyperlink || object.text || "";
                        }
                        if (object.formula || object.result) {
                            return object.result || "";
                        }
                        if (value instanceof Date) {
                            if (dateFormat) {
                                return dateUTC ? dayjs.utc(value).format(dateFormat) : dayjs(value).format(dateFormat);
                            }
                            return dateUTC ? dayjs.utc(value).format() : dayjs(value).format();
                        }
                        if (object.error) {
                            return object.error;
                        }
                        return JSON.stringify(value);
                    }
                    return value;
                });

            const includeEmptyRows = writeOptions.includeEmptyRows === undefined || writeOptions.includeEmptyRows;
            let lastRow = 1;
            if (worksheet) {
                worksheet.eachRow((row, rowNumber) => {
                    if (includeEmptyRows) {
                        while (lastRow++ < rowNumber - 1) {
                            csvStream.write([]);
                        }
                    }
                    const { values } = row;
                    values.shift();
                    csvStream.write(values.map(map));
                    lastRow = rowNumber;
                });
            }
            csvStream.end();
        });
    }

    /**
     * Writes a worksheet to a CSV file (`encoding`, utf8 by default), as {@link write}.
     *
     * @throws (rejects) as {@link write}
     */
    writeFile(filename: string, options?: Partial<CsvWriteOptions>): Promise<void> {
        options = options || {};

        const streamOptions = {
            encoding: (options.encoding || "utf8") as BufferEncoding,
        };
        const stream = fs.createWriteStream(filename, streamOptions);

        return this.write(stream, options);
    }

    /**
     * Writes a worksheet as CSV, as {@link write}.
     *
     * @returns the CSV bytes
     */
    async writeBuffer(options?: Partial<CsvWriteOptions>): Promise<Buffer<ArrayBuffer>> {
        const stream = new StreamBuf();
        await this.write(stream, options);
        return stream.read();
    }
}

export default CSV;
