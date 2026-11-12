const fs = require("fs");
const { pipeline } = require("stream");
const fastCsv = require("fast-csv");
const customParseFormat = require("dayjs/plugin/customParseFormat");
const utc = require("dayjs/plugin/utc");
const dayjs = require("dayjs").extend(customParseFormat).extend(utc);
const StreamBuf = require("../utils/stream-buf");
const colCache = require("../utils/col-cache");

const {
    fs: { exists },
} = require("../utils/utils");

const SpecialValues = {
    true: true,
    false: false,
    "#N/A": { error: "#N/A" },
    "#REF!": { error: "#REF!" },
    "#NAME?": { error: "#NAME?" },
    "#DIV/0!": { error: "#DIV/0!" },
    "#NULL!": { error: "#NULL!" },
    "#VALUE!": { error: "#VALUE!" },
    "#NUM!": { error: "#NUM!" },
};
/* eslint-ensable quote-props */

// plain integer literal (no decimal point or exponent)
const INTEGER = /^\s*[+-]?\d+\s*$/;

// [name, key] per header name, or undefined for a column to skip (an undefined or null name). Keys are unique:
// a repeated name (or empty name) is numbered "name_2", "name_3"... skipping names already in use.
function uniqueKeys(names) {
    const used = new Set(names);
    const next = new Map();
    const taken = new Set();
    return names.map((name) => {
        if (name === undefined || name === null) {
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

class CSV {
    constructor(workbook) {
        this.workbook = workbook;
        this.worksheet = null;
    }

    async readFile(filename, options) {
        options = options || {};
        if (!(await exists(filename))) {
            throw new Error(`File not found: ${filename}`);
        }
        const stream = fs.createReadStream(filename);
        const worksheet = await this.read(stream, options);
        stream.close();
        return worksheet;
    }

    read(stream, options) {
        options = options || {};

        return new Promise((resolve, reject) => {
            const worksheet = this.workbook.addWorksheet(options.sheetName);

            const dateFormats = options.dateFormats || ["YYYY-MM-DD[T]HH:mm:ssZ", "YYYY-MM-DD[T]HH:mm:ss", "MM-DD-YYYY", "YYYY-MM-DD"];
            const map =
                options.map ||
                function (datum) {
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
                    const dt = dateFormats.reduce((matchingDate, currentDateFormat) => {
                        if (matchingDate) {
                            return matchingDate;
                        }
                        const dayjsObj = dayjs(datum, currentDateFormat, true);
                        if (dayjsObj.isValid()) {
                            return dayjsObj;
                        }
                        return null;
                    }, null);
                    if (dt) {
                        return new Date(dt.valueOf());
                    }
                    // own properties only: '__proto__' or 'constructor' are plain text, not Object internals
                    if (Object.prototype.hasOwnProperty.call(SpecialValues, datum)) {
                        return SpecialValues[datum];
                    }
                    return datum;
                };

            // With parserOptions.headers, fast-csv turns each row into an object keyed by the header names, so values
            // under repeated (or empty, or "__proto__") names would be lost. Hand it the column indexes as names
            // instead (it still applies its column count checks) and keep the real names here.
            const parserOptions = { ...options.parserOptions };
            let headers; // per column: [name, key], or undefined for a column to skip
            const setHeaders = (names) => {
                if (names.length > colCache.MAX_COLS) {
                    throw new Error(`CSV header has ${names.length} names: a worksheet has at most ${colCache.MAX_COLS} columns`);
                }
                headers = uniqueKeys(names);
                return headers.map((header, index) => header && String(index));
            };
            if (parserOptions.headers === true) {
                parserOptions.headers = setHeaders;
            } else if (typeof parserOptions.headers === "function") {
                const transform = parserOptions.headers;
                parserOptions.headers = (row) => setHeaders(transform(row));
            } else if (Array.isArray(parserOptions.headers)) {
                parserOptions.headers = setHeaders(parserOptions.headers);
            }

            const csvStream = fastCsv.parse(parserOptions);
            // an error thrown by a listener would escape fast-csv (e.g. from its 'headers' event) and crash the
            // process: fail the stream instead, which rejects the promise
            const guard = (listener) => (arg) => {
                try {
                    listener(arg);
                } catch (error) {
                    csvStream.destroy(error);
                }
            };
            csvStream
                .on(
                    "headers",
                    guard(() => {
                        // the header names become the first row, and the column keys
                        worksheet.columns = headers.map((header) => (header ? { header: header[0], key: header[1] } : {}));
                    })
                )
                .on(
                    "data",
                    guard((data) => {
                        if (headers) {
                            worksheet.addRow(headers.map((header, index) => (header ? map(data[index], index) : undefined)));
                        } else {
                            worksheet.addRow(data.map(map));
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
     */
    createInputStream() {
        throw new Error(
            "`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md"
        );
    }

    write(stream, options) {
        return new Promise((resolve, reject) => {
            options = options || {};
            // const encoding = options.encoding || 'utf8';
            // const separator = options.separator || ',';
            // const quoteChar = options.quoteChar || '\'';

            const worksheet = this.workbook.getWorksheet(options.sheetName || options.sheetId);

            const csvStream = fastCsv.format(options.formatterOptions);
            stream.on("finish", () => {
                resolve();
            });
            csvStream.on("error", reject);
            csvStream.pipe(stream);

            const { dateFormat, dateUTC } = options;
            const map =
                options.map ||
                ((value) => {
                    if (value) {
                        if (value.text || value.hyperlink) {
                            return value.hyperlink || value.text || "";
                        }
                        if (value.formula || value.result) {
                            return value.result || "";
                        }
                        if (value instanceof Date) {
                            if (dateFormat) {
                                return dateUTC ? dayjs.utc(value).format(dateFormat) : dayjs(value).format(dateFormat);
                            }
                            return dateUTC ? dayjs.utc(value).format() : dayjs(value).format();
                        }
                        if (value.error) {
                            return value.error;
                        }
                        if (typeof value === "object") {
                            return JSON.stringify(value);
                        }
                    }
                    return value;
                });

            const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
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

    writeFile(filename, options) {
        options = options || {};

        const streamOptions = {
            encoding: options.encoding || "utf8",
        };
        const stream = fs.createWriteStream(filename, streamOptions);

        return this.write(stream, options);
    }

    async writeBuffer(options) {
        const stream = new StreamBuf();
        await this.write(stream, options);
        return stream.read();
    }
}

module.exports = CSV;
