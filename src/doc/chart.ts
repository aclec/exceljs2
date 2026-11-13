// The charts created by `worksheet.addChart` and `workbook.addChartsheet` (and their streaming counterparts,
// src/stream/xlsx/stream-chart.ts): the options are validated and the data ranges resolved to worksheets when a chart is
// added (validateChart: a clear Error for invalid options), the cached values read from the cells when the workbook is
// written (chartSpaceSpec, from a reader of the cells that the document and the streaming writer each provide). A chart is
// written as the charts read from a file are (src/xlsx/chart-parts.ts): `toModel` gives the anchor and the chart part
// (src/xlsx/chart-space.ts) that the worksheet model carries.

import type { ChartGrouping, ChartLegendPosition, ChartOptions, ChartType, ImageEditAs } from "../../index";
import colCache from "../utils/col-cache";
import { dateToExcel, excelToDate } from "../utils/utils";
import { EMU_PER_PIXEL, MAX_COORDINATE, renderChartSpace, chartAnchorModel } from "../xlsx/chart-space";
import { decodeReference } from "./sheet-reference";
import { ValueType } from "./enums";
import Anchor, { type AnchorWorksheet } from "./anchor";
import type Cell from "./cell";
import type Worksheet from "./worksheet";
import type Workbook from "./workbook";
import type {
    AnchorModel,
    ChartAbsoluteAnchorSpec,
    ChartAnchorModel,
    ChartAnchorSpec,
    ChartAxisSpec,
    ChartDataRef,
    ChartNumberCache,
    ChartPoint,
    ChartSeriesSpec,
    ChartSpaceSpec,
    ChartStringCache,
} from "../types";

const CHART_TYPES: readonly ChartType[] = ["bar", "column", "line", "pie", "doughnut", "area", "scatter"];
const LEGEND_POSITIONS: readonly ChartLegendPosition[] = ["r", "l", "t", "b", "tr"];
const EDIT_AS: readonly ImageEditAs[] = ["twoCell", "oneCell", "absolute"];
/** The groupings of each type of chart that has one. */
const GROUPINGS: Readonly<Partial<Record<ChartType, readonly ChartGrouping[]>>> = {
    bar: ["clustered", "stacked", "percentStacked"],
    column: ["clustered", "stacked", "percentStacked"],
    line: ["standard", "stacked", "percentStacked"],
    area: ["standard", "stacked", "percentStacked"],
};
/** The most series of a chart (Excel's limit). */
const MAX_SERIES = 255;
// a color: 6 hexadecimal digits, optionally after "#" (anchored, one character class: linear)
const COLOR = /^#?[0-9A-Fa-f]{6}$/;
// a sheet name written without quotes in a formula: an identifier that is neither a cell (A1) nor a R1C1 reference
// (anchored, no nested quantifier: linear; a sheet name has at most 31 characters anyway)
const PLAIN_SHEET_NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/;
const CELL_LIKE = /^[A-Za-z]{1,3}[0-9]+$/;
const R1C1_LIKE = /^(?:[Rr][0-9]*)?(?:[Cc][0-9]*)?$/;
/** The number format of a date without one (the format ExcelJS gives a date cell, numFmtId 14). */
const DATE_FORMAT = "mm-dd-yy";

/** A worksheet whose cells a chart shows (of the document, or of the streaming writer): its name when it is written. */
export interface ChartDataSheet {
    readonly name: string;
}

/** A range of cells of a worksheet (1-based, inclusive). */
export interface ChartRange<S extends ChartDataSheet> {
    readonly sheet: S;
    readonly top: number;
    readonly left: number;
    readonly bottom: number;
    readonly right: number;
}

/** A series once validated: its ranges resolved to worksheets. */
export interface ChartSeriesDefinition<S extends ChartDataSheet> {
    readonly name?: string | ChartRange<S> | undefined;
    readonly categories?: ChartRange<S> | undefined;
    readonly values: ChartRange<S>;
    readonly color?: string | undefined;
}

/** A chart once validated: everything but the cached values of its ranges, read when the workbook is written. */
export interface ChartDefinition<S extends ChartDataSheet> {
    readonly type: ChartType;
    readonly grouping?: ChartGrouping | undefined;
    readonly marker: boolean;
    readonly line: boolean;
    readonly title?: string | undefined;
    readonly legend?: ChartLegendPosition | undefined;
    readonly series: readonly ChartSeriesDefinition<S>[];
    readonly categoryAxis: ChartAxisSpec;
    readonly valueAxis: ChartAxisSpec;
    /** Where the chart is in its worksheet; absent for the chart of a chartsheet (it fills the sheet). */
    readonly anchor?: Omit<ChartAnchorSpec, "name"> | undefined;
}

/** Where the options of a chart are validated ({@link validateChart}): the worksheets of its workbook. */
export interface ChartContext<S extends ChartDataSheet> {
    /**
     * The worksheet that holds the chart: a range without a sheet name is on it, and the chart needs an `anchor`.
     * `undefined` for the chart of a chartsheet: every range names its worksheet, and there is no anchor.
     */
    readonly sheet: S | undefined;
    /** The sizes of the cells of that worksheet: they place a corner given as a fractional `{ col, row }`. */
    readonly cellSizes?: AnchorWorksheet | undefined;
    /** @returns the worksheet of a name (case-insensitive), if any */
    findSheet(name: string): S | undefined;
}

/** The kinds of the values of a {@link ChartRangeValues}. */
const NUMBER = 0;
const DATE = 1;
const TEXT = 2;

/** The capacity of the arrays of a {@link ChartRangeValues} when it gets its first value. */
const INITIAL_CAPACITY = 16;

/** Two digits. */
function pad(n: number): string {
    return String(n).padStart(2, "0");
}

/**
 * The text of a date in a cache of texts (categories mixing texts and dates, a series name): `yyyy-mm-dd`, with
 * ` hh:mm:ss` when it has a time (UTC, as ExcelJS reads and writes dates). ExcelJS has no number formatter: Excel shows
 * the date in the format of its cell once it recalculates the chart.
 */
function dateText(serial: number, date1904: boolean): string {
    const date = excelToDate(serial, date1904);
    const day = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
    const hours = date.getUTCHours();
    const minutes = date.getUTCMinutes();
    const seconds = date.getUTCSeconds();
    return hours || minutes || seconds ? `${day} ${pad(hours)}:${pad(minutes)}:${pad(seconds)}` : day;
}

/**
 * The values of the cells of a range for a chart: its number of cells, and the cells with a value by increasing index
 * (0-based, row by row). Stored compactly, as the streaming writer holds them until the workbook is committed: typed
 * arrays of the indexes, of the numbers (a date as its serial number) and of their kinds, the texts in a map of the
 * positions that have one, the number format of the first number: 13 bytes per number (up to twice that while the arrays
 * grow), no object per cell.
 */
export class ChartRangeValues {
    /** The number of cells of the range (`<c:ptCount>`). */
    readonly count: number;
    /** The number format of the first number (`mm-dd-yy` for a date without one). */
    formatCode: string | undefined;
    private readonly date1904: boolean;
    private indexes = new Int32Array(0);
    private numbers = new Float64Array(0);
    private kinds = new Uint8Array(0);
    private readonly texts = new Map<number, string>();
    private size = 0;

    /**
     * @param count - the number of cells of the range
     * @param date1904 - the date system of the workbook
     */
    constructor(count: number, date1904: boolean) {
        this.count = count;
        this.date1904 = date1904;
    }

    /** Whether every value is a number or a date, and there is one: categories cached as numbers. */
    get numeric(): boolean {
        return this.size > 0 && this.texts.size === 0;
    }

    /**
     * Adds the value of a cell (the result of a formula): nothing for an empty cell, nor for a cell merged into another
     * one (empty, as Excel reads it).
     * @param index - its 0-based index in the range, greater than those added before
     */
    addCell(index: number, cell: Cell | undefined): void {
        if (!cell || cell.type === ValueType.Merge) {
            return;
        }
        const value: unknown = cell.type === ValueType.Formula ? cell.result : cell.value;
        if (value === null || value === undefined || value === "") {
            return;
        }
        if (typeof value === "number") {
            if (Number.isFinite(value)) {
                this.push(index, NUMBER, value, cell.numFmt);
            } else {
                this.pushText(index, String(value));
            }
        } else if (value instanceof Date) {
            if (Number.isFinite(value.getTime())) {
                this.push(index, DATE, dateToExcel(value, this.date1904), cell.numFmt ?? DATE_FORMAT);
            }
        } else if (typeof value === "string") {
            this.pushText(index, value);
        } else if (typeof value === "boolean") {
            this.pushText(index, value ? "TRUE" : "FALSE");
        } else if (isObject(value) && typeof member(value, "error") === "string") {
            this.pushText(index, String(member(value, "error")));
        } else {
            // rich text, hyperlink: its text
            const { text } = cell;
            if (text) {
                this.pushText(index, text);
            }
        }
    }

    /** The numbers (and dates), for a cache of numbers. */
    *numberPoints(): Generator<ChartPoint<number>> {
        for (let i = 0; i < this.size; i++) {
            if (this.kinds[i] !== TEXT) {
                yield { index: this.indexes[i] ?? 0, value: this.numbers[i] ?? 0 };
            }
        }
    }

    /** The texts of the values (a number as its text, a date as {@link dateText}), for a cache of texts. */
    *textPoints(): Generator<ChartPoint<string>> {
        for (let i = 0; i < this.size; i++) {
            const index = this.indexes[i] ?? 0;
            const number = this.numbers[i] ?? 0;
            const kind = this.kinds[i];
            if (kind === NUMBER) {
                yield { index, value: String(number) };
            } else if (kind === DATE) {
                yield { index, value: dateText(number, this.date1904) };
            } else {
                yield { index, value: this.texts.get(i) ?? "" };
            }
        }
    }

    /** Drops the values. */
    clear(): void {
        this.indexes = new Int32Array(0);
        this.numbers = new Float64Array(0);
        this.kinds = new Uint8Array(0);
        this.texts.clear();
        this.size = 0;
    }

    private pushText(index: number, text: string): void {
        this.texts.set(this.size, text);
        this.push(index, TEXT, 0, undefined);
    }

    private push(index: number, kind: number, number: number, formatCode: string | undefined): void {
        if (this.size === this.indexes.length) {
            // doubled: amortized constant time, at most twice the memory of the values
            const capacity = Math.max(INITIAL_CAPACITY, this.size * 2);
            const indexes = new Int32Array(capacity);
            indexes.set(this.indexes);
            const numbers = new Float64Array(capacity);
            numbers.set(this.numbers);
            const kinds = new Uint8Array(capacity);
            kinds.set(this.kinds);
            this.indexes = indexes;
            this.numbers = numbers;
            this.kinds = kinds;
        }
        this.indexes[this.size] = index;
        this.numbers[this.size] = number;
        this.kinds[this.size] = kind;
        this.size++;
        if (kind !== TEXT) {
            this.formatCode ??= formatCode;
        }
    }
}

/**
 * Reads the values of the cells of a range: `undefined` when they are not known (the streaming writer: a row of the range
 * was committed before the chart was added).
 */
export type ChartCellReader<S extends ChartDataSheet> = (range: ChartRange<S>) => ChartRangeValues | undefined;

/** Where the chart of a chartsheet is in its drawing: the size Excel gives it (the sheet zooms it to fit the window). */
const CHARTSHEET_ANCHOR: ChartAbsoluteAnchorSpec["absolute"] = { x: 0, y: 0, cx: 9304587, cy: 6077107 };

/** Whether a value is a non-null object (not an array). */
function isObject(value: unknown): value is object {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** A member of an object given as options (own or inherited, as a destructuring reads it). */
function member(value: object, key: string): unknown {
    return Reflect.get(value, key);
}

/** A value as an error message shows it: short, never throws. */
function show(value: unknown): string {
    if (typeof value === "string") {
        return JSON.stringify(value.length > 64 ? `${value.slice(0, 64)}...` : value);
    }
    if (typeof value === "number" || typeof value === "boolean" || value === undefined || value === null) {
        return String(value);
    }
    return Array.isArray(value) ? "an array" : `a ${typeof value}`;
}

/** The sheet name of a formula: quoted (`'My Sheet'`, a quote doubled) unless it is a plain identifier. */
export function quoteSheetName(name: string): string {
    if (PLAIN_SHEET_NAME.test(name) && !CELL_LIKE.test(name) && !R1C1_LIKE.test(name)) {
        return name;
    }
    return `'${name.replaceAll("'", "''")}'`;
}

/** The formula of a range: `Sheet1!$B$2:$B$5`, `'My Sheet'!$A$1` for a single cell. */
function rangeFormula(range: ChartRange<ChartDataSheet>): string {
    const corner = (row: number, col: number): string => `$${colCache.n2l(col)}$${row}`;
    const tl = corner(range.top, range.left);
    const cells = range.top === range.bottom && range.left === range.right ? tl : `${tl}:${corner(range.bottom, range.right)}`;
    return `${quoteSheetName(range.sheet.name)}!${cells}`;
}

/** The values of the cells of a range of a document worksheet, now. */
function readRange(range: ChartRange<Worksheet>, date1904: boolean): ChartRangeValues {
    const { sheet } = range;
    const width = range.right - range.left + 1;
    const values = new ChartRangeValues((range.bottom - range.top + 1) * width, date1904);
    for (let row = range.top; row <= range.bottom; row++) {
        const found = sheet.findRow(row);
        if (found) {
            for (let col = range.left; col <= range.right; col++) {
                values.addCell((row - range.top) * width + col - range.left, found.findCell(col));
            }
        }
    }
    return values;
}

/** The cached numbers of a range (the cells without a number have no point). */
function numberCache(values: ChartRangeValues): ChartNumberCache {
    return {
        type: "number",
        formatCode: values.formatCode || "General",
        count: values.count,
        points: iterable(() => values.numberPoints()),
    };
}

/** The cached texts of a range (the empty cells have no point). */
function stringCache(values: ChartRangeValues): ChartStringCache {
    return { type: "string", count: values.count, points: iterable(() => values.textPoints()) };
}

/** An iterable that can be iterated more than once (a generator function called again each time). */
function iterable<T>(items: () => Iterator<T>): Iterable<T> {
    return { [Symbol.iterator]: items };
}

/**
 * The chart part of a chart, its cached values read now.
 * @param read - the values of the cells of a range
 * @param date1904 - the date system of the workbook
 */
export function chartSpaceSpec<S extends ChartDataSheet>(
    definition: ChartDefinition<S>,
    read: ChartCellReader<S>,
    date1904: boolean
): ChartSpaceSpec {
    const series = definition.series.map((item): ChartSeriesSpec => {
        // the values of the ranges of the series; a series with a range whose values are not known has no cache at all,
        // so that a viewer that does not recalculate does not show a part of it (labels without values...)
        const values = new Map<ChartRange<S>, ChartRangeValues>();
        for (const range of seriesRanges(item)) {
            const data = read(range);
            if (!data) {
                values.clear();
                break;
            }
            values.set(range, data);
        }
        let name: ChartSeriesSpec["name"];
        if (typeof item.name === "object") {
            const data = values.get(item.name);
            name = { formula: rangeFormula(item.name), cache: data && stringCache(data) };
        } else {
            ({ name } = item);
        }
        let categories: ChartDataRef<ChartNumberCache | ChartStringCache> | undefined;
        if (item.categories) {
            const data = values.get(item.categories);
            // numbers (or dates) as numbers, anything else as texts
            const numeric = !!data && data.numeric;
            categories = {
                formula: rangeFormula(item.categories),
                cache: data && (numeric ? numberCache(data) : stringCache(data)),
            };
        }
        const data = values.get(item.values);
        return {
            name,
            categories,
            values: { formula: rangeFormula(item.values), cache: data && numberCache(data) },
            color: item.color,
        };
    });
    return {
        type: definition.type,
        grouping: definition.grouping,
        marker: definition.marker,
        line: definition.line,
        title: definition.title,
        legend: definition.legend,
        date1904,
        series,
        categoryAxis: definition.categoryAxis,
        valueAxis: definition.valueAxis,
    };
}

/**
 * The anchor and the chart part of a chart, its cached values read now (see {@link chartSpaceSpec}).
 * @param name - the name of its graphic frame (`Chart 1`)
 */
export function chartModel<S extends ChartDataSheet>(
    definition: ChartDefinition<S>,
    name: string,
    read: ChartCellReader<S>,
    date1904: boolean
): ChartAnchorModel {
    const xml = renderChartSpace(chartSpaceSpec(definition, read, date1904));
    const anchor: ChartAnchorSpec | ChartAbsoluteAnchorSpec = definition.anchor
        ? { ...definition.anchor, name }
        : { absolute: CHARTSHEET_ANCHOR, name };
    return chartAnchorModel(anchor, xml);
}

/** The ranges of a series: its name (a cell), categories and values. */
export function* seriesRanges<S extends ChartDataSheet>(item: ChartSeriesDefinition<S>): Generator<ChartRange<S>> {
    if (typeof item.name === "object") {
        yield item.name;
    }
    if (item.categories) {
        yield item.categories;
    }
    yield item.values;
}

/**
 * The context of a chart of the document: the worksheets of the workbook.
 * @param worksheet - the worksheet of the chart, undefined for the chart of a chartsheet
 */
function documentContext(workbook: Workbook, worksheet: Worksheet | undefined): ChartContext<Worksheet> {
    return {
        sheet: worksheet,
        cellSizes: worksheet,
        findSheet(name) {
            const lower = name.toLowerCase();
            return workbook.worksheets.find((sheet) => sheet.name.toLowerCase() === lower);
        },
    };
}

/**
 * The reader of the cells of the document worksheets, now.
 * @param removed - the message of the Error thrown for a worksheet that was removed from the workbook
 */
function documentReader(workbook: Workbook, removed: (sheet: Worksheet) => string): ChartCellReader<Worksheet> {
    const date1904 = !!workbook.properties.date1904;
    return (range) => {
        const { sheet } = range;
        if (sheet.workbook.getWorksheet(sheet.id) !== sheet) {
            throw new Error(removed(sheet));
        }
        return readRange(range, date1904);
    };
}

/** A chart of a worksheet, created by `worksheet.addChart` (see index.d.ts `ChartOptions`). */
class Chart {
    /** The worksheet that holds the chart. */
    readonly worksheet: Worksheet;
    /** The type of the chart. */
    readonly type: ChartType;
    private readonly definition: ChartDefinition<Worksheet>;

    /**
     * @throws Error for invalid options: the message names the option (a series range that is not one row or one column, a
     *   worksheet that does not exist, a grouping of another type of chart...)
     */
    constructor(worksheet: Worksheet, options: ChartOptions) {
        this.worksheet = worksheet;
        this.definition = validateChart(options, documentContext(worksheet.workbook, worksheet));
        this.type = this.definition.type;
    }

    /** The title of the chart. */
    get title(): string | undefined {
        return this.definition.title;
    }

    /** Removes the chart from its worksheet. */
    remove(): void {
        const charts = this.worksheet._addedCharts;
        const index = charts.indexOf(this);
        if (index !== -1) {
            charts.splice(index, 1);
        }
    }

    /**
     * The chart, its data read from the cells now: the type, options and cached values of its chart part.
     * @throws Error when a worksheet of its data was removed from the workbook
     */
    get spec(): ChartSpaceSpec {
        const { workbook } = this.worksheet;
        return chartSpaceSpec(this.definition, this.reader(), !!workbook.properties.date1904);
    }

    /**
     * The chart as a chart read from a file is carried (its anchor, and its chart part generated now: see {@link spec}).
     * @param name - the name of its graphic frame (`Chart 1`)
     * @throws Error when a worksheet of its data was removed from the workbook
     */
    toModel(name: string): ChartAnchorModel {
        const { workbook } = this.worksheet;
        return chartModel(this.definition, name, this.reader(), !!workbook.properties.date1904);
    }

    private reader(): ChartCellReader<Worksheet> {
        return documentReader(
            this.worksheet.workbook,
            (sheet) =>
                `A ${this.type} chart of the worksheet ${show(this.worksheet.name)} uses the data of the worksheet ${show(sheet.name)}, which was removed: remove the chart (chart.remove())`
        );
    }
}

/** The chart of a chartsheet created by `workbook.addChartsheet` (see index.d.ts `ChartsheetChartOptions`). */
export class ChartsheetChart {
    private readonly workbook: Workbook;
    private readonly definition: ChartDefinition<Worksheet>;

    /**
     * @throws Error for invalid options (as {@link Chart}; every range names its worksheet, and there is no anchor)
     */
    constructor(workbook: Workbook, options: unknown) {
        this.workbook = workbook;
        this.definition = validateChart(options, documentContext(workbook, undefined));
    }

    /**
     * The chart of the drawing of the chartsheet (an absolute anchor), its chart part generated now.
     * @param chartsheet - the name of the chartsheet, for the message
     * @throws Error when a worksheet of its data was removed from the workbook
     */
    toModel(chartsheet: string): ChartAnchorModel {
        const read = documentReader(
            this.workbook,
            (sheet) =>
                `The chartsheet ${show(chartsheet)} uses the data of the worksheet ${show(sheet.name)}, which was removed: remove the chartsheet (workbook.removeChartsheet(${show(chartsheet)}))`
        );
        return chartModel(this.definition, "Chart 1", read, !!this.workbook.properties.date1904);
    }
}

/**
 * Validates the options of a chart, and resolves its ranges to the worksheets of the context.
 * @throws Error for invalid options: the message names the option
 */
export function validateChart<S extends ChartDataSheet>(options: unknown, context: ChartContext<S>): ChartDefinition<S> {
    return new ChartValidator(context).chart(options);
}

/** Validates the options of a chart (see {@link validateChart}). */
class ChartValidator<S extends ChartDataSheet> {
    private readonly context: ChartContext<S>;

    constructor(context: ChartContext<S>) {
        this.context = context;
    }

    /** @throws Error for invalid options */
    chart(options: unknown): ChartDefinition<S> {
        if (!isObject(options)) {
            throw new Error(`Invalid chart options: ${show(options)}, expected an object`);
        }
        const type = member(options, "type");
        const chartType = CHART_TYPES.find((name) => name === type);
        if (!chartType) {
            throw new Error(`Invalid chart type: ${show(type)}, expected one of ${CHART_TYPES.join(", ")}`);
        }
        const pie = chartType === "pie" || chartType === "doughnut";
        const scatter = chartType === "scatter";

        const grouping = member(options, "grouping");
        const groupings = GROUPINGS[chartType];
        let chartGrouping: ChartGrouping | undefined;
        if (grouping !== undefined) {
            chartGrouping = groupings?.find((name) => name === grouping);
            if (!chartGrouping) {
                throw new Error(
                    groupings
                        ? `Invalid grouping for a ${chartType} chart: ${show(grouping)}, expected one of ${groupings.join(", ")}`
                        : `A ${chartType} chart has no grouping`
                );
            }
        }

        const marker = this.flag(options, "marker", chartType === "line" || scatter, chartType, true);
        const line = this.flag(options, "line", scatter, chartType, false);
        if (scatter && !marker && !line) {
            throw new Error("A scatter chart needs markers or lines (marker: false and line: false show nothing)");
        }

        const title = this.text(member(options, "title"), "title");

        const legendOption = member(options, "legend");
        let legend: ChartLegendPosition | undefined = "r";
        if (legendOption === false) {
            legend = undefined;
        } else if (legendOption !== undefined) {
            if (!isObject(legendOption)) {
                throw new Error(`Invalid chart legend: ${show(legendOption)}, expected { position } or false`);
            }
            const position = member(legendOption, "position");
            if (position !== undefined) {
                legend = LEGEND_POSITIONS.find((name) => name === position);
                if (!legend) {
                    throw new Error(`Invalid chart legend position: ${show(position)}, expected one of ${LEGEND_POSITIONS.join(", ")}`);
                }
            }
        }

        const seriesOption = member(options, "series");
        if (!Array.isArray(seriesOption) || !seriesOption.length || seriesOption.length > MAX_SERIES) {
            throw new Error(`Invalid chart series: ${show(seriesOption)}, expected an array of 1 to ${MAX_SERIES} series`);
        }
        const series = seriesOption.map((item: unknown, index) => this.series(item, index + 1, scatter));

        const axesOption = member(options, "axes");
        let axes: object = {};
        if (axesOption !== undefined) {
            if (pie) {
                throw new Error(`A ${chartType} chart has no axes`);
            }
            if (!isObject(axesOption)) {
                throw new Error(`Invalid chart axes: ${show(axesOption)}, expected an object`);
            }
            axes = axesOption;
        }
        const [categoryKey, valueKey] = scatter ? ["x", "y"] : ["category", "value"];
        for (const key of Object.keys(axes)) {
            if (key !== categoryKey && key !== valueKey) {
                throw new Error(`Invalid chart axis ${show(key)}: a ${chartType} chart has the axes ${categoryKey} and ${valueKey}`);
            }
        }
        const categoryAxis = this.axis(member(axes, categoryKey), categoryKey, scatter, false);
        const valueAxis = this.axis(member(axes, valueKey), valueKey, true, true);

        const anchorOption = member(options, "anchor");
        let anchor: ChartDefinition<S>["anchor"];
        if (this.context.sheet) {
            anchor = this.anchor(anchorOption);
        } else if (anchorOption !== undefined) {
            throw new Error("The chart of a chartsheet has no anchor: it fills the sheet");
        }
        return { type: chartType, grouping: chartGrouping, marker, line, title, legend, series, categoryAxis, valueAxis, anchor };
    }

    /** A boolean option of some types of chart (`marker`, `line`). */
    private flag(options: object, key: string, applies: boolean, type: ChartType, defaultValue: boolean): boolean {
        const value = member(options, key);
        if (value === undefined) {
            return defaultValue;
        }
        if (!applies) {
            throw new Error(`The option ${key} does not apply to a ${type} chart`);
        }
        if (typeof value !== "boolean") {
            throw new Error(`Invalid chart option ${key}: ${show(value)}, expected a boolean`);
        }
        return value;
    }

    /** An optional text (a title). */
    private text(value: unknown, what: string): string | undefined {
        if (value !== undefined && typeof value !== "string") {
            throw new Error(`Invalid chart ${what}: ${show(value)}, expected a string`);
        }
        return value;
    }

    /** An axis: `values` for a value axis (min, max allowed), `defaultGridlines` as Excel draws them. */
    private axis(value: unknown, key: string, values: boolean, defaultGridlines: boolean): ChartAxisSpec {
        if (value === undefined) {
            return { majorGridlines: defaultGridlines, visible: true };
        }
        if (!isObject(value)) {
            throw new Error(`Invalid chart axis ${key}: ${show(value)}, expected an object`);
        }
        const axis = value;
        const what = `axis ${key}`;
        const bound = (name: "min" | "max"): number | undefined => {
            const bound = member(axis, name);
            if (bound === undefined) {
                return undefined;
            }
            if (!values) {
                throw new Error(`The chart ${what} is a category axis: it has no ${name}`);
            }
            if (typeof bound !== "number" || !Number.isFinite(bound)) {
                throw new Error(`Invalid chart ${what} ${name}: ${show(bound)}, expected a finite number`);
            }
            return bound;
        };
        const min = bound("min");
        const max = bound("max");
        if (min !== undefined && max !== undefined && min >= max) {
            throw new Error(`Invalid chart ${what}: min (${min}) must be less than max (${max})`);
        }
        const numFmt = this.text(member(axis, "numFmt"), `${what} numFmt`);
        if (numFmt === "") {
            throw new Error(`Invalid chart ${what} numFmt: an empty string`);
        }
        const boolean = (name: "majorGridlines" | "visible", defaultValue: boolean): boolean => {
            const flag = member(axis, name);
            if (flag === undefined) {
                return defaultValue;
            }
            if (typeof flag !== "boolean") {
                throw new Error(`Invalid chart ${what} ${name}: ${show(flag)}, expected a boolean`);
            }
            return flag;
        };
        return {
            title: this.text(member(axis, "title"), `${what} title`),
            min,
            max,
            numFmt,
            majorGridlines: boolean("majorGridlines", defaultGridlines),
            visible: boolean("visible", true),
        };
    }

    /** A series (`number`: 1-based, for the messages). */
    private series(value: unknown, number: number, scatter: boolean): ChartSeriesDefinition<S> {
        const what = `series ${number}`;
        if (!isObject(value)) {
            throw new Error(`Invalid chart ${what}: ${show(value)}, expected an object`);
        }
        const [categoriesKey, valuesKey] = scatter ? ["xValues", "yValues"] : ["categories", "values"];
        const [otherCategories, otherValues] = scatter ? ["categories", "values"] : ["xValues", "yValues"];
        for (const key of [otherCategories, otherValues]) {
            if (member(value, key) !== undefined) {
                throw new Error(
                    `Invalid chart ${what}: ${key} is for ${scatter ? "other" : "scatter"} charts, a ${scatter ? "scatter" : "category"} chart takes ${categoriesKey} and ${valuesKey}`
                );
            }
        }
        const valuesOption = member(value, valuesKey);
        if (valuesOption === undefined) {
            throw new Error(`Invalid chart ${what}: ${valuesKey} is required (a range such as "B2:B10")`);
        }
        const values = this.range(valuesOption, `${what} ${valuesKey}`, true);
        const categoriesOption = member(value, categoriesKey);
        const categories = categoriesOption === undefined ? undefined : this.range(categoriesOption, `${what} ${categoriesKey}`, true);

        const nameOption = member(value, "name");
        let name: string | ChartRange<S> | undefined;
        if (typeof nameOption === "string" || nameOption === undefined) {
            name = nameOption;
        } else if (isObject(nameOption) && typeof member(nameOption, "ref") === "string") {
            name = this.range(member(nameOption, "ref"), `${what} name`, false);
            if (name.top !== name.bottom || name.left !== name.right) {
                throw new Error(`Invalid chart ${what} name: the reference must be one cell`);
            }
        } else {
            throw new Error(`Invalid chart ${what} name: ${show(nameOption)}, expected a string or { ref: "B1" }`);
        }

        const colorOption = member(value, "color");
        let color: string | undefined;
        if (colorOption !== undefined) {
            if (typeof colorOption !== "string" || !COLOR.test(colorOption)) {
                throw new Error(`Invalid chart ${what} color: ${show(colorOption)}, expected 6 hexadecimal digits RRGGBB`);
            }
            color = colorOption.replace("#", "").toUpperCase();
        }
        return { name, categories, values, color };
    }

    /**
     * A range of cells (`"B2:B10"` of the worksheet of the chart, `"Sheet1!$B$2:$B$10"`, `"'My Sheet'!B2:B10"`).
     * @param line - the range must be one row or one column
     */
    private range(value: unknown, what: string, line: boolean): ChartRange<S> {
        const reference = typeof value === "string" ? decodeReference(value) : undefined;
        if (!reference) {
            throw new Error(`Invalid chart ${what}: ${show(value)}, expected a range such as "B2:B10" or "'My Sheet'!$B$2:$B$10"`);
        }
        let { sheet } = this.context;
        if (reference.sheetName !== undefined) {
            sheet = this.context.findSheet(reference.sheetName);
            if (!sheet) {
                throw new Error(`Invalid chart ${what}: there is no worksheet ${show(reference.sheetName)}`);
            }
        } else if (!sheet) {
            throw new Error(
                `Invalid chart ${what}: ${show(value)} names no worksheet (the ranges of the chart of a chartsheet are such as "Data!B2:B10")`
            );
        }
        if (line && reference.top !== reference.bottom && reference.left !== reference.right) {
            throw new Error(`Invalid chart ${what}: ${show(value)} is not one row or one column`);
        }
        return { sheet, top: reference.top, left: reference.left, bottom: reference.bottom, right: reference.right };
    }

    /** Where the chart is: a range of cells (`"E2:L20"`), two corners, or a corner and a size. */
    private anchor(value: unknown): NonNullable<ChartDefinition<S>["anchor"]> {
        if (typeof value === "string") {
            const reference = decodeReference(value);
            if (!reference || reference.sheetName !== undefined) {
                throw new Error(`Invalid chart anchor: ${show(value)}, expected a range of cells such as "E2:L20"`);
            }
            // the chart covers the cells: from the top-left corner of the first one to the bottom-right corner of the last
            return {
                from: { nativeCol: reference.left - 1, nativeColOff: 0, nativeRow: reference.top - 1, nativeRowOff: 0 },
                to: { nativeCol: reference.right, nativeColOff: 0, nativeRow: reference.bottom, nativeRowOff: 0 },
            };
        }
        if (!isObject(value)) {
            throw new Error(`Invalid chart anchor: ${show(value)}, expected a range ("E2:L20"), { tl, br } or { tl, ext }`);
        }
        const from = this.position(member(value, "tl"), "tl");
        const br = member(value, "br");
        const ext = member(value, "ext");
        const editAs = member(value, "editAs");
        if ((br === undefined) === (ext === undefined)) {
            throw new Error("Invalid chart anchor: expected either br (the bottom-right corner) or ext (the size in pixels)");
        }
        if (br !== undefined) {
            const to = this.position(br, "br");
            if (to.nativeCol < from.nativeCol || (to.nativeCol === from.nativeCol && to.nativeColOff < from.nativeColOff)) {
                throw new Error("Invalid chart anchor: br is left of tl");
            }
            if (to.nativeRow < from.nativeRow || (to.nativeRow === from.nativeRow && to.nativeRowOff < from.nativeRowOff)) {
                throw new Error("Invalid chart anchor: br is above tl");
            }
            let editAsValue: ImageEditAs | undefined;
            if (editAs !== undefined) {
                editAsValue = EDIT_AS.find((name) => name === editAs);
                if (!editAsValue) {
                    throw new Error(`Invalid chart anchor editAs: ${show(editAs)}, expected one of ${EDIT_AS.join(", ")}`);
                }
            }
            return { from, to, editAs: editAsValue };
        }
        if (editAs !== undefined) {
            throw new Error("Invalid chart anchor: editAs applies to an anchor with two corners (tl and br)");
        }
        if (!isObject(ext)) {
            throw new Error(`Invalid chart anchor ext: ${show(ext)}, expected { width, height } in pixels`);
        }
        const size = (name: "width" | "height"): number => {
            const length = member(ext, name);
            if (typeof length !== "number" || !Number.isFinite(length) || length <= 0) {
                throw new Error(`Invalid chart anchor ext ${name}: ${show(length)}, expected a positive number of pixels`);
            }
            // written in EMU: at most the largest coordinate of DrawingML
            if (Math.floor(length * EMU_PER_PIXEL) > MAX_COORDINATE) {
                throw new Error(
                    `Invalid chart anchor ext ${name}: ${show(length)}, expected a positive number of pixels up to ${Math.floor(MAX_COORDINATE / EMU_PER_PIXEL)}`
                );
            }
            return length;
        };
        return { from, ext: { width: size("width"), height: size("height") } };
    }

    /** A corner: `{ col, row }` (fractional, 0-based) or `{ nativeCol, nativeRow, nativeColOff?, nativeRowOff? }`. */
    private position(value: unknown, what: string): AnchorModel {
        if (!isObject(value)) {
            throw new Error(`Invalid chart anchor ${what}: ${show(value)}, expected { col, row } or { nativeCol, nativeRow }`);
        }
        const index = (name: string, max: number, integer: boolean): number => {
            const n = member(value, name);
            if (typeof n !== "number" || !Number.isFinite(n) || n < 0 || n >= max || (integer && !Number.isInteger(n))) {
                throw new Error(
                    `Invalid chart anchor ${what}.${name}: ${show(n)}, expected ${integer ? "an integer" : "a number"} from 0 to less than ${max}`
                );
            }
            return n;
        };
        const offset = (name: string): number => {
            const n = member(value, name);
            if (n === undefined) {
                return 0;
            }
            if (typeof n !== "number" || !Number.isInteger(n) || n < 0 || n > MAX_COORDINATE) {
                throw new Error(
                    `Invalid chart anchor ${what}.${name}: ${show(n)}, expected an offset in EMU (an integer from 0 to ${MAX_COORDINATE})`
                );
            }
            return n;
        };
        if (member(value, "nativeCol") !== undefined) {
            return {
                nativeCol: index("nativeCol", colCache.MAX_COLS, true),
                nativeColOff: offset("nativeColOff"),
                nativeRow: index("nativeRow", colCache.MAX_ROWS, true),
                nativeRowOff: offset("nativeRowOff"),
            };
        }
        const col = index("col", colCache.MAX_COLS, false);
        const row = index("row", colCache.MAX_ROWS, false);
        // the offset in the cell, from the size of the column and the row (as an image's corner: doc/anchor.ts)
        const corner = new Anchor(this.context.cellSizes, { col, row }).model;
        const largest = Math.max(corner.nativeColOff, corner.nativeRowOff);
        if (largest > MAX_COORDINATE) {
            throw new Error(`Invalid chart anchor ${what}: the offset in its cell (${largest} EMU) is beyond ${MAX_COORDINATE} EMU`);
        }
        return corner;
    }
}

export default Chart;
