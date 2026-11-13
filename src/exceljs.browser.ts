/**
 * Entry point of `dist/exceljs(.min).js`. Since v5 the bundle targets ES2020 and no longer ships core-js/regenerator
 * polyfills, so it is identical to the bare build.
 */
import ExcelJS from "./exceljs.bare";

export = ExcelJS;
