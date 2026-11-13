/**
 * Entry point of `dist/exceljs.bare(.min).js`: the in-memory workbook API, without the Node-only streaming classes.
 */
import Workbook from "./doc/workbook";
import enums from "./doc/enums";

const ExcelJS = {
    Workbook,
    ...enums,
};

export = ExcelJS;
