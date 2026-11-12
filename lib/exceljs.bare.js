// Entry point of dist/exceljs.bare(.min).js
const ExcelJS = {
    Workbook: require("./doc/workbook"),
};

Object.assign(ExcelJS, require("./doc/enums"));

module.exports = ExcelJS;
