/**
 * Entry point for Node.js and Bun (`require("@aclec/exceljs2")`).
 *
 * The public API and its types are documented in `index.d.ts`.
 */
import Workbook from "./doc/workbook";
import ModelContainer from "./doc/modelcontainer";
import WorkbookWriter from "./stream/xlsx/workbook-writer";
import WorkbookReader from "./stream/xlsx/workbook-reader";
import enums from "./doc/enums";
import headerFooter from "./doc/header-footer";

const ExcelJS = {
    Workbook,
    ModelContainer,
    headerFooter,
    stream: {
        xlsx: {
            WorkbookWriter,
            WorkbookReader,
        },
    },
    ...enums,
};

// `export =` keeps `module.exports = ExcelJS`, the CommonJS shape of the package entry point
export = ExcelJS;
