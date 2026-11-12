const _ = require("../../../utils/under-dash");

const XmlStream = require("../../../utils/xml-stream");

const BaseXform = require("../base-xform");
const StaticXform = require("../static-xform");
const ListXform = require("../list-xform");
const DefinedNameXform = require("./defined-name-xform");
const SheetXform = require("./sheet-xform");
const WorkbookViewXform = require("./workbook-view-xform");
const WorkbookPropertiesXform = require("./workbook-properties-xform");
const WorkbookCalcPropertiesXform = require("./workbook-calc-properties-xform");
const WorkbookPivotCacheXform = require("./workbook-pivot-cache-xform");

// Worksheets are stored in an array indexed by sheetId. The file format allows any unsigned int, but Excel
// numbers sheets from 1 as they are added, so a huge id is hostile: sheetId="4294967294" made an array whose
// every traversal (workbook.worksheets, getWorksheet(name), addWorksheet, write) took minutes.
const MAX_SHEET_ID = 65535;

// Keeps valid ids; gives the lowest free ids to sheets whose id is not an integer in 1..MAX_SHEET_ID or is
// already used by an earlier sheet (which used to replace that sheet).
function renumberSheets(sheets) {
    const used = new Set();
    const invalid = sheets.filter((sheet) => {
        const valid = Number.isInteger(sheet.id) && sheet.id >= 1 && sheet.id <= MAX_SHEET_ID && !used.has(sheet.id);
        if (valid) {
            used.add(sheet.id);
        }
        return !valid;
    });
    let id = 1;
    invalid.forEach((sheet) => {
        while (used.has(id)) {
            id++;
        }
        sheet.id = id;
        used.add(id);
    });
}

class WorkbookXform extends BaseXform {
    constructor() {
        super();

        this.map = {
            fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
            workbookPr: new WorkbookPropertiesXform(),
            bookViews: new ListXform({
                tag: "bookViews",
                count: false,
                childXform: new WorkbookViewXform(),
            }),
            sheets: new ListXform({ tag: "sheets", count: false, childXform: new SheetXform() }),
            definedNames: new ListXform({
                tag: "definedNames",
                count: false,
                childXform: new DefinedNameXform(),
            }),
            calcPr: new WorkbookCalcPropertiesXform(),
            pivotCaches: new ListXform({
                tag: "pivotCaches",
                count: false,
                childXform: new WorkbookPivotCacheXform(),
            }),
        };
    }

    prepare(model) {
        model.sheets = model.worksheets;

        // collate all the print areas from all of the sheets and add them to the defined names
        const printAreas = [];
        let index = 0; // sheets is sparse array - calc index manually
        model.sheets.forEach((sheet) => {
            // quoted sheet name for a formula reference: embedded quotes are doubled
            const sheetRef = `'${sheet.name.replace(/'/g, "''")}'`;
            if (sheet.pageSetup && sheet.pageSetup.printArea) {
                // a sheet has at most one _xlnm.Print_Area: several areas are comma-separated ranges
                const ranges = sheet.pageSetup.printArea.split("&&").map((printArea) => {
                    const printAreaComponents = printArea.split(":").map((component) => `$${component}`);
                    return `${sheetRef}!${printAreaComponents.join(":")}`;
                });
                printAreas.push({
                    name: "_xlnm.Print_Area",
                    ranges,
                    localSheetId: index,
                });
            }

            if (sheet.pageSetup && (sheet.pageSetup.printTitlesRow || sheet.pageSetup.printTitlesColumn)) {
                const ranges = [];

                if (sheet.pageSetup.printTitlesColumn) {
                    const titlesColumns = sheet.pageSetup.printTitlesColumn.split(":");
                    ranges.push(`${sheetRef}!$${titlesColumns[0]}:$${titlesColumns[1]}`);
                }

                if (sheet.pageSetup.printTitlesRow) {
                    const titlesRows = sheet.pageSetup.printTitlesRow.split(":");
                    ranges.push(`${sheetRef}!$${titlesRows[0]}:$${titlesRows[1]}`);
                }

                const definedName = {
                    name: "_xlnm.Print_Titles",
                    ranges,
                    localSheetId: index,
                };

                printAreas.push(definedName);
            }
            index++;
        });
        if (printAreas.length) {
            model.definedNames = model.definedNames.concat(printAreas);
        }

        (model.media || []).forEach((medium, i) => {
            // assign name
            medium.name = medium.type + (i + 1);
        });
    }

    render(xmlStream, model) {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("workbook", WorkbookXform.WORKBOOK_ATTRIBUTES);

        this.map.fileVersion.render(xmlStream);
        this.map.workbookPr.render(xmlStream, model.properties);
        this.map.bookViews.render(xmlStream, model.views);
        this.map.sheets.render(xmlStream, model.sheets);
        this.map.definedNames.render(xmlStream, model.definedNames);
        this.map.calcPr.render(xmlStream, model.calcProperties);
        this.map.pivotCaches.render(xmlStream, model.pivotTables);

        xmlStream.closeNode();
    }

    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "workbook":
                return true;
            default:
                this.parser = this.map[node.name];
                if (this.parser) {
                    this.parser.parseOpen(node);
                }
                return true;
        }
    }

    parseText(text) {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    parseClose(name) {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "workbook":
                this.model = {
                    sheets: this.map.sheets.model,
                    properties: this.map.workbookPr.model || {},
                    views: this.map.bookViews.model,
                    calcProperties: {},
                };
                if (this.map.definedNames.model) {
                    this.model.definedNames = this.map.definedNames.model;
                }

                return false;
            default:
                // not quite sure how we get here!
                return true;
        }
    }

    reconcile(model) {
        // keyed by ids from the file: no prototype, so that an id such as "constructor" is unknown
        const rels = (model.workbookRels || []).reduce((map, rel) => {
            map[rel.Id] = rel;
            return map;
        }, Object.create(null));

        renumberSheets(model.sheets || []);

        // reconcile sheet ids, rIds and names
        const worksheets = [];
        let worksheet;
        let index = 0;

        (model.sheets || []).forEach((sheet) => {
            const rel = rels[sheet.rId];
            if (!rel) {
                return;
            }
            // if rel.Target start with `[space]/xl/` or `/xl/` , then it will be replaced with `''` and spliced behind `xl/`,
            // otherwise it will be spliced directly behind `xl/`. i.g.
            worksheet = model.worksheetHash[`xl/${rel.Target.replace(/^(\s|\/xl\/)+/, "")}`];
            // If there are "chartsheets" in the file, rel.Target will
            // come out as chartsheets/sheet1.xml or similar here, and
            // that won't be in model.worksheetHash.
            // As we don't have the infrastructure to support chartsheets,
            // we will ignore them for now:
            if (worksheet) {
                worksheet.name = sheet.name;
                worksheet.id = sheet.id;
                worksheet.state = sheet.state;
                worksheets[index++] = worksheet;
            }
        });

        // reconcile print areas
        const definedNames = [];
        _.each(model.definedNames, (definedName) => {
            if (definedName.name === "_xlnm.Print_Area") {
                worksheet = worksheets[definedName.localSheetId];
                if (worksheet) {
                    if (!worksheet.pageSetup) {
                        worksheet.pageSetup = {};
                    }
                    // keep the reference as written (minus sheet and $): an area may be a cell, a range,
                    // or whole columns/rows (A:C, 1:3); #REF! areas are dropped
                    const areas = definedName.ranges
                        .map((range) => range.slice(range.lastIndexOf("!") + 1).replace(/\$/g, ""))
                        .filter((area) => area && !area.startsWith("#"));
                    if (worksheet.pageSetup.printArea) {
                        // legacy files may hold several Print_Area names for the same sheet
                        areas.unshift(worksheet.pageSetup.printArea);
                    }
                    worksheet.pageSetup.printArea = areas.join("&&");
                }
            } else if (definedName.name === "_xlnm.Print_Titles") {
                worksheet = worksheets[definedName.localSheetId];
                if (worksheet) {
                    if (!worksheet.pageSetup) {
                        worksheet.pageSetup = {};
                    }

                    const rangeString = definedName.ranges.join(",");

                    const dollarRegex = /\$/g;

                    const rowRangeRegex = /\$\d+:\$\d+/;
                    const rowRangeMatches = rangeString.match(rowRangeRegex);

                    if (rowRangeMatches && rowRangeMatches.length) {
                        const range = rowRangeMatches[0];
                        worksheet.pageSetup.printTitlesRow = range.replace(dollarRegex, "");
                    }

                    const columnRangeRegex = /\$[A-Z]+:\$[A-Z]+/;
                    const columnRangeMatches = rangeString.match(columnRangeRegex);

                    if (columnRangeMatches && columnRangeMatches.length) {
                        const range = columnRangeMatches[0];
                        worksheet.pageSetup.printTitlesColumn = range.replace(dollarRegex, "");
                    }
                }
            } else {
                definedNames.push(definedName);
            }
        });
        model.definedNames = definedNames;

        // used by sheets to build their image models
        model.media.forEach((media, i) => {
            media.index = i;
        });
    }
}

WorkbookXform.WORKBOOK_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "x15",
    "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
};
WorkbookXform.STATIC_XFORMS = {
    fileVersion: new StaticXform({
        tag: "fileVersion",
        $: { appName: "xl", lastEdited: 5, lowestEdited: 5, rupBuild: 9303 },
    }),
};

module.exports = WorkbookXform;
