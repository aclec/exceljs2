/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import * as fs from "fs";
import { PassThrough } from "stream";
import { ZipWriter } from "../utils/zip-stream";
import { readZip } from "../utils/zip-reader";
import StreamBuf from "../utils/stream-buf";
import RelType from "./rel-type";

import utils from "../utils/utils";
import XmlStream from "../utils/xml-stream";
import { bufferToString } from "../utils/browser-buffer-decode";
import { checkMaxMergedCells } from "../utils/merge-limits";
import Pacer from "../utils/pacer";

import StylesXform from "./xform/style/styles-xform";

import CoreXform from "./xform/core/core-xform";
import SharedStringsXform from "./xform/strings/shared-strings-xform";
import RelationshipsXform from "./xform/core/relationships-xform";
import ContentTypesXform from "./xform/core/content-types-xform";
import AppXform from "./xform/core/app-xform";
import WorkbookXform from "./xform/book/workbook-xform";
import WorksheetXform from "./xform/sheet/worksheet-xform";
import DrawingXform from "./xform/drawing/drawing-xform";
import TableXform from "./xform/table/table-xform";
import PivotCacheRecordsXform from "./xform/pivot-table/pivot-cache-records-xform";
import PivotCacheDefinitionXform from "./xform/pivot-table/pivot-cache-definition-xform";
import PivotTableXform from "./xform/pivot-table/pivot-table-xform";
import CommentsXform from "./xform/comment/comments-xform";
import VmlNotesXform from "./xform/comment/vml-notes-xform";

import theme1Xml from "./xml/theme1";

function fsReadFileAsync(filename, options) {
    return new Promise((resolve, reject) => {
        fs.readFile(filename, options, (error, data) => {
            if (error) {
                reject(error);
            } else {
                resolve(data);
            }
        });
    });
}

// image file extensions: alphanumeric only (no path separators, no XML)
const IMAGE_EXTENSION = /^[a-z0-9]{1,5}$/i;

// the text of an XML part, in chunks for the SAX parser, pausing between them (utils/pacer)
async function* xmlChunks(content, pacer) {
    const chunkSize = 16 * 1024;
    for (let i = 0; i < content.length; i += chunkSize) {
        yield content.substring(i, i + chunkSize);
        const pause = pacer.pause();
        if (pause) {
            await pause;
        }
    }
}

// The parts that load() reads (keep in sync with its switch): the others (pivot tables, charts, custom XML, printer
// settings...) are not inflated.
const READ_PARTS = new Set([
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "docProps/app.xml",
    "docProps/core.xml",
]);
const READ_PART_PATTERNS = [
    /xl\/worksheets\/sheet(\d+)[.]xml/,
    /xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/,
    /xl\/theme\/([a-zA-Z0-9]+)[.]xml/,
    /xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/,
    /xl\/drawings\/([a-zA-Z0-9]+)[.]xml/,
    /xl\/(comments\d+|comments\/comment\d+)[.]xml$/,
    /xl\/tables\/(table\d+)[.]xml/,
    /xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/,
    /xl\/drawings\/(vmlDrawing\d+|commentsDrawing\d+)[.]vml$/,
];

function isReadPart(name) {
    return READ_PARTS.has(name) || READ_PART_PATTERNS.some((pattern) => pattern.test(name));
}

class XLSX {
    constructor(workbook) {
        this.workbook = workbook;
    }

    // ===============================================================================
    // Workbook
    // =========================================================================
    // Read

    async readFile(filename, options) {
        if (!(await utils.fs.exists(filename))) {
            throw new Error(`File not found: ${filename}`);
        }
        const stream = fs.createReadStream(filename);
        try {
            const workbook = await this.read(stream, options);
            stream.close();
            return workbook;
        } catch (error) {
            stream.close();
            throw error;
        }
    }

    parseRels(stream) {
        const xform = new RelationshipsXform();
        return xform.parseStream(stream);
    }

    parseWorkbook(stream) {
        const xform = new WorkbookXform();
        return xform.parseStream(stream);
    }

    parseSharedStrings(stream) {
        const xform = new SharedStringsXform();
        return xform.parseStream(stream);
    }

    reconcile(model, options) {
        Pacer.runSteps(this._reconcileSteps(model, options));
    }

    // reconcile() worksheet row by worksheet row: the caller may pause between rows
    *_reconcileSteps(model, options) {
        const workbookXform = new WorkbookXform();
        const worksheetXform = new WorksheetXform(options);
        const drawingXform = new DrawingXform();
        const tableXform = new TableXform();

        workbookXform.reconcile(model);

        // reconcile drawings with their rels
        const drawingOptions = {
            media: model.media,
            mediaIndex: model.mediaIndex,
        };
        Object.keys(model.drawings).forEach((name) => {
            const drawing = model.drawings[name];
            // part names and relationship ids come from the file: ignore inherited members
            const drawingRel = Object.prototype.hasOwnProperty.call(model.drawingRels, name) && model.drawingRels[name];
            if (drawingRel) {
                drawingOptions.rels = drawingRel.reduce((o, rel) => {
                    o[rel.Id] = rel;
                    return o;
                }, Object.create(null));
                (drawing.anchors || []).forEach((anchor) => {
                    const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
                    if (hyperlinks && drawingOptions.rels[hyperlinks.rId]) {
                        hyperlinks.hyperlink = drawingOptions.rels[hyperlinks.rId].Target;
                        delete hyperlinks.rId;
                    }
                });
                drawingXform.reconcile(drawing, drawingOptions);
            }
        });

        // reconcile tables with the default styles
        const tableOptions = {
            styles: model.styles,
        };
        Object.values(model.tables).forEach((table) => {
            tableXform.reconcile(table, tableOptions);
        });

        const sheetOptions = {
            styles: model.styles,
            sharedStrings: model.sharedStrings,
            media: model.media,
            mediaIndex: model.mediaIndex,
            date1904: model.properties && model.properties.date1904,
            drawings: model.drawings,
            comments: model.comments,
            tables: model.tables,
            vmlDrawings: model.vmlDrawings,
            // the parts already claimed by a worksheet (see WorksheetXform.reconcile)
            claimedParts: new Set(),
        };
        for (const worksheet of model.worksheets) {
            worksheet.relationships = model.worksheetRels[worksheet.sheetNo];
            yield* worksheetXform.reconcileSteps(worksheet, sheetOptions);
        }

        // delete unnecessary parts
        delete model.worksheetHash;
        delete model.worksheetRels;
        delete model.globalRels;
        delete model.sharedStrings;
        delete model.workbookRels;
        delete model.sheetDefs;
        delete model.styles;
        delete model.mediaIndex;
        delete model.drawings;
        delete model.drawingRels;
        delete model.vmlDrawings;
        delete model.loadBudget;
    }

    async _processWorksheetEntry(stream, model, sheetNo, options, path) {
        const xform = new WorksheetXform({ ...options, loadBudget: model.loadBudget });
        const worksheet = await xform.parseStream(stream);
        worksheet.sheetNo = sheetNo;
        model.worksheetHash[path] = worksheet;
        model.worksheets.push(worksheet);
    }

    async _processCommentEntry(stream, model, entryName) {
        const xform = new CommentsXform();
        model.comments[entryName] = await xform.parseStream(stream);
    }

    async _processTableEntry(stream, model, entryName) {
        const xform = new TableXform();
        model.tables[entryName] = await xform.parseStream(stream);
    }

    async _processWorksheetRelsEntry(stream, model, sheetNo) {
        const xform = new RelationshipsXform();
        const relationships = await xform.parseStream(stream);
        model.worksheetRels[sheetNo] = relationships;
    }

    async _processMediaEntry(entry, model, filename) {
        const lastDot = filename.lastIndexOf(".");
        // if we can't determine extension, ignore it
        if (lastDot >= 1) {
            const extension = filename.substr(lastDot + 1);
            const name = filename.substr(0, lastDot);
            await new Promise((resolve, reject) => {
                const streamBuf = new StreamBuf();
                streamBuf.on("finish", () => {
                    model.mediaIndex[filename] = model.media.length;
                    model.mediaIndex[name] = model.media.length;
                    const medium = {
                        type: "image",
                        name,
                        extension,
                        buffer: streamBuf.toBuffer(),
                    };
                    model.media.push(medium);
                    resolve();
                });
                entry.on("error", (error) => {
                    reject(error);
                });
                entry.pipe(streamBuf);
            });
        }
    }

    async _processDrawingEntry(entry, model, name) {
        const xform = new DrawingXform();
        const drawing = await xform.parseStream(entry);
        // null when the part is not a spreadsheet drawing (e.g. a chart's <c:userShapes>)
        if (drawing) {
            model.drawings[name] = drawing;
        }
    }

    async _processDrawingRelsEntry(entry, model, name) {
        const xform = new RelationshipsXform();
        const relationships = await xform.parseStream(entry);
        model.drawingRels[name] = relationships;
    }

    async _processVmlDrawingEntry(entry, model, entryName) {
        const xform = new VmlNotesXform();
        model.vmlDrawings[entryName] = await xform.parseStream(entry);
    }

    async _processThemeEntry(entry, model, name) {
        await new Promise((resolve, reject) => {
            // TODO: stream entry into buffer and store the xml in the model.themes[]
            const stream = new StreamBuf();
            entry.on("error", reject);
            stream.on("error", reject);
            stream.on("finish", () => {
                model.themes[name] = stream.read().toString();
                resolve();
            });
            entry.pipe(stream);
        });
    }

    /**
     * @deprecated since version 4.0. You should use `#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
     */
    createInputStream() {
        throw new Error(
            "`XLSX#createInputStream` is deprecated. You should use `XLSX#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md"
        );
    }

    async read(stream, options) {
        // TODO: Remove once node v8 is deprecated
        // Detect and upgrade old streams
        if (!stream[Symbol.asyncIterator] && stream.pipe) {
            stream = stream.pipe(new PassThrough());
        }
        const chunks = [];
        for await (const chunk of stream) {
            chunks.push(chunk);
        }
        return this.load(Buffer.concat(chunks), options);
    }

    async load(data, options) {
        let buffer;
        if (options && options.base64) {
            buffer = Buffer.from(data.toString(), "base64");
        } else {
            buffer = data;
        }

        const model = {
            worksheets: [],
            worksheetHash: {},
            worksheetRels: [],
            themes: {},
            media: [],
            mediaIndex: {},
            drawings: {},
            drawingRels: {},
            comments: {},
            tables: {},
            vmlDrawings: {},
        };
        // the cap on merged cells applies to the whole workbook, and is kept by the loaded workbook
        const maxMergedCells =
            options && options.maxMergedCells !== undefined ? checkMaxMergedCells(options.maxMergedCells) : this.workbook.maxMergedCells;
        model.loadBudget = WorksheetXform.createLoadBudget(maxMergedCells);

        const pacer = new Pacer();

        // like JSZip did, accept a Blob (e.g. a File from an <input type="file">)
        if (typeof Blob !== "undefined" && buffer instanceof Blob) {
            buffer = new Uint8Array(await buffer.arrayBuffer());
        }
        // the limits (maxEntrySize...) are checked while inflating
        for (const entry of readZip(buffer, options)) {
            const pause = pacer.pause();
            if (pause) {
                await pause;
            }
            if (!entry.dir) {
                let entryName = entry.name;
                if (entryName[0] === "/") {
                    entryName = entryName.substr(1);
                }
                if (!isReadPart(entryName)) {
                    continue;
                }
                const data = await entry.data();
                let stream;
                if (
                    entryName.match(/xl\/media\//) ||
                    // themes are not parsed as stream
                    entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/)
                ) {
                    stream = new PassThrough();
                    stream.write(Buffer.from(data.buffer, data.byteOffset, data.byteLength));
                    stream.end();
                } else {
                    let content;
                    // https://www.npmjs.com/package/process
                    if (process.browser) {
                        // running in browser, use TextDecoder if possible
                        content = bufferToString(data);
                    } else {
                        // running in node.js
                        content = Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString("utf8");
                    }
                    stream = xmlChunks(content, pacer);
                }
                switch (entryName) {
                    case "_rels/.rels":
                        model.globalRels = await this.parseRels(stream);
                        break;

                    case "xl/workbook.xml": {
                        const workbook = await this.parseWorkbook(stream);
                        model.sheets = workbook.sheets;
                        model.definedNames = workbook.definedNames;
                        model.views = workbook.views;
                        model.properties = workbook.properties;
                        model.calcProperties = workbook.calcProperties;
                        break;
                    }

                    case "xl/_rels/workbook.xml.rels":
                        model.workbookRels = await this.parseRels(stream);
                        break;

                    case "xl/sharedStrings.xml":
                        model.sharedStrings = new SharedStringsXform();
                        await model.sharedStrings.parseStream(stream);
                        break;

                    case "xl/styles.xml":
                        model.styles = new StylesXform();
                        await model.styles.parseStream(stream);
                        break;

                    case "docProps/app.xml": {
                        const appXform = new AppXform();
                        const appProperties = await appXform.parseStream(stream);
                        // undefined when the part has no <Properties> root (e.g. a prefixed <ep:Properties>)
                        if (appProperties) {
                            model.company = appProperties.company;
                            model.manager = appProperties.manager;
                        }
                        break;
                    }

                    case "docProps/core.xml": {
                        const coreXform = new CoreXform();
                        const coreProperties = await coreXform.parseStream(stream);
                        Object.assign(model, coreProperties);
                        break;
                    }

                    default: {
                        let match = entryName.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
                        if (match) {
                            await this._processWorksheetEntry(stream, model, match[1], options, entryName);
                            break;
                        }
                        match = entryName.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
                        if (match) {
                            await this._processWorksheetRelsEntry(stream, model, match[1]);
                            break;
                        }
                        match = entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
                        if (match) {
                            await this._processThemeEntry(stream, model, match[1]);
                            break;
                        }
                        match = entryName.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
                        if (match) {
                            await this._processMediaEntry(stream, model, match[1]);
                            break;
                        }
                        match = entryName.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
                        if (match) {
                            await this._processDrawingEntry(stream, model, match[1]);
                            break;
                        }
                        // xl/comments1.xml (Excel), xl/comments/comment1.xml (openpyxl)
                        match = entryName.match(/xl\/(comments\d+|comments\/comment\d+)[.]xml$/);
                        if (match) {
                            await this._processCommentEntry(stream, model, entryName);
                            break;
                        }
                        match = entryName.match(/xl\/tables\/(table\d+)[.]xml/);
                        if (match) {
                            await this._processTableEntry(stream, model, entryName);
                            break;
                        }
                        match = entryName.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
                        if (match) {
                            await this._processDrawingRelsEntry(stream, model, match[1]);
                            break;
                        }
                        // vmlDrawing1.vml (Excel), commentsDrawing1.vml (openpyxl)
                        match = entryName.match(/xl\/drawings\/(vmlDrawing\d+|commentsDrawing\d+)[.]vml$/);
                        if (match) {
                            await this._processVmlDrawingEntry(stream, model, entryName);
                            break;
                        }
                    }
                }
            }
        }

        await pacer.run(this._reconcileSteps(model, options));

        // apply model
        this.workbook.maxMergedCells = maxMergedCells;
        this.workbook.model = model;
        return this.workbook;
    }

    // =========================================================================
    // Write

    async addMedia(zip, model) {
        await Promise.all(
            model.media.map(async (medium) => {
                if (medium.type === "image") {
                    const filename = `xl/media/${medium.name}.${medium.extension}`;
                    if (medium.filename) {
                        const data = await fsReadFileAsync(medium.filename);
                        return zip.append(data, { name: filename });
                    }
                    if (medium.buffer) {
                        return zip.append(medium.buffer, { name: filename });
                    }
                    if (medium.base64) {
                        const dataimg64 = medium.base64;
                        const content = dataimg64.substring(dataimg64.indexOf(",") + 1);
                        return zip.append(content, { name: filename, base64: true });
                    }
                }
                throw new Error("Unsupported media");
            })
        );
    }

    addDrawings(zip, model) {
        const drawingXform = new DrawingXform();
        const relsXform = new RelationshipsXform();

        model.worksheets.forEach((worksheet) => {
            const { drawing } = worksheet;
            if (drawing) {
                drawingXform.prepare(drawing, {});
                let xml = drawingXform.toXml(drawing);
                zip.append(xml, { name: `xl/drawings/${drawing.name}.xml` });

                xml = relsXform.toXml(drawing.rels);
                zip.append(xml, { name: `xl/drawings/_rels/${drawing.name}.xml.rels` });
            }
        });
    }

    addTables(zip, model) {
        const tableXform = new TableXform();

        model.worksheets.forEach((worksheet) => {
            const { tables } = worksheet;
            tables.forEach((table) => {
                tableXform.prepare(table, {});
                const tableXml = tableXform.toXml(table);
                zip.append(tableXml, { name: `xl/tables/${table.target}` });
            });
        });
    }

    addPivotTables(zip, model) {
        if (!model.pivotTables.length) return;

        const pivotTable = model.pivotTables[0];

        const pivotCacheRecordsXform = new PivotCacheRecordsXform();
        const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
        const pivotTableXform = new PivotTableXform();
        const relsXform = new RelationshipsXform();

        // pivot cache records
        // --------------------------------------------------
        // copy of the source data.
        //
        // Note: cells in the columns of the source data which are part
        // of the "rows" or "columns" of the pivot table configuration are
        // replaced by references to their __cache field__ identifiers.
        // See "pivot cache definition" below.

        let xml = pivotCacheRecordsXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotCache/pivotCacheRecords1.xml" });

        // pivot cache definition
        // --------------------------------------------------
        // cache source (source data):
        //    ref="A1:E7" on sheet="Sheet1"
        // cache fields:
        //    - 0: "A" (a1, a2, a3)
        //    - 1: "B" (b1, b2)
        //    - ...

        xml = pivotCacheDefinitionXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotCache/pivotCacheDefinition1.xml" });

        xml = relsXform.toXml([
            {
                Id: "rId1",
                Type: XLSX.RelType.PivotCacheRecords,
                Target: "pivotCacheRecords1.xml",
            },
        ]);
        zip.append(xml, { name: "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels" });

        // pivot tables (on destination worksheet)
        // --------------------------------------------------
        // location: ref="A3:E15"
        // pivotFields
        // rowFields and rowItems
        // colFields and colItems
        // dataFields
        // pivotTableStyleInfo

        xml = pivotTableXform.toXml(pivotTable);
        zip.append(xml, { name: "xl/pivotTables/pivotTable1.xml" });

        xml = relsXform.toXml([
            {
                Id: "rId1",
                Type: XLSX.RelType.PivotCacheDefinition,
                Target: "../pivotCache/pivotCacheDefinition1.xml",
            },
        ]);
        zip.append(xml, { name: "xl/pivotTables/_rels/pivotTable1.xml.rels" });
    }

    async addContentTypes(zip, model) {
        const xform = new ContentTypesXform();
        const xml = xform.toXml(model);
        zip.append(xml, { name: "[Content_Types].xml" });
    }

    async addApp(zip, model) {
        const xform = new AppXform();
        const xml = xform.toXml(model);
        zip.append(xml, { name: "docProps/app.xml" });
    }

    async addCore(zip, model) {
        const coreXform = new CoreXform();
        zip.append(coreXform.toXml(model), { name: "docProps/core.xml" });
    }

    async addThemes(zip, model) {
        // workbook rels and content types always reference theme1.xml, so it must be written
        // even when the workbook was read from a file without a theme
        const themes = { theme1: theme1Xml, ...model.themes };
        Object.keys(themes).forEach((name) => {
            const xml = themes[name];
            const path = `xl/theme/${name}.xml`;
            zip.append(xml, { name: path });
        });
    }

    async addOfficeRels(zip) {
        const xform = new RelationshipsXform();
        const xml = xform.toXml([
            { Id: "rId1", Type: XLSX.RelType.OfficeDocument, Target: "xl/workbook.xml" },
            { Id: "rId2", Type: XLSX.RelType.CoreProperties, Target: "docProps/core.xml" },
            { Id: "rId3", Type: XLSX.RelType.ExtenderProperties, Target: "docProps/app.xml" },
        ]);
        zip.append(xml, { name: "_rels/.rels" });
    }

    async addWorkbookRels(zip, model) {
        let count = 1;
        const relationships = [
            { Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: "styles.xml" },
            { Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: "theme/theme1.xml" },
        ];
        if (model.sharedStrings.count) {
            relationships.push({
                Id: `rId${count++}`,
                Type: XLSX.RelType.SharedStrings,
                Target: "sharedStrings.xml",
            });
        }
        if ((model.pivotTables || []).length) {
            const pivotTable = model.pivotTables[0];
            pivotTable.rId = `rId${count++}`;
            relationships.push({
                Id: pivotTable.rId,
                Type: XLSX.RelType.PivotCacheDefinition,
                Target: "pivotCache/pivotCacheDefinition1.xml",
            });
        }
        model.worksheets.forEach((worksheet) => {
            worksheet.rId = `rId${count++}`;
            relationships.push({
                Id: worksheet.rId,
                Type: XLSX.RelType.Worksheet,
                Target: `worksheets/sheet${worksheet.id}.xml`,
            });
        });
        const xform = new RelationshipsXform();
        const xml = xform.toXml(relationships);
        zip.append(xml, { name: "xl/_rels/workbook.xml.rels" });
    }

    async addSharedStrings(zip, model) {
        if (model.sharedStrings && model.sharedStrings.count) {
            zip.append(model.sharedStrings.xml, { name: "xl/sharedStrings.xml" });
        }
    }

    async addStyles(zip, model) {
        const { xml } = model.styles;
        if (xml) {
            zip.append(xml, { name: "xl/styles.xml" });
        }
    }

    async addWorkbook(zip, model) {
        const xform = new WorkbookXform();
        zip.append(xform.toXml(model), { name: "xl/workbook.xml" });
    }

    async addWorksheets(zip, model, pacer = new Pacer()) {
        // preparation phase
        const worksheetXform = new WorksheetXform();
        const relationshipsXform = new RelationshipsXform();
        const commentsXform = new CommentsXform();
        const vmlNotesXform = new VmlNotesXform();

        // the blocks of VML shape ids claimed by the notes of each worksheet (VmlNotesXform)
        let nextShapeBlock = 1;

        // write sheets
        for (const worksheet of model.worksheets) {
            let xmlStream = new XmlStream();
            await pacer.run(worksheetXform.renderSteps(xmlStream, worksheet));
            // appended whole, so that its sizes are in the local header (the zip layer compresses it at once)
            zip.append(xmlStream.xml, { name: `xl/worksheets/sheet${worksheet.id}.xml` });

            if (worksheet.rels && worksheet.rels.length) {
                xmlStream = new XmlStream();
                relationshipsXform.render(xmlStream, worksheet.rels);
                zip.append(xmlStream.xml, { name: `xl/worksheets/_rels/sheet${worksheet.id}.xml.rels` });
            }

            if (worksheet.comments.length > 0) {
                xmlStream = new XmlStream();
                commentsXform.render(xmlStream, worksheet);
                zip.append(xmlStream.xml, { name: `xl/comments${worksheet.id}.xml` });

                xmlStream = new XmlStream();
                worksheet.shapeBlock = nextShapeBlock;
                nextShapeBlock += VmlNotesXform.blockCount(worksheet.comments.length);
                vmlNotesXform.render(xmlStream, worksheet);
                zip.append(xmlStream.xml, { name: `xl/drawings/vmlDrawing${worksheet.id}.vml` });
            }
        }
    }

    async _finalize(zip) {
        // rejects on any failure of the zip writer or of the output stream
        await zip.finalize();
        return this;
    }

    prepareModel(model, options) {
        Pacer.runSteps(this._prepareModelSteps(model, options));
    }

    // prepareModel() worksheet row by worksheet row: the caller may pause between rows
    *_prepareModelSteps(model, options) {
        // ensure following properties have sane values
        model.creator = model.creator || "ExcelJS";
        model.lastModifiedBy = model.lastModifiedBy || "ExcelJS";
        model.created = model.created || new Date();
        model.modified = model.modified || new Date();

        model.useSharedStrings = options.useSharedStrings !== undefined ? options.useSharedStrings : true;
        model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

        // Manage the shared strings
        model.sharedStrings = new SharedStringsXform();

        // add a style manager to handle cell formats, fonts, etc.
        model.styles = model.useStyles ? new StylesXform(true) : new StylesXform.Mock();

        // prepare all of the things before the render
        const workbookXform = new WorkbookXform();
        const worksheetXform = new WorksheetXform();

        workbookXform.prepare(model);

        const worksheetOptions = {
            sharedStrings: model.sharedStrings,
            styles: model.styles,
            date1904: model.properties.date1904,
            drawingsCount: 0,
            media: model.media,
        };
        worksheetOptions.drawings = model.drawings = [];
        worksheetOptions.commentRefs = model.commentRefs = [];
        let tableCount = 0;
        model.tables = [];
        for (const worksheet of model.worksheets) {
            // assign unique filenames to tables
            worksheet.tables.forEach((table) => {
                tableCount++;
                table.target = `table${tableCount}.xml`;
                table.id = tableCount;
                model.tables.push(table);
            });

            yield* worksheetXform.prepareSteps(worksheet, worksheetOptions);
        }

        // TODO: workbook drawing list
    }

    async write(stream, options) {
        options = options || {};
        const { model } = this.workbook;
        // before anything is written: the extension ends up in a zip entry name and in XML
        model.media.forEach((medium) => {
            if (medium.type === "image" && !IMAGE_EXTENSION.test(medium.extension)) {
                throw new Error(`Invalid image extension: ${JSON.stringify(medium.extension)}`);
            }
        });
        const zip = new ZipWriter(options.zip);
        zip.pipe(stream);

        const pacer = new Pacer();
        await pacer.run(this._prepareModelSteps(model, options));

        // render
        await this.addContentTypes(zip, model);
        await this.addOfficeRels(zip, model);
        await this.addWorkbookRels(zip, model);
        await this.addWorksheets(zip, model, pacer);
        await this.addSharedStrings(zip, model); // always after worksheets
        await this.addDrawings(zip, model);
        await this.addTables(zip, model);
        await this.addPivotTables(zip, model);
        await Promise.all([this.addThemes(zip, model), this.addStyles(zip, model)]);
        await this.addMedia(zip, model);
        await Promise.all([this.addApp(zip, model), this.addCore(zip, model)]);
        await this.addWorkbook(zip, model);
        // a stream reports a failed write on a later tick: let it do so before the archive is completed. An output that
        // cannot report errors (the collector of writeBuffer) needs no turn, so that writing a small workbook does not wait for
        // a timer (fake timers of test frameworks never run it)
        if (typeof stream.on === "function") {
            await Pacer.nextTurn();
        }
        return this._finalize(zip);
    }

    writeFile(filename, options) {
        const stream = fs.createWriteStream(filename);

        return new Promise((resolve, reject) => {
            stream.on("finish", () => {
                resolve();
            });
            stream.on("error", (error) => {
                reject(error);
            });

            this.write(stream, options)
                .then(() => {
                    stream.end();
                })
                .catch((err) => {
                    reject(err);
                });
        });
    }

    async writeBuffer(options) {
        const chunks = [];
        const collector = {
            write(chunk) {
                chunks.push(chunk);
            },
            end() {},
        };
        await this.write(collector, options);
        return Buffer.concat(chunks);
    }
}

XLSX.RelType = RelType;

export default XLSX;
