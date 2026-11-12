const _ = require("../../../utils/under-dash");

const colCache = require("../../../utils/col-cache");
const Range = require("../../../doc/range");
const XmlStream = require("../../../utils/xml-stream");

const RelType = require("../../rel-type");

const Merges = require("./merges");

const BaseXform = require("../base-xform");
const ListXform = require("../list-xform");
const RowXform = require("./row-xform");
const ColXform = require("./col-xform");
const DimensionXform = require("./dimension-xform");
const HyperlinkXform = require("./hyperlink-xform");
const MergeCellXform = require("./merge-cell-xform");
const DataValidationsXform = require("./data-validations-xform");
const SheetPropertiesXform = require("./sheet-properties-xform");
const SheetFormatPropertiesXform = require("./sheet-format-properties-xform");
const SheetViewXform = require("./sheet-view-xform");
const SheetProtectionXform = require("./sheet-protection-xform");
const PageMarginsXform = require("./page-margins-xform");
const PageSetupXform = require("./page-setup-xform");
const PrintOptionsXform = require("./print-options-xform");
const AutoFilterXform = require("./auto-filter-xform");
const PictureXform = require("./picture-xform");
const DrawingXform = require("./drawing-xform");
const TablePartXform = require("./table-part-xform");
const RowBreaksXform = require("./row-breaks-xform");
const HeaderFooterXform = require("./header-footer-xform");
const ConditionalFormattingsXform = require("./cf/conditional-formattings-xform");
const ExtListXform = require("./ext-lst-xform");

class LegacyDrawingXform extends DrawingXform {
    get tag() {
        return "legacyDrawing";
    }
}

function getOwn(object, key) {
    return object && Object.prototype.hasOwnProperty.call(object, key) ? object[key] : undefined;
}

// Parts are keyed by their zip entry name. A worksheet relationship target is relative to
// xl/worksheets/ or absolute from the package root (openpyxl writes /xl/comments/comment1.xml).
function getPart(parts, rel) {
    if (!rel || !rel.Target || rel.TargetMode === "External") {
        return undefined;
    }
    const segments = rel.Target.startsWith("/") ? [] : ["xl", "worksheets"];
    rel.Target.split("/").forEach((segment) => {
        if (segment === "..") {
            segments.pop();
        } else if (segment && segment !== ".") {
            segments.push(segment);
        }
    });
    return getOwn(parts, segments.join("/"));
}

// A <hyperlink> has a relationship (external target), a location (place in the workbook, or
// anchor in the external target), or both. exceljs writes internal links with both, the
// target being the location itself ("Sheet2!A1" or "#'Sheet2'!A1"): don't append it twice.
function getHyperlinkTarget(relTarget, location) {
    if (!location) {
        return relTarget;
    }
    const anchor = location.startsWith("#") ? location : `#${location}`;
    if (!relTarget) {
        return anchor;
    }
    if (relTarget === location || relTarget.includes("#")) {
        return relTarget;
    }
    return `${relTarget}${anchor}`;
}

// The r attribute of <row> and <c> is optional (ECMA-376 CT_Row, CT_Cell): a row without it
// follows the previous row, a cell without it follows the previous cell of its row.
function inferMissingReferences(row, previousRowNumber) {
    if (!(row.number > 0)) {
        const addressed = row.cells.find((cell) => cell.address);
        row.number = addressed ? colCache.decodeAddress(addressed.address).row : previousRowNumber + 1;
    }
    row.cells.forEach((cell, index) => {
        if (!cell.address) {
            const previousCol = index ? colCache.decodeAddress(row.cells[index - 1].address).col : 0;
            cell.address = colCache.encodeAddress(row.number, previousCol + 1);
        }
    });
}

// Every cell of a merge is materialised when a worksheet is loaded. The worksheet caps its own merged cells, the
// load budget caps those of the whole workbook: one sheet may hold a full-column merge, but not every sheet.
const MAX_LOADED_MERGED_CELLS = colCache.MAX_ROWS;

function spendMergedCells(loadBudget, mergeCells) {
    (mergeCells || []).forEach((ref) => {
        loadBudget.mergedCells -= new Range(ref).count;
        if (loadBudget.mergedCells < 0) {
            throw new Error(`Cannot merge ${ref}: the workbook would exceed ${MAX_LOADED_MERGED_CELLS} merged cells`);
        }
    });
}

const mergeRule = (rule, extRule) => {
    Object.keys(extRule).forEach((key) => {
        const value = rule[key];
        const extValue = extRule[key];
        if (value === undefined && extValue !== undefined) {
            rule[key] = extValue;
        }
    });
};

const mergeConditionalFormattings = (model, extModel) => {
    // conditional formattings are rendered in worksheet.conditionalFormatting and also in
    // worksheet.extLst.ext.x14:conditionalFormattings
    // some (e.g. dataBar) are even spread across both!
    if (!extModel || !extModel.length) {
        return model;
    }
    if (!model || !model.length) {
        return extModel;
    }

    // index model rules by x14Id. Both keys come from the file: without a prototype, an id such
    // as "__proto__" cannot reach Object.prototype
    const cfMap = Object.create(null);
    const ruleMap = Object.create(null);
    model.forEach((cf) => {
        cfMap[cf.ref] = cf;
        cf.rules.forEach((rule) => {
            const { x14Id } = rule;
            if (x14Id) {
                ruleMap[x14Id] = rule;
            }
        });
    });

    extModel.forEach((extCf) => {
        extCf.rules.forEach((extRule) => {
            const rule = ruleMap[extRule.x14Id];
            if (rule) {
                // merge with matching rule
                mergeRule(rule, extRule);
            } else if (cfMap[extCf.ref]) {
                // reuse existing cf ref
                cfMap[extCf.ref].rules.push(extRule);
            } else {
                // create new cf
                model.push({
                    ref: extCf.ref,
                    rules: [extRule],
                });
            }
        });
    });

    // need to cope with rules in extModel that don't exist in model
    return model;
};

class WorkSheetXform extends BaseXform {
    constructor(options) {
        super();

        const { maxRows, maxCols, ignoreNodes, loadBudget } = options || {};

        this.ignoreNodes = ignoreNodes || [];
        this.loadBudget = loadBudget;

        this.map = {
            sheetPr: new SheetPropertiesXform(),
            dimension: new DimensionXform(),
            sheetViews: new ListXform({
                tag: "sheetViews",
                count: false,
                childXform: new SheetViewXform(),
            }),
            sheetFormatPr: new SheetFormatPropertiesXform(),
            cols: new ListXform({ tag: "cols", count: false, childXform: new ColXform() }),
            sheetData: new ListXform({
                tag: "sheetData",
                count: false,
                empty: true,
                childXform: new RowXform({ maxItems: maxCols }),
                maxItems: maxRows,
            }),
            autoFilter: new AutoFilterXform(),
            mergeCells: new ListXform({ tag: "mergeCells", count: true, childXform: new MergeCellXform() }),
            rowBreaks: new RowBreaksXform(),
            hyperlinks: new ListXform({
                tag: "hyperlinks",
                count: false,
                childXform: new HyperlinkXform(),
            }),
            pageMargins: new PageMarginsXform(),
            dataValidations: new DataValidationsXform({ loadBudget }),
            pageSetup: new PageSetupXform(),
            headerFooter: new HeaderFooterXform(),
            printOptions: new PrintOptionsXform(),
            picture: new PictureXform(),
            drawing: new DrawingXform(),
            legacyDrawing: new LegacyDrawingXform(),
            sheetProtection: new SheetProtectionXform(),
            tableParts: new ListXform({ tag: "tableParts", count: true, childXform: new TablePartXform() }),
            conditionalFormatting: new ConditionalFormattingsXform(),
            extLst: new ExtListXform(),
        };
    }

    prepare(model, options) {
        options.merges = new Merges();
        model.hyperlinks = options.hyperlinks = [];
        model.comments = options.comments = [];

        options.formulae = {};
        options.siFormulae = 0;
        this.map.cols.prepare(model.cols, options);
        this.map.sheetData.prepare(model.rows, options);
        this.map.conditionalFormatting.prepare(model.conditionalFormattings, options);

        model.mergeCells = options.merges.mergeCells;

        // prepare relationships
        const rels = (model.rels = []);

        function nextRid(r) {
            return `rId${r.length + 1}`;
        }

        model.hyperlinks.forEach((hyperlink) => {
            const rId = nextRid(rels);
            hyperlink.rId = rId;
            rels.push({
                Id: rId,
                Type: RelType.Hyperlink,
                Target: hyperlink.target,
                TargetMode: "External",
            });
        });

        // prepare comment relationships
        if (model.comments.length > 0) {
            const comment = {
                Id: nextRid(rels),
                Type: RelType.Comments,
                Target: `../comments${model.id}.xml`,
            };
            rels.push(comment);
            const vmlDrawing = {
                Id: nextRid(rels),
                Type: RelType.VmlDrawing,
                Target: `../drawings/vmlDrawing${model.id}.vml`,
            };
            rels.push(vmlDrawing);

            model.comments.forEach((item) => {
                item.refAddress = colCache.decodeAddress(item.ref);
            });

            options.commentRefs.push({
                commentName: `comments${model.id}`,
                vmlDrawing: `vmlDrawing${model.id}`,
            });
        }

        // one image relationship per distinct image of this sheet's drawing
        const imageRIds = {};
        let bookImage;
        model.media.forEach((medium) => {
            if (medium.type === "background") {
                const rId = nextRid(rels);
                bookImage = options.media[medium.imageId];
                rels.push({
                    Id: rId,
                    Type: RelType.Image,
                    Target: `../media/${bookImage.name}.${bookImage.extension}`,
                });
                model.background = {
                    rId,
                };
                model.image = options.media[medium.imageId];
            } else if (medium.type === "image") {
                let { drawing } = model;
                bookImage = options.media[medium.imageId];
                if (!drawing) {
                    drawing = model.drawing = {
                        rId: nextRid(rels),
                        name: `drawing${++options.drawingsCount}`,
                        anchors: [],
                        rels: [],
                    };
                    options.drawings.push(drawing);
                    rels.push({
                        Id: drawing.rId,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                        Target: `../drawings/${drawing.name}.xml`,
                    });
                }
                let rIdImage = imageRIds[medium.imageId];
                if (!rIdImage) {
                    rIdImage = imageRIds[medium.imageId] = nextRid(drawing.rels);
                    drawing.rels.push({
                        Id: rIdImage,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        Target: `../media/${bookImage.name}.${bookImage.extension}`,
                    });
                }

                const anchor = {
                    picture: {
                        rId: rIdImage,
                    },
                    range: medium.range,
                };
                if (medium.hyperlinks && medium.hyperlinks.hyperlink) {
                    const rIdHyperLink = nextRid(drawing.rels);
                    anchor.picture.hyperlinks = {
                        tooltip: medium.hyperlinks.tooltip,
                        rId: rIdHyperLink,
                    };
                    drawing.rels.push({
                        Id: rIdHyperLink,
                        Type: RelType.Hyperlink,
                        Target: medium.hyperlinks.hyperlink,
                        TargetMode: "External",
                    });
                }
                drawing.anchors.push(anchor);
            }
        });

        // prepare tables
        model.tables.forEach((table) => {
            // relationships
            const rId = nextRid(rels);
            table.rId = rId;
            rels.push({
                Id: rId,
                Type: RelType.Table,
                Target: `../tables/${table.target}`,
            });

            // dynamic styles
            table.columns.forEach((column) => {
                const { style } = column;
                if (style) {
                    column.dxfId = options.styles.addDxfStyle(style);
                }
            });
        });

        // prepare pivot tables
        if ((model.pivotTables || []).length) {
            rels.push({
                Id: nextRid(rels),
                Type: RelType.PivotTable,
                Target: "../pivotTables/pivotTable1.xml",
            });
        }

        // prepare ext items
        this.map.extLst.prepare(model, options);
    }

    render(xmlStream, model) {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("worksheet", WorkSheetXform.WORKSHEET_ATTRIBUTES);

        const sheetFormatPropertiesModel = model.properties
            ? {
                  defaultRowHeight: model.properties.defaultRowHeight,
                  dyDescent: model.properties.dyDescent,
                  outlineLevelCol: model.properties.outlineLevelCol,
                  outlineLevelRow: model.properties.outlineLevelRow,
              }
            : undefined;
        if (model.properties && model.properties.defaultColWidth) {
            sheetFormatPropertiesModel.defaultColWidth = model.properties.defaultColWidth;
        }
        const sheetPropertiesModel = {
            outlineProperties: model.properties && model.properties.outlineProperties,
            tabColor: model.properties && model.properties.tabColor,
            pageSetup:
                model.pageSetup && model.pageSetup.fitToPage
                    ? {
                          fitToPage: model.pageSetup.fitToPage,
                      }
                    : undefined,
        };
        const pageMarginsModel = model.pageSetup && model.pageSetup.margins;
        const printOptionsModel = {
            showRowColHeaders: model.pageSetup && model.pageSetup.showRowColHeaders,
            showGridLines: model.pageSetup && model.pageSetup.showGridLines,
            horizontalCentered: model.pageSetup && model.pageSetup.horizontalCentered,
            verticalCentered: model.pageSetup && model.pageSetup.verticalCentered,
        };
        const sheetProtectionModel = model.sheetProtection;

        this.map.sheetPr.render(xmlStream, sheetPropertiesModel);
        this.map.dimension.render(xmlStream, model.dimensions);
        this.map.sheetViews.render(xmlStream, model.views);
        this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel);
        this.map.cols.render(xmlStream, model.cols);
        this.map.sheetData.render(xmlStream, model.rows);
        this.map.sheetProtection.render(xmlStream, sheetProtectionModel); // Note: must be after sheetData and before autoFilter
        this.map.autoFilter.render(xmlStream, model.autoFilter);
        this.map.mergeCells.render(xmlStream, model.mergeCells);
        this.map.conditionalFormatting.render(xmlStream, model.conditionalFormattings); // Note: must be before dataValidations
        this.map.dataValidations.render(xmlStream, model.dataValidations);

        // For some reason hyperlinks have to be after the data validations
        this.map.hyperlinks.render(xmlStream, model.hyperlinks);

        this.map.printOptions.render(xmlStream, printOptionsModel); // Note: must be before pageMargins
        this.map.pageMargins.render(xmlStream, pageMarginsModel);
        this.map.pageSetup.render(xmlStream, model.pageSetup);
        this.map.headerFooter.render(xmlStream, model.headerFooter);
        this.map.rowBreaks.render(xmlStream, model.rowBreaks);
        // CT_Worksheet order: drawing, legacyDrawing, picture, tableParts, extLst
        this.map.drawing.render(xmlStream, model.drawing); // Note: must be after rowBreaks
        if (model.rels) {
            // add a <legacyDrawing /> node for the comments (vml drawing)
            model.rels.forEach((rel) => {
                if (rel.Type === RelType.VmlDrawing) {
                    xmlStream.leafNode("legacyDrawing", { "r:id": rel.Id });
                }
            });
        }
        this.map.picture.render(xmlStream, model.background);
        this.map.tableParts.render(xmlStream, model.tables);

        this.map.extLst.render(xmlStream, model);

        xmlStream.closeNode();
    }

    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        if (node.name === "worksheet") {
            _.each(this.map, (xform) => {
                xform.reset();
            });
            return true;
        }

        if (this.map[node.name] && !this.ignoreNodes.includes(node.name)) {
            this.parser = this.map[node.name];
            this.parser.parseOpen(node);
        }
        return true;
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
            case "worksheet": {
                const properties = this.map.sheetFormatPr.model || {};
                if (this.map.sheetPr.model && this.map.sheetPr.model.tabColor) {
                    properties.tabColor = this.map.sheetPr.model.tabColor;
                }
                if (this.map.sheetPr.model && this.map.sheetPr.model.outlineProperties) {
                    properties.outlineProperties = this.map.sheetPr.model.outlineProperties;
                }
                const sheetProperties = {
                    fitToPage:
                        (this.map.sheetPr.model && this.map.sheetPr.model.pageSetup && this.map.sheetPr.model.pageSetup.fitToPage) || false,
                    margins: this.map.pageMargins.model,
                };
                const pageSetup = Object.assign(sheetProperties, this.map.pageSetup.model, this.map.printOptions.model);
                const conditionalFormattings = mergeConditionalFormattings(
                    this.map.conditionalFormatting.model,
                    this.map.extLst.model && this.map.extLst.model["x14:conditionalFormattings"]
                );
                if (this.loadBudget) {
                    spendMergedCells(this.loadBudget, this.map.mergeCells.model);
                }
                this.model = {
                    dimensions: this.map.dimension.model,
                    cols: this.map.cols.model,
                    rows: this.map.sheetData.model,
                    mergeCells: this.map.mergeCells.model,
                    hyperlinks: this.map.hyperlinks.model,
                    dataValidations: this.map.dataValidations.model,
                    properties,
                    views: this.map.sheetViews.model,
                    pageSetup,
                    headerFooter: this.map.headerFooter.model,
                    background: this.map.picture.model,
                    drawing: this.map.drawing.model,
                    tables: this.map.tableParts.model,
                    conditionalFormattings,
                };

                if (this.map.autoFilter.model) {
                    this.model.autoFilter = this.map.autoFilter.model;
                }
                if (this.map.sheetProtection.model) {
                    this.model.sheetProtection = this.map.sheetProtection.model;
                }
                if (this.map.legacyDrawing.model) {
                    this.model.legacyDrawing = this.map.legacyDrawing.model;
                }

                return false;
            }

            default:
                // not quite sure how we get here!
                return true;
        }
    }

    reconcile(model, options) {
        // options.merges = new Merges();
        // options.merges.reconcile(model.mergeCells, model.rows);
        const relationships = model.relationships || [];
        const rels = relationships.reduce((h, rel) => {
            h[rel.Id] = rel;
            return h;
        }, Object.create(null));
        // a relationship may point at a part that is missing or was not recognised: ignore it
        const findPart = (type, parts) =>
            relationships
                .filter((rel) => rel.Type === type)
                .map((rel) => getPart(parts, rel))
                .find(Boolean);
        const commentsPart = findPart(RelType.Comments, options.comments);
        if (commentsPart) {
            model.comments = commentsPart.comments;
            // the notes vml is the one named by <legacyDrawing> (<legacyDrawingHF> has header/footer
            // images); without that element, fall back to the first vml drawing relationship
            const vmlRel = model.legacyDrawing
                ? rels[model.legacyDrawing.rId]
                : relationships.find((rel) => rel.Type === RelType.VmlDrawing);
            const vmlPart = getPart(options.vmlDrawings, vmlRel);
            if (vmlPart) {
                model.comments.forEach((comment, index) => {
                    comment.note = Object.assign({}, comment.note, vmlPart.comments[index]);
                });
            }
        }
        options.commentsMap = (model.comments || []).reduce((h, comment) => {
            if (comment.ref) {
                h[comment.ref] = comment;
            }
            return h;
        }, Object.create(null));
        options.hyperlinkMap = (model.hyperlinks || []).reduce((h, hyperlink) => {
            const rel = hyperlink.rId && rels[hyperlink.rId];
            const target = getHyperlinkTarget(rel && rel.Target, hyperlink.target);
            if (target) {
                h[hyperlink.address] = target;
            }
            return h;
        }, {});
        // shared formula masters by si (from the file)
        options.formulae = Object.create(null);

        // compact the rows and cells
        model.rows = (model.rows && model.rows.filter(Boolean)) || [];
        let previousRowNumber = 0;
        model.rows.forEach((row) => {
            row.cells = (row.cells && row.cells.filter(Boolean)) || [];
            inferMissingReferences(row, previousRowNumber);
            previousRowNumber = row.number;
        });

        this.map.cols.reconcile(model.cols, options);
        this.map.sheetData.reconcile(model.rows, options);
        this.map.conditionalFormatting.reconcile(model.conditionalFormattings, options);

        model.media = [];
        if (model.drawing) {
            const drawingRel = rels[model.drawing.rId];
            const match = drawingRel && drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
            const drawing = match && getOwn(options.drawings, match[1]);
            if (drawing) {
                drawing.anchors.forEach((anchor) => {
                    if (anchor.medium) {
                        const image = {
                            type: "image",
                            imageId: anchor.medium.index,
                            range: anchor.range,
                            hyperlinks: anchor.picture.hyperlinks,
                        };
                        model.media.push(image);
                    }
                });
            }
        }

        const backgroundRel = model.background && rels[model.background.rId];
        if (backgroundRel) {
            const target = backgroundRel.Target.split("/media/")[1];
            const imageId = getOwn(options.mediaIndex, target);
            if (imageId !== undefined) {
                model.media.push({
                    type: "background",
                    imageId,
                });
            }
        }

        model.tables = (model.tables || []).map((tablePart) => getPart(options.tables, rels[tablePart.rId])).filter(Boolean);

        delete model.relationships;
        delete model.legacyDrawing;
        delete model.hyperlinks;
        delete model.comments;
    }
}

WorkSheetXform.WORKSHEET_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "x14ac",
    "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
};

// budgets shared by all the worksheets of one load
WorkSheetXform.createLoadBudget = () => ({
    mergedCells: MAX_LOADED_MERGED_CELLS,
    dataValidationCells: DataValidationsXform.MAX_EXPANDED_CELLS,
});

module.exports = WorkSheetXform;
