import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { PivotTableModel } from "../../../types";
import { xmlEncode } from "../../../utils/utils";

import BaseXform from "../base-xform";

/**
 * The records of the pivot cache of a pivot table, `xl/pivotCache/pivotCacheRecords1.xml`: a copy of the source rows
 * (below the header row). The values of a row or column field are written as the index of the value in the shared
 * items of the field, the others as numbers or strings. Render only (parsing is not supported: the inherited no-op
 * methods).
 */
class PivotCacheRecordsXform extends BaseXform<PivotTableModel> {
    /** The namespaces of `<pivotCacheRecords>`. */
    static readonly PIVOT_CACHE_RECORDS_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "xr",
        "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    };

    declare map: Record<string, never>;

    constructor() {
        super();

        this.map = {};
    }

    get tag(): string {
        // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheRecords.html
        return "pivotCacheRecords";
    }

    /**
     * @throws Error when a value of a row or column field is not in the shared items of the field, or when a source row
     *   has a value beyond the header columns
     */
    override render(xmlStream: XmlStream, model: PivotTableModel): void {
        const { sourceSheet, cacheFields } = model;
        const sourceBodyRows = sourceSheet.getSheetValues().slice(2);

        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, {
            ...PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES,
            count: sourceBodyRows.length,
        });
        xmlStream.writeXml(renderTable());
        xmlStream.closeNode();

        // Helpers

        function renderTable(): string {
            const rowsInXML = sourceBodyRows.map((row) => {
                const realRow = row.slice(1);
                return [...renderRowLines(realRow)].join("");
            });
            return rowsInXML.join("");
        }

        function* renderRowLines(row: readonly unknown[]): Generator<string, void, undefined> {
            // PivotCache Record: http://www.datypic.com/sc/ooxml/e-ssml_r-1.html
            // Note: pretty-printing this for now to ease debugging.
            yield "\n  <r>";
            for (const [index, cellValue] of row.entries()) {
                yield "\n    ";
                // the cache fields are the columns of the header row
                const cacheField = cacheFields[index];
                if (!cacheField) {
                    throw new Error(`The source of the pivot table has a value in column ${index + 1}, which has no header.`);
                }
                yield renderCell(cellValue, cacheField.sharedItems);
            }
            yield "\n  </r>";
        }

        function renderCell(value: unknown, sharedItems: readonly unknown[] | null): string {
            // no shared items
            // --------------------------------------------------
            if (sharedItems === null) {
                if (typeof value === "number" && Number.isFinite(value)) {
                    // Numeric value: http://www.datypic.com/sc/ooxml/e-ssml_n-2.html
                    return `<n v="${value}" />`;
                }
                // Character Value: http://www.datypic.com/sc/ooxml/e-ssml_s-2.html
                return `<s v="${xmlEncode(String(value))}" />`;
            }

            // shared items
            // --------------------------------------------------
            const sharedItemsIndex = sharedItems.indexOf(value);
            if (sharedItemsIndex < 0) {
                throw new Error(`${JSON.stringify(value)} not in sharedItems ${JSON.stringify(sharedItems)}`);
            }
            // Shared Items Index: http://www.datypic.com/sc/ooxml/e-ssml_x-9.html
            return `<x v="${sharedItemsIndex}" />`;
        }
    }
}

export default PivotCacheRecordsXform;
