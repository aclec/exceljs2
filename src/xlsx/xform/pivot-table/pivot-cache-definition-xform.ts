import type { XmlAttributes } from "../../../utils/xml-stream";
import XmlStream from "../../../utils/xml-stream";
import type { PivotTableModel } from "../../../types";
import BaseXform from "../base-xform";
import CacheField from "./cache-field";

/**
 * The pivot cache definition of a pivot table, `xl/pivotCache/pivotCacheDefinition1.xml`: its source range and its cache
 * fields. Render only (parsing is not supported: the inherited no-op methods).
 */
class PivotCacheDefinitionXform extends BaseXform<PivotTableModel> {
    /** The namespaces of `<pivotCacheDefinition>`. */
    static readonly PIVOT_CACHE_DEFINITION_ATTRIBUTES: XmlAttributes = {
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
        // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheDefinition.html
        return "pivotCacheDefinition";
    }

    override render(xmlStream: XmlStream, model: PivotTableModel): void {
        const { sourceSheet, cacheFields } = model;

        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, {
            ...PivotCacheDefinitionXform.PIVOT_CACHE_DEFINITION_ATTRIBUTES,
            "r:id": "rId1",
            refreshOnLoad: "1", // important for our implementation to work
            refreshedBy: "Author",
            refreshedDate: "45125.026046874998",
            createdVersion: "8",
            refreshedVersion: "8",
            minRefreshableVersion: "3",
            recordCount: cacheFields.length + 1,
        });

        xmlStream.openNode("cacheSource", { type: "worksheet" });
        xmlStream.leafNode("worksheetSource", {
            ref: sourceSheet.dimensions.shortRange,
            sheet: sourceSheet.name,
        });
        xmlStream.closeNode();

        xmlStream.openNode("cacheFields", { count: cacheFields.length });
        // Note: keeping this pretty-printed for now to ease debugging.
        // (the values of the fields are encoded by CacheField.render)
        xmlStream.writeXml(cacheFields.map((cacheField) => new CacheField(cacheField).render()).join("\n    "));
        xmlStream.closeNode();

        xmlStream.closeNode();
    }
}

export default PivotCacheDefinitionXform;
