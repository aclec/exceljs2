import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { ContentTypesModel } from "../../../types";

import BaseXform from "../base-xform";

/**
 * `[Content_Types].xml`: the content type of each part of the package, from the write model. Render only: parsing
 * accepts nothing.
 */
class ContentTypesXform extends BaseXform<ContentTypesModel> {
    /** The namespace of `<Types>`. */
    static readonly PROPERTY_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
    };

    override render(xmlStream: XmlStream, model: ContentTypesModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("Types", ContentTypesXform.PROPERTY_ATTRIBUTES);

        // one Default per image extension (a Set: the extensions are the application's data)
        const imageTypes = new Set<string>();
        for (const medium of model.media || []) {
            if (medium.type === "image") {
                const imageType = medium.extension;
                if (!imageTypes.has(imageType)) {
                    imageTypes.add(imageType);
                    xmlStream.leafNode("Default", { Extension: imageType, ContentType: `image/${imageType}` });
                }
            }
        }

        xmlStream.leafNode("Default", {
            Extension: "rels",
            ContentType: "application/vnd.openxmlformats-package.relationships+xml",
        });
        xmlStream.leafNode("Default", { Extension: "xml", ContentType: "application/xml" });

        xmlStream.leafNode("Override", {
            PartName: "/xl/workbook.xml",
            ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        });

        for (const worksheet of model.worksheets) {
            const name = `/xl/worksheets/sheet${worksheet.id}.xml`;
            xmlStream.leafNode("Override", {
                PartName: name,
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            });
        }

        if ((model.pivotTables || []).length) {
            // Note(2023-10-06): assuming at most one pivot table for now.
            xmlStream.leafNode("Override", {
                PartName: "/xl/pivotCache/pivotCacheDefinition1.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
            });
            xmlStream.leafNode("Override", {
                PartName: "/xl/pivotCache/pivotCacheRecords1.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
            });
            xmlStream.leafNode("Override", {
                PartName: "/xl/pivotTables/pivotTable1.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
            });
        }

        xmlStream.leafNode("Override", {
            PartName: "/xl/theme/theme1.xml",
            ContentType: "application/vnd.openxmlformats-officedocument.theme+xml",
        });
        xmlStream.leafNode("Override", {
            PartName: "/xl/styles.xml",
            ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        });

        const hasSharedStrings = model.sharedStrings && model.sharedStrings.count;
        if (hasSharedStrings) {
            xmlStream.leafNode("Override", {
                PartName: "/xl/sharedStrings.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            });
        }

        if (model.tables) {
            for (const table of model.tables) {
                xmlStream.leafNode("Override", {
                    PartName: `/xl/tables/${table.target}`,
                    ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                });
            }
        }

        if (model.drawings) {
            for (const drawing of model.drawings) {
                xmlStream.leafNode("Override", {
                    PartName: `/xl/drawings/${drawing.name}.xml`,
                    ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
                });
            }
        }

        if (model.commentRefs) {
            xmlStream.leafNode("Default", {
                Extension: "vml",
                ContentType: "application/vnd.openxmlformats-officedocument.vmlDrawing",
            });

            for (const { commentName } of model.commentRefs) {
                xmlStream.leafNode("Override", {
                    PartName: `/xl/${commentName}.xml`,
                    ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                });
            }
        }

        xmlStream.leafNode("Override", {
            PartName: "/docProps/core.xml",
            ContentType: "application/vnd.openxmlformats-package.core-properties+xml",
        });
        xmlStream.leafNode("Override", {
            PartName: "/docProps/app.xml",
            ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        });

        xmlStream.closeNode();
    }

    /** @returns `false`: this part is not parsed */
    override parseOpen(): boolean {
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default ContentTypesXform;
