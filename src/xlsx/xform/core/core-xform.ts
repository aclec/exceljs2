import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { CoreModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import DateXform from "../simple/date-xform";
import StringXform from "../simple/string-xform";
import IntegerXform from "../simple/integer-xform";

// an Invalid Date (set by the app) is not written
function validDate(date: Date | undefined): Date | undefined {
    return date instanceof Date && Number.isNaN(date.getTime()) ? undefined : date;
}

/**
 * `docProps/core.xml`, the core properties (Dublin Core terms). Other Dublin Core terms (dc:contributor,
 * dcterms:available...) are skipped when parsing. A date that is not valid is neither written nor read.
 */
class CoreXform extends BaseXform<CoreModel> {
    /** Formats a date as W3CDTF (ISO 8601 without milliseconds). */
    static DateFormat(this: void, dt: Date): string {
        return dt.toISOString().replace(/[.]\d{3}/, "");
    }

    /**
     * Parses a W3CDTF date. A date that is not valid is read as undefined: it used to load as an Invalid Date, on which
     * every later write failed (toISOString throws).
     */
    static DateParse(this: void, str: string | undefined): Date | undefined {
        return validDate(new Date(str ?? NaN));
    }

    /** The attributes of a W3CDTF date element. */
    static readonly DateAttrs: XmlAttributes = { "xsi:type": "dcterms:W3CDTF" };

    /** The namespaces of `<cp:coreProperties>`. */
    static readonly CORE_PROPERTY_ATTRIBUTES: XmlAttributes = {
        "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "xmlns:dc": "http://purl.org/dc/elements/1.1/",
        "xmlns:dcterms": "http://purl.org/dc/terms/",
        "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    };

    declare map: {
        readonly "dc:creator": StringXform;
        readonly "dc:title": StringXform;
        readonly "dc:subject": StringXform;
        readonly "dc:description": StringXform;
        readonly "dc:identifier": StringXform;
        readonly "dc:language": StringXform;
        readonly "cp:keywords": StringXform;
        readonly "cp:category": StringXform;
        readonly "cp:lastModifiedBy": StringXform;
        readonly "cp:lastPrinted": DateXform;
        readonly "cp:revision": IntegerXform;
        readonly "cp:version": StringXform;
        readonly "cp:contentStatus": StringXform;
        readonly "cp:contentType": StringXform;
        readonly "dcterms:created": DateXform;
        readonly "dcterms:modified": DateXform;
    };

    /** The depth inside an element that is skipped (not modelled); 0 outside such an element. */
    private skipDepth = 0;

    constructor() {
        super();

        this.map = {
            "dc:creator": new StringXform({ tag: "dc:creator" }),
            "dc:title": new StringXform({ tag: "dc:title" }),
            "dc:subject": new StringXform({ tag: "dc:subject" }),
            "dc:description": new StringXform({ tag: "dc:description" }),
            "dc:identifier": new StringXform({ tag: "dc:identifier" }),
            "dc:language": new StringXform({ tag: "dc:language" }),
            "cp:keywords": new StringXform({ tag: "cp:keywords" }),
            "cp:category": new StringXform({ tag: "cp:category" }),
            "cp:lastModifiedBy": new StringXform({ tag: "cp:lastModifiedBy" }),
            "cp:lastPrinted": new DateXform({ tag: "cp:lastPrinted", format: CoreXform.DateFormat, parse: CoreXform.DateParse }),
            "cp:revision": new IntegerXform({ tag: "cp:revision" }),
            "cp:version": new StringXform({ tag: "cp:version" }),
            "cp:contentStatus": new StringXform({ tag: "cp:contentStatus" }),
            "cp:contentType": new StringXform({ tag: "cp:contentType" }),
            "dcterms:created": new DateXform({
                tag: "dcterms:created",
                attrs: CoreXform.DateAttrs,
                format: CoreXform.DateFormat,
                parse: CoreXform.DateParse,
            }),
            "dcterms:modified": new DateXform({
                tag: "dcterms:modified",
                attrs: CoreXform.DateAttrs,
                format: CoreXform.DateFormat,
                parse: CoreXform.DateParse,
            }),
        };
    }

    override render(xmlStream: XmlStream, model: CoreModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("cp:coreProperties", CoreXform.CORE_PROPERTY_ATTRIBUTES);

        const { map } = this;
        map["dc:creator"].render(xmlStream, model.creator);
        map["dc:title"].render(xmlStream, model.title);
        map["dc:subject"].render(xmlStream, model.subject);
        map["dc:description"].render(xmlStream, model.description);
        map["dc:identifier"].render(xmlStream, model.identifier);
        map["dc:language"].render(xmlStream, model.language);
        map["cp:keywords"].render(xmlStream, model.keywords);
        map["cp:category"].render(xmlStream, model.category);
        map["cp:lastModifiedBy"].render(xmlStream, model.lastModifiedBy);
        map["cp:lastPrinted"].render(xmlStream, validDate(model.lastPrinted));
        map["cp:revision"].render(xmlStream, model.revision);
        map["cp:version"].render(xmlStream, model.version);
        map["cp:contentStatus"].render(xmlStream, model.contentStatus);
        map["cp:contentType"].render(xmlStream, model.contentType);
        map["dcterms:created"].render(xmlStream, validDate(model.created));
        map["dcterms:modified"].render(xmlStream, validDate(model.modified));

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the part is accepted (unknown ones are skipped with their content) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth++;
            return true;
        }
        switch (node.name) {
            case "cp:coreProperties":
            case "coreProperties":
                this.skipDepth = 0;
                return true;
            default:
                this.parser = ownEntry(this.map, node.name);
                if (this.parser) {
                    this.parser.parseOpen(node);
                    return true;
                }
                // other Dublin Core terms (dc:contributor, dcterms:available...) are not modelled
                this.skipDepth = 1;
                return true;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @throws Error on an end tag that matches no start tag of the part */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth--;
            return true;
        }
        switch (name) {
            case "cp:coreProperties":
            case "coreProperties": {
                // (`?? undefined`: a child's model is null only after a reset, which this xform never does)
                const { map } = this;
                this.model = {
                    creator: map["dc:creator"].model ?? undefined,
                    title: map["dc:title"].model ?? undefined,
                    subject: map["dc:subject"].model ?? undefined,
                    description: map["dc:description"].model ?? undefined,
                    identifier: map["dc:identifier"].model ?? undefined,
                    language: map["dc:language"].model ?? undefined,
                    keywords: map["cp:keywords"].model ?? undefined,
                    category: map["cp:category"].model ?? undefined,
                    lastModifiedBy: map["cp:lastModifiedBy"].model ?? undefined,
                    lastPrinted: map["cp:lastPrinted"].model ?? undefined,
                    revision: map["cp:revision"].model ?? undefined,
                    contentStatus: map["cp:contentStatus"].model ?? undefined,
                    contentType: map["cp:contentType"].model ?? undefined,
                    created: map["dcterms:created"].model ?? undefined,
                    modified: map["dcterms:modified"].model ?? undefined,
                };
                return false;
            }
            default:
                throw new Error(`Unexpected xml node in parseClose: ${name}`);
        }
    }
}

export default CoreXform;
