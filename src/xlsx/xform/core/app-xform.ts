import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { AppModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import StringXform from "../simple/string-xform";

import AppHeadingPairsXform from "./app-heading-pairs-xform";
import AppTitleOfPartsXform from "./app-titles-of-parts-xform";

/**
 * `docProps/app.xml`, the extended properties: the application, the worksheet names (rendered only), the company and
 * the manager (rendered and parsed). The other properties are written with fixed values and not parsed.
 */
class AppXform extends BaseXform<AppModel> {
    /** Formats a date as W3CDTF (ISO 8601 without the fraction of seconds). */
    static DateFormat(this: void, dt: Date): string {
        return dt.toISOString().replace(/[.]\d{3,6}/, "");
    }

    /** The attributes of a W3CDTF date element. */
    static readonly DateAttrs: XmlAttributes = { "xsi:type": "dcterms:W3CDTF" };

    /** The namespaces of `<Properties>`. */
    static readonly PROPERTY_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    };

    // the key "TitleOfParts" does not match the element (TitlesOfParts): its content is skipped as an unknown element
    declare map: {
        readonly Company: StringXform;
        readonly Manager: StringXform;
        readonly HeadingPairs: AppHeadingPairsXform;
        readonly TitleOfParts: AppTitleOfPartsXform;
    };

    constructor() {
        super();

        this.map = {
            Company: new StringXform({ tag: "Company" }),
            Manager: new StringXform({ tag: "Manager" }),
            HeadingPairs: new AppHeadingPairsXform(),
            TitleOfParts: new AppTitleOfPartsXform(),
        };
    }

    /** @param model - `worksheets` is required to render */
    override render(xmlStream: XmlStream, model: AppModel & Required<Pick<AppModel, "worksheets">>): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("Properties", AppXform.PROPERTY_ATTRIBUTES);

        xmlStream.leafNode("Application", undefined, "Microsoft Excel");
        xmlStream.leafNode("DocSecurity", undefined, "0");
        xmlStream.leafNode("ScaleCrop", undefined, "false");

        this.map.HeadingPairs.render(xmlStream, model.worksheets);
        this.map.TitleOfParts.render(xmlStream, model.worksheets);
        this.map.Company.render(xmlStream, model.company || "");
        this.map.Manager.render(xmlStream, model.manager);

        xmlStream.leafNode("LinksUpToDate", undefined, "false");
        xmlStream.leafNode("SharedDoc", undefined, "false");
        xmlStream.leafNode("HyperlinksChanged", undefined, "false");
        xmlStream.leafNode("AppVersion", undefined, "16.0300");

        xmlStream.closeNode();
    }

    /** @returns `true` for `<Properties>` and the properties that are parsed, `false` for the others (skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "Properties":
                return true;
            default:
                this.parser = ownEntry(this.map, node.name);
                if (this.parser) {
                    this.parser.parseOpen(node);
                    return true;
                }

                // there's a lot we don't bother to parse
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "Properties":
                // (`?? undefined`: a child's model is null only after a reset, which this xform never does)
                this.model = {
                    // never set: the titles of parts are not parsed
                    worksheets: this.map.TitleOfParts.model ?? undefined,
                    company: this.map.Company.model ?? undefined,
                    manager: this.map.Manager.model ?? undefined,
                };
                return false;
            default:
                return true;
        }
    }
}

export default AppXform;
