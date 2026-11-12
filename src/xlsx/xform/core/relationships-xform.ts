import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { RelationshipModel } from "../../../types";
import BaseXform from "../base-xform";

import RelationshipXform from "./relationship-xform";

/** A relationships part (`_rels/*.rels`): the list of its `<Relationship>` elements. */
class RelationshipsXform extends BaseXform<RelationshipModel[]> {
    /** The namespace of `<Relationships>`. */
    static readonly RELATIONSHIPS_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
    };

    declare map: { readonly Relationship: RelationshipXform };

    constructor() {
        super();

        this.map = {
            Relationship: new RelationshipXform(),
        };
    }

    override render(xmlStream: XmlStream, model: readonly RelationshipModel[]): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("Relationships", RelationshipsXform.RELATIONSHIPS_ATTRIBUTES);

        const relationshipXform = this.map.Relationship;
        for (const relationship of model) {
            relationshipXform.render(xmlStream, relationship);
        }

        xmlStream.closeNode();
    }

    /**
     * @returns `true` for `<Relationships>` and its relationships
     * @throws Error on any other element
     */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "Relationships":
                this.model = [];
                return true;
            case "Relationship":
                this.parser = this.map.Relationship;
                this.parser.parseOpen(node);
                return true;
            default:
                throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @throws Error on an unexpected end tag */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // <Relationships> opened the list before its items; RelationshipXform has its model once open
                this.model!.push(this.map.Relationship.model!);
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "Relationships":
                return false;
            default:
                throw new Error(`Unexpected xml node in parseClose: ${name}`);
        }
    }
}

export default RelationshipsXform;
