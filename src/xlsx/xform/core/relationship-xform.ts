import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RelationshipModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * A `<Relationship>` of a relationships part: the model is the attributes of the element (`Id`, `Type`, `Target`,
 * `TargetMode`), rendered and parsed as is (parsing keeps every attribute, as strings from the file).
 */
class RelationshipXform extends BaseXform<RelationshipModel> {
    override render(xmlStream: XmlStream, model: RelationshipModel): void {
        xmlStream.leafNode("Relationship", model);
    }

    /** @returns `true` for `<Relationship>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "Relationship":
                // the attributes object of the parser (a missing Id / Type / Target is undefined)
                this.model = node.attributes as RelationshipModel;
                return true;
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default RelationshipXform;
