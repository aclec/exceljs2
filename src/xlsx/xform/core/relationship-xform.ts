import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RelationshipModel } from "../../../types";
import BaseXform from "../base-xform";

// the attributes of a <Relationship> (in their order), or null without one of the required ones
function toRelationship(attributes: SaxTag["attributes"]): RelationshipModel | null {
    const { Id, Type, Target, TargetMode } = attributes;
    if (Id === undefined || Type === undefined || Target === undefined) {
        return null;
    }
    const model: RelationshipModel = { ...attributes, Id, Type, Target };
    if (TargetMode !== undefined && TargetMode !== "External" && TargetMode !== "Internal") {
        delete model.TargetMode;
    }
    return model;
}

/**
 * A `<Relationship>` of a relationships part: the model is the attributes of the element (`Id`, `Type`, `Target`,
 * `TargetMode`), rendered and parsed as is (parsing keeps every attribute, as strings from the file). A relationship
 * without `Id`, `Type` or `Target` (all required, ECMA-376 Part 2 CT_Relationship) parses as `null`; an invalid
 * `TargetMode` (neither `External` nor `Internal`) is dropped (the default: `Internal`).
 */
class RelationshipXform extends BaseXform<RelationshipModel> {
    override render(xmlStream: XmlStream, model: RelationshipModel): void {
        xmlStream.leafNode("Relationship", model);
    }

    /** @returns `true` for `<Relationship>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "Relationship":
                this.model = toRelationship(node.attributes);
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
